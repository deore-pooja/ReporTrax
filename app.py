"""
Report Builder - end to end Flask web app
-----------------------------------------
* Sign up / Sign in (name, email, password) - SQLite + hashed passwords
* Dashboard with a "Report" card
* Report page: your fixed Excel format, every sheet available from a dropdown on the
  left, the selected sheet shown on the right, every edit auto-saved per user.

The fixed Excel format lives in  template/report_template.xlsx  - replace that file
with your own workbook and the app picks up its sheets automatically.
"""
import colorsys
import io
import json
import os
import re
import shutil
import secrets
import sqlite3
import threading
from functools import wraps

from flask import (Flask, g, jsonify, redirect, render_template, request,
                   send_file, session, url_for, flash, abort)
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles.colors import COLOR_INDEX
from openpyxl.utils import column_index_from_string, get_column_letter
from PIL import Image, ImageOps
from pdf_export import build_pdf, header_ranges, page_blocks, image_slots, dimension_icons, icon_svg, area_bars, photo_rects, AREA_OPTIONS
from sanitize import strip_drawings
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
USER_FILES = os.path.join(DATA_DIR, "user_files")
DB_PATH = os.path.join(DATA_DIR, "app.db")
TEMPLATE_PATH = os.path.join(BASE_DIR, "template", "report_template.xlsx")
SECRET_PATH = os.path.join(DATA_DIR, "secret.key")

MAX_ROWS = 1000   # safety caps for what is drawn in the browser
MAX_COLS = 2300   # QUALITY CHECK has 200 cards (last column 2199)

os.makedirs(USER_FILES, exist_ok=True)
os.makedirs(os.path.dirname(TEMPLATE_PATH), exist_ok=True)

app = Flask(__name__)
if not os.path.exists(SECRET_PATH):
    with open(SECRET_PATH, "w") as f:
        f.write(secrets.token_hex(32))
with open(SECRET_PATH) as f:
    app.secret_key = f.read().strip()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  MAX_CONTENT_LENGTH=25 * 1024 * 1024)   # photos from a phone camera

_locks = {}
_locks_guard = threading.Lock()


def user_lock(uid):
    with _locks_guard:
        return _locks.setdefault(uid, threading.Lock())


# --------------------------------------------------------------------------- DB
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db:
        db.close()


def init_db():
    with sqlite3.connect(DB_PATH) as db:
        db.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)""")


def ensure_template():
    """If no template was supplied yet, create a small sample so the app still runs."""
    if os.path.exists(TEMPLATE_PATH):
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Sample report - replace template/report_template.xlsx with your own file"
    for i, h in enumerate(["Item", "Owner", "Status", "Notes"], start=1):
        ws.cell(row=3, column=i, value=h)
    for name in ("Details", "Metrics", "Remarks"):
        s = wb.create_sheet(name)
        s["A1"] = name
    wb.save(TEMPLATE_PATH)


# ------------------------------------------------------------------------- auth
def login_required(view):
    @wraps(view)
    def wrapped(*a, **kw):
        if not session.get("uid"):
            return redirect(url_for("signin"))
        return view(*a, **kw)
    return wrapped


@app.route("/")
def index():
    return redirect(url_for("dashboard" if session.get("uid") else "signin"))


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@app.route("/signup", methods=["GET", "POST"])
def signup():
    if request.method == "POST":
        name = request.form.get("name", "").strip()
        email = request.form.get("email", "").strip().lower()
        pw = request.form.get("password", "")
        error = None
        if not name:
            error = "Enter your name."
        elif not EMAIL_RE.match(email):
            error = "Enter a valid email address."
        elif len(pw) < 6:
            error = "Password must be at least 6 characters."
        if not error:
            db = get_db()
            try:
                cur = db.execute(
                    "INSERT INTO users (name, email, password_hash) VALUES (?,?,?)",
                    (name, email, generate_password_hash(pw)))
                db.commit()
            except sqlite3.IntegrityError:
                error = "An account with this email already exists. Sign in instead."
            else:
                session.clear()
                session["uid"] = cur.lastrowid
                session["name"] = name
                return redirect(url_for("dashboard"))
        flash(error, "error")
        return render_template("signup.html", name=name, email=email)
    return render_template("signup.html")


@app.route("/signin", methods=["GET", "POST"])
def signin():
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        pw = request.form.get("password", "")
        row = get_db().execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if row and check_password_hash(row["password_hash"], pw):
            session.clear()
            session["uid"] = row["id"]
            session["name"] = row["name"]
            return redirect(url_for("dashboard"))
        flash("Email or password is incorrect.", "error")
        return render_template("signin.html", email=email)
    return render_template("signin.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("signin"))


@app.route("/dashboard")
@login_required
def dashboard():
    return render_template("dashboard.html", name=session.get("name"))


# ----------------------------------------------------------------------- report
def user_file(uid):
    """Each user edits a private copy of the fixed template; created on first use.
    Pictures/shapes are stripped from the copy (they are not used in this app)."""
    path = os.path.join(USER_FILES, f"user_{uid}.xlsx")
    if not os.path.exists(path):
        ensure_template()
        strip_drawings(TEMPLATE_PATH, path)
    return path


_wb_cache = {}   # uid -> (mtime, Workbook); avoids re-parsing the big workbook on every click


def get_wb(uid):
    """Call while holding user_lock(uid)."""
    path = user_file(uid)
    mtime = os.path.getmtime(path)
    hit = _wb_cache.get(uid)
    if hit and hit[0] == mtime:
        return hit[1]
    wb = load_workbook(path)
    _wb_cache[uid] = (mtime, wb)
    return wb


def sheet_has_data(ws):
    """True if the sheet holds at least one non-empty cell (images do not count)."""
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if v is not None and str(v).strip() != "":
                return True
    return False


def visible_sheets(wb):
    """Only visible sheets that actually contain data - empty pages are ignored."""
    return [ws for ws in wb.worksheets
            if ws.sheet_state == "visible" and sheet_has_data(ws)]


# ---- colours -----------------------------------------------------------------
THEME_ORDER = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3",
               "accent4", "accent5", "accent6", "hlink", "folHlink"]
DEFAULT_THEME = ["FFFFFF", "000000", "E7E6E6", "44546A", "4472C4", "ED7D31",
                 "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"]


def theme_colors(wb):
    cols = list(DEFAULT_THEME)
    try:
        xml = wb.loaded_theme.decode("utf-8")
        for i, key in enumerate(THEME_ORDER):
            m = re.search(r"<a:%s>.*?(?:srgbClr val=\"(\w{6})\"|lastClr=\"(\w{6})\")" % key, xml, re.S)
            if m:
                cols[i] = m.group(1) or m.group(2)
    except Exception:
        pass
    return cols


def apply_tint(hex6, tint):
    if not tint:
        return hex6
    r, g, b = (int(hex6[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 + tint) if tint < 0 else l * (1 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


def color_of(c, theme):
    """Return '#rrggbb' or None."""
    try:
        if c is None:
            return None
        if c.type == "rgb" and isinstance(c.rgb, str) and len(c.rgb) in (6, 8):
            return "#" + c.rgb[-6:]
        if c.type == "theme" and 0 <= c.theme < len(theme):
            return "#" + apply_tint(theme[c.theme], c.tint)
        if c.type == "indexed" and 0 <= c.indexed < len(COLOR_INDEX):
            return "#" + COLOR_INDEX[c.indexed][-6:]
    except Exception:
        pass
    return None


BORDER_PX = {"thin": 1, "hair": 1, "dotted": 1, "dashed": 1, "medium": 2,
             "mediumDashed": 2, "thick": 3, "double": 3}


def border_css(side, name, theme):
    if side and side.style:
        col = color_of(side.color, theme) or "#333333"
        return f"border-{name}:{BORDER_PX.get(side.style, 1)}px solid {col}"
    return None


def cell_style(cell, theme, box=None):
    """CSS for a cell. `box` = (top_cell, bottom_cell, left_cell, right_cell) for merged areas."""
    css = []
    f = cell.font
    if f:
        if f.b:
            css.append("font-weight:700")
        if f.i:
            css.append("font-style:italic")
        if f.u:
            css.append("text-decoration:underline")
        if f.sz:
            css.append(f"font-size:{round(f.sz * 1.333)}px")
        col = color_of(f.color, theme)
        if col and col.lower() != "#000000":
            css.append(f"color:{col}")
    if cell.fill and cell.fill.fill_type == "solid":
        col = color_of(cell.fill.fgColor, theme)
        if col and col.lower() != "#ffffff":
            css.append(f"background:{col}")
    al = cell.alignment
    if al and al.horizontal in ("left", "center", "right"):
        css.append(f"text-align:{al.horizontal}")
    elif al and al.horizontal == "centerContinuous":
        css.append("text-align:center")
    t, b, l, r = box or (cell, cell, cell, cell)
    for part in (border_css(t.border.top, "top", theme), border_css(b.border.bottom, "bottom", theme),
                 border_css(l.border.left, "left", theme), border_css(r.border.right, "right", theme)):
        if part:
            css.append(part)
    return ";".join(css)


def fmt_number(v, fmt):
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        return "" if v is None else str(v)
    m = re.search(r"0\.(0+)", fmt or "")
    if m:
        return f"{v:.{len(m.group(1))}f}"
    if isinstance(v, float):
        return ("%.10g" % v)
    return str(v)


# ---- tiny, safe formula engine (+ - * / ^, SUM, AVERAGE, MIN, MAX, ROUND, ABS, COUNT, IF) ------
REF_RE = re.compile(r"(?<![A-Za-z0-9_])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(])")
ALLOWED_NAMES = {"SUM", "AVERAGE", "MIN", "MAX", "ROUND", "ABS", "COUNT", "IF", "C", "R"}


def _num(v):
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return 0


def eval_cell(ws, ref, depth=0, cache=None):
    cache = {} if cache is None else cache
    if ref in cache:
        return cache[ref]
    v = ws[ref].value
    if isinstance(v, str) and v.startswith("="):
        if depth > 20:
            return "#CYCLE"
        v = eval_formula(ws, v, depth + 1, cache)
    cache[ref] = v
    return v


def eval_formula(ws, formula, depth=0, cache=None):
    expr = formula[1:].strip()
    bare = re.sub(r'"[^"]*"', "", expr)
    if "!" in bare or "[" in bare or not re.fullmatch(r"[A-Za-z0-9_$:.,()+\-*/^ <>=%]*", bare):
        return "#UNSUPPORTED"
    names = set(re.findall(r"[A-Za-z_]+", REF_RE.sub("", bare)))
    if any(n.upper() not in ALLOWED_NAMES for n in names):
        return "#UNSUPPORTED"
    def _ref(m):
        if m.group(3):
            return f'R("{m.group(1)}{m.group(2)}","{m.group(3)}{m.group(4)}")'
        return f'C("{m.group(1)}{m.group(2)}")'
    py = REF_RE.sub(_ref, expr)
    py = py.replace("^", "**").replace("<>", "!=")
    py = re.sub(r"(?<![<>!=])=(?!=)", "==", py)

    def C(ref):
        return _num(eval_cell(ws, ref, depth, cache))

    def R(a, b):
        out = []
        for row in ws[a:b]:
            for cell in row:
                v = eval_cell(ws, cell.coordinate, depth, cache)
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(v)
                elif isinstance(v, str):
                    try:
                        out.append(float(v))
                    except ValueError:
                        pass
        return out

    def flat(args):
        for a in args:
            if isinstance(a, list):
                yield from a
            else:
                yield a

    env = {"C": C, "R": R,
           "SUM": lambda *a: sum(flat(a)),
           "AVERAGE": lambda *a: (sum(flat(a)) / len(list(flat(a)))) if list(flat(a)) else 0,
           "MIN": lambda *a: min(flat(a), default=0), "MAX": lambda *a: max(flat(a), default=0),
           "COUNT": lambda *a: len(list(flat(a))), "ABS": abs,
           "ROUND": lambda x, n=0: round(x, int(n)), "IF": lambda c, a, b=0: a if c else b}
    py = re.sub(r"\b(sum|average|min|max|round|abs|count|if)\(", lambda m: m.group(1).upper() + "(", py)
    try:
        return eval(py, {"__builtins__": {}}, env)
    except ZeroDivisionError:
        return "#DIV/0!"
    except Exception:
        return "#ERROR"


def calc_values(ws):
    """All formula cells of a sheet -> {(r, c): shown_text}"""
    out, cache = {}, {}
    for row in ws.iter_rows():
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                res = eval_cell(ws, cell.coordinate, 0, cache)
                out[(cell.row, cell.column)] = fmt_number(res, cell.number_format) if not isinstance(res, str) else res
    return out


# ---- sheet -> JSON ---------------------------------------------------------------
def used_columns(ws):
    cols = set()
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and str(c.value).strip() != "":
                cols.add(c.column)
    for rng in ws.merged_cells.ranges:
        cols.update(range(rng.min_col, rng.max_col + 1))
    return sorted(cols)


def used_rows_limit(ws):
    last = 1
    for row in ws.iter_rows():
        for c in row:
            if c.value is not None and str(c.value).strip() != "":
                last = max(last, c.row)
    for rng in ws.merged_cells.ranges:
        last = max(last, rng.max_row)
    return min(max(last, 1), ws.max_row, MAX_ROWS)


def find_bands(ws):
    """Normal sheets = one band (all columns). Very wide sheets (cards laid out side by side)
    are split into one band per block of adjacent used columns."""
    cols = used_columns(ws) or [1]
    if cols[-1] - cols[0] <= 40:
        return [(1, min(max(cols[-1] + 1, 2), ws.max_column, MAX_COLS))]
    bands, start, prev = [], cols[0], cols[0]
    for c in cols[1:]:
        if c - prev > 1:
            bands.append((max(1, start - 1), prev + 1))
            start = c
        prev = c
    bands.append((max(1, start - 1), prev + 1))
    return bands


def column_widths(ws, c0, c1):
    default = ws.sheet_format.defaultColWidth or 8.43
    widths, hidden = {}, set()
    for dim in ws.column_dimensions.values():
        lo = dim.min or column_index_from_string(dim.index)
        hi = dim.max or lo
        for c in range(max(lo, c0), min(hi, c1) + 1):
            widths[c] = dim.width if dim.width else default
            if dim.hidden:
                hidden.add(c)
    return [0 if c in hidden else int(widths.get(c, default) * 7 + 5) for c in range(c0, c1 + 1)]


def sheet_payload(wb, ws, band=0, uid=None):
    theme = theme_colors(wb)
    bands = find_bands(ws)
    band = max(0, min(band, len(bands) - 1))
    c0, c1 = bands[band]
    max_r = used_rows_limit(ws)
    blocks = page_blocks(ws)                                  # show every page down to its blue bottom bar
    if blocks:
        max_r = min(max(max_r, max(b[1] for b in blocks)), ws.max_row, MAX_ROWS)

    spans, hidden, boxes = {}, set(), {}
    for rng in ws.merged_cells.ranges:
        spans[(rng.min_row, rng.min_col)] = (rng.max_row - rng.min_row + 1, rng.max_col - rng.min_col + 1)
        boxes[(rng.min_row, rng.min_col)] = (
            ws.cell(rng.min_row, rng.min_col), ws.cell(rng.max_row, rng.min_col),
            ws.cell(rng.min_row, rng.min_col), ws.cell(rng.min_row, rng.max_col))
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    hidden.add((r, c))

    calc = calc_values(ws)
    slots = image_slots(ws)                                   # photo areas (Take Photo / Upload)
    heads = {(m.min_row, m.min_col) for m in header_ranges(ws)}   # page header (title + logo)
    icons = dimension_icons(ws)                               # door / window / wall-tile symbols
    bars = area_bars(ws)                                      # orange bar = drop-down list of areas
    lays = load_layout(uid) if uid else {}
    default_h = (ws.sheet_format.defaultRowHeight or 15) * 1.333
    rows = []
    for r in range(1, max_r + 1):
        rdim = ws.row_dimensions.get(r)
        if rdim and rdim.hidden:
            continue
        cells = []
        for c in range(c0, c1 + 1):
            if (r, c) in hidden:
                continue
            cell = ws.cell(row=r, column=c)
            raw = cell.value
            is_formula = isinstance(raw, str) and raw.startswith("=")
            rs, cs = spans.get((r, c), (1, 1))
            if is_formula:
                shown = calc.get((r, c), "")
            elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
                shown = fmt_number(raw, cell.number_format) if "0." in (cell.number_format or "") else str(raw)
            else:
                shown = "" if raw is None else str(raw)
            multi = rs > 1 or cs > 2 or "\n" in shown or len(shown) > 40 or bool(cell.alignment and cell.alignment.wrap_text)
            cells.append({"c": c, "v": shown, "f": raw if is_formula else None, "rs": rs, "cs": cs,
                          "m": 1 if multi else 0,
                          "va": (cell.alignment.vertical if cell.alignment and cell.alignment.vertical else "bottom"),
                          "s": cell_style(cell, theme, boxes.get((r, c)))})
            if (r, c) in heads:
                cells[-1]["hdr"] = 1
            if (r, c) in icons:
                cells[-1]["ico"] = icon_svg(icons[(r, c)])
            if (r, c) in bars:
                cells[-1]["dd"] = AREA_OPTIONS
            if (r, c) in slots:
                vers = []
                for i in (0, 1):
                    p = img_path(uid, ws.title, r, c, i) if uid else None
                    vers.append(int(os.path.getmtime(p) * 1000) if p and os.path.exists(p) else 0)
                cells[-1]["img"] = {"mode": slots[(r, c)], "v": vers,
                                    "lay": norm_layout(lays.get(layout_key(ws.title, r, c)))}
        rows.append({"r": r, "h": int((rdim.height * 1.333) if rdim and rdim.height else default_h), "cells": cells})
    return {"cols": column_widths(ws, c0, c1), "rows": rows, "band": band, "bands": len(bands),
            "gridlines": bool(ws.sheet_view.showGridLines is not False)}


@app.route("/report")
@login_required
def report():
    return render_template("report.html", name=session.get("name"))


@app.route("/api/sheets")
@login_required
def api_sheets():
    uid = session["uid"]
    with user_lock(uid):
        wb = get_wb(uid)
        names = [ws.title for ws in visible_sheets(wb)]
    return jsonify(sheets=names)


@app.route("/api/sheet")
@login_required
def api_sheet():
    uid = session["uid"]
    name = request.args.get("name", "")
    try:
        band = int(request.args.get("band", 0))
    except ValueError:
        band = 0
    with user_lock(uid):
        wb = get_wb(uid)
        if name not in wb.sheetnames or wb[name].sheet_state != "visible" or not sheet_has_data(wb[name]):
            abort(404)
        return jsonify(name=name, **sheet_payload(wb, wb[name], band, uid))


def coerce(v):
    if v is None:
        return None
    v = str(v)
    if v.strip() == "":
        return None
    if re.fullmatch(r"-?\d+", v.strip()):
        try:
            return int(v.strip())
        except ValueError:
            return v
    if re.fullmatch(r"-?\d*\.\d+", v.strip()):
        return float(v.strip())
    return v


@app.route("/api/save", methods=["POST"])
@login_required
def api_save():
    """Body: {sheet: 'Name', cells: [{r: 3, c: 2, v: 'text'}, ...]}"""
    uid = session["uid"]
    data = request.get_json(force=True, silent=True) or {}
    sheet, cells = data.get("sheet"), data.get("cells", [])
    with user_lock(uid):
        path = user_file(uid)
        wb = get_wb(uid)
        if sheet not in wb.sheetnames or wb[sheet].sheet_state != "visible":
            return jsonify(ok=False, error="Unknown sheet"), 404
        ws = wb[sheet]
        for item in cells:
            try:
                r, c = int(item["r"]), int(item["c"])
            except (KeyError, ValueError, TypeError):
                continue
            if not (1 <= r <= MAX_ROWS and 1 <= c <= MAX_COLS):
                continue
            cell = ws.cell(row=r, column=c)
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue  # formulas are protected
            try:
                cell.value = coerce(item.get("v"))
            except AttributeError:   # merged (read-only) cell
                continue
        tmp = path + ".tmp"
        wb.save(tmp)
        os.replace(tmp, path)
        _wb_cache[uid] = (os.path.getmtime(path), wb)
        calc = [{"r": r, "c": c, "v": v} for (r, c), v in calc_values(ws).items()]
        ver = data_version(uid)
    return jsonify(ok=True, saved=len(cells), calc=calc, ver=ver)


# ------------------------------------------------------------------------ photos
def img_path(uid, sheet, r, c, i=0):
    """Each picture area holds up to two photos: i = 0 (first) or 1 (second)."""
    slug = re.sub(r"[^A-Za-z0-9]+", "_", sheet).strip("_") or "sheet"
    suffix = "_b" if int(i) == 1 else ""
    return os.path.join(USER_FILES, f"images_{uid}", f"{slug}_{int(r)}_{int(c)}{suffix}.jpg")


def _slot_args(src):
    try:
        i = int(src.get("i", 0) or 0)
        return str(src.get("sheet", "")), int(src.get("r")), int(src.get("c")), (1 if i == 1 else 0)
    except (TypeError, ValueError):
        abort(400)


# ---- layout of the photos in a picture area (1 or 2 photos, side by side / stacked, split) ----
def layout_path(uid):
    return os.path.join(USER_FILES, f"layout_{uid}.json")


def layout_key(sheet, r, c):
    return f"{sheet}|{int(r)}|{int(c)}"


def norm_layout(d):
    d = d if isinstance(d, dict) else {}
    try:
        split = max(20, min(80, int(d.get("split", 50))))
    except (TypeError, ValueError):
        split = 50
    return {"n": 2 if d.get("n") == 2 else 1,
            "mode": "stack" if d.get("mode") == "stack" else "side", "split": split}


def load_layout(uid):
    try:
        with open(layout_path(uid)) as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def data_version(uid):
    """Changes whenever this login's report data, photos or photo layout change (used for live sync)."""
    parts = []
    for p in (user_file(uid), layout_path(uid), os.path.join(USER_FILES, f"images_{uid}")):
        try:
            parts.append(str(os.stat(p).st_mtime_ns))
        except OSError:
            parts.append("0")
    d = os.path.join(USER_FILES, f"images_{uid}")
    if os.path.isdir(d):
        for n in sorted(os.listdir(d)):
            try:
                parts.append(str(os.stat(os.path.join(d, n)).st_mtime_ns))
            except OSError:
                pass
    return "-".join(parts)


@app.route("/api/version")
@login_required
def api_version():
    uid = session["uid"]
    with user_lock(uid):
        return jsonify(v=data_version(uid))


@app.route("/api/layout", methods=["POST"])
@login_required
def api_layout():
    uid = session["uid"]
    data = request.get_json(force=True, silent=True) or {}
    sheet, r, c, _i = _slot_args(data)
    with user_lock(uid):
        wb = get_wb(uid)
        if sheet not in wb.sheetnames or image_slots(wb[sheet]).get((r, c)) != "both":
            return jsonify(ok=False, error="This area takes a single picture."), 400
        lays = load_layout(uid)
        lays[layout_key(sheet, r, c)] = norm_layout(data)
        tmp = layout_path(uid) + ".tmp"
        with open(tmp, "w") as f:
            json.dump(lays, f)
        os.replace(tmp, layout_path(uid))
        return jsonify(ok=True, v=data_version(uid))


@app.route("/api/image")
@login_required
def api_image():
    sheet, r, c, i = _slot_args(request.args)
    path = img_path(session["uid"], sheet, r, c, i)
    if not os.path.exists(path):
        abort(404)
    return send_file(path, mimetype="image/jpeg", max_age=0)


@app.route("/api/image", methods=["POST"])
@login_required
def api_image_upload():
    """Form fields: sheet, r, c, i (0 or 1), file  (a photo taken with the camera or picked from the device)."""
    uid = session["uid"]
    sheet, r, c, i = _slot_args(request.form)
    with user_lock(uid):
        wb = get_wb(uid)
        mode = image_slots(wb[sheet]).get((r, c)) if sheet in wb.sheetnames else None
        if mode is None or (i == 1 and mode != "both"):
            return jsonify(ok=False, error="This area does not take a picture."), 400
    f = request.files.get("file")
    if not f:
        return jsonify(ok=False, error="No picture received."), 400
    try:
        im = ImageOps.exif_transpose(Image.open(f.stream))
        if im.mode in ("RGBA", "LA", "P"):
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im.convert("RGBA"), mask=im.convert("RGBA").split()[-1])
            im = bg
        im = im.convert("RGB")
        im.thumbnail((1600, 1600))
    except Exception:
        return jsonify(ok=False, error="That file is not a valid picture."), 400
    path = img_path(uid, sheet, r, c, i)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp.jpg"
    im.save(tmp, "JPEG", quality=85, optimize=True)
    os.replace(tmp, path)
    with user_lock(uid):
        return jsonify(ok=True, v=int(os.path.getmtime(path) * 1000), ver=data_version(uid))


@app.route("/api/image/delete", methods=["POST"])
@login_required
def api_image_delete():
    data = request.get_json(force=True, silent=True) or {}
    sheet, r, c, i = _slot_args(data)
    uid = session["uid"]
    path = img_path(uid, sheet, r, c, i)
    if os.path.exists(path):
        os.remove(path)
    with user_lock(uid):
        return jsonify(ok=True, ver=data_version(uid))


# ---------------------------------------------------------------------- downloads
def _xlsx_with_photos(uid):
    """The user's workbook as a fresh in-memory copy with their photos placed in the picture areas."""
    wb = load_workbook(user_file(uid))
    lays = load_layout(uid)
    for ws in wb.worksheets:
        for (r, c), mode in image_slots(ws).items():
            rng = next((m for m in ws.merged_cells.ranges if m.min_row == r and m.min_col == c), None)
            if rng is None:
                continue
            col_w = column_widths(ws, rng.min_col, rng.max_col)
            row_h = [(ws.row_dimensions[i].height or ws.sheet_format.defaultRowHeight or 15) * 1.333
                     for i in range(rng.min_row, rng.max_row + 1)]
            area_w, area_h = sum(col_w), sum(row_h)
            rects = photo_rects(norm_layout(lays.get(layout_key(ws.title, r, c))) if mode == "both" else None)
            for i, (fx, fy, fw, fh) in enumerate(rects):
                path = img_path(uid, ws.title, r, c, i)
                if not os.path.exists(path):
                    continue
                pad = 4
                sx, sy = area_w * fx + pad, area_h * fy + pad
                aw, ah = area_w * fw - 2 * pad, area_h * fh - 2 * pad
                pic = XLImage(path)
                k = min(aw / pic.width, ah / pic.height)
                w, h = pic.width * k, pic.height * k
                off_x, off_y = sx + (aw - w) / 2, sy + (ah - h) / 2          # px from the area's corner
                col = rng.min_col
                for cw in col_w:                                              # keep the anchor inside a cell
                    if off_x < cw:
                        break
                    off_x -= cw
                    col += 1
                row = rng.min_row
                for rh in row_h:
                    if off_y < rh:
                        break
                    off_y -= rh
                    row += 1
                pic.anchor = OneCellAnchor(
                    _from=AnchorMarker(col=col - 1, colOff=int(off_x * 9525), row=row - 1, rowOff=int(off_y * 9525)),
                    ext=XDRPositiveSize2D(int(w * 9525), int(h * 9525)))
                ws.add_image(pic)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/report/excel")
@login_required
def download_excel():
    uid = session["uid"]
    with user_lock(uid):
        try:
            buf = _xlsx_with_photos(uid)
        except Exception:                      # never lose the download because of a picture
            with open(user_file(uid), "rb") as f:
                buf = io.BytesIO(f.read())
    return send_file(buf, as_attachment=True, download_name="my_report.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/report/pdf")
@login_required
def download_pdf():
    """Every page of the report as one A4 PDF (same look as the sample inspection report)."""
    uid = session["uid"]
    with user_lock(uid):
        wb = get_wb(uid)
        theme = theme_colors(wb)
        lays = load_layout(uid)
        data = build_pdf(visible_sheets(wb),
                         color=lambda col: color_of(col, theme),
                         calc=calc_values, fmt_number=fmt_number,
                         image_path=lambda sheet, r, c, i=0: img_path(uid, sheet, r, c, i),
                         layout=lambda sheet, r, c: norm_layout(lays.get(layout_key(sheet, r, c))))
    inline = request.args.get("inline") == "1"          # preview = shown in the page, not downloaded
    return send_file(io.BytesIO(data), as_attachment=not inline, download_name="my_report.pdf",
                     mimetype="application/pdf")


init_db()
ensure_template()

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
