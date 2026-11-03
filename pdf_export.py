"""PDF export for the report.

Every page of the Excel format is drawn straight onto an A4 page with reportlab, using the
real cell sizes, colours, merged areas and borders of the workbook, so the PDF looks like the
sample inspection report (header + logo, orange titles, blue bottom bar, footer line).

* Cover / contents / introduction / carpet area / summary / disclaimer / thank-you pages are
  printed as they are.
* QUALITY CHECK: every card that has an element, an observation or a photo becomes one page
  (empty cards are skipped).
* DIMENSION MEASUREMENT is printed only when something was typed into its table.
* Photos added with "Take Photo" / "Upload" are drawn inside the picture area of their page.
"""
import io
import os
import re

from openpyxl.utils import column_index_from_string
from reportlab.lib.colors import HexColor, black
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FONT_DIR = os.path.join(BASE_DIR, "static", "fonts")
LOGO_PATH = os.path.join(BASE_DIR, "static", "img", "logo.jpg")

TITLE = "CIVIL-VENTURE HOME INSPECTION REPORT"
PEACH = "#FCE4D6"
BAR_COLORS = {"FF92ABF0", "FF93AFFF"}          # the blue bar that closes every page

PAGE_W, PAGE_H = 595.28, 841.89                 # A4 portrait, points
MARGIN_X, MARGIN_TOP, MARGIN_BOTTOM = 30, 34, 58
LINE = 1.22                                     # line height / font size (Calibri-like)
ASCENT = 0.95

BORDER_PT = {"thin": 0.8, "hair": 0.4, "dotted": 0.8, "dashed": 0.8, "medium": 1.6,
             "mediumDashed": 1.6, "thick": 2.4, "double": 2.4}

# --------------------------------------------------------------------- page layout helpers


def header_ranges(ws):
    """The empty 9-column x 4-row merged block at the top of every page / card."""
    out = []
    for m in ws.merged_cells.ranges:
        if (m.max_row - m.min_row == 3 and m.max_col - m.min_col == 8
                and ws.cell(m.min_row, m.min_col).value in (None, "")):
            out.append(m)
    return sorted(out, key=lambda m: (m.min_col, m.min_row))


# Quality-check pages: the orange bar under the title is a drop-down list of rooms / areas
AREA_OPTIONS = ["HALL", "KITCHEN", "BEDROOM 1", "BEDROOM 2", "BEDROOM 3", "BEDROOM 4",
                "TOILET 1", "TOILET 2", "TOILET 3", "DRY BALCONY", "HALL BALCONY",
                "BED 1 BALCONY", "BED 2 BALCONY", "BED 3 BALCONY", "TOP TERRACE", "GARDEN",
                "GF STAIRS", "1ST FLOOR STAIRS", "2ND FLOOR STAIRS", "PORCH", "FAMILY LOUNGE",
                "PASSAGE", "HALL LOBBY"]


def area_bars(ws):
    """{(row, col): True} for the orange bar of every QUALITY CHECK card (merged B7:J8 of each card)."""
    out = {}
    if ws.title == "QUALITY CHECK":
        for m in ws.merged_cells.ranges:
            if m.min_row == 7 and m.max_row == 8 and m.max_col - m.min_col == 8:
                out[(7, m.min_col)] = True
    return out


def image_slots(ws):
    """Picture areas: {(row, col) of the merged area's top-left cell: 'both' | 'upload'}.
    QUALITY CHECK cards -> Take Photo + Upload.  CARPET AREA drawing -> Upload Image only."""
    out = {}
    if ws.title == "QUALITY CHECK":
        for m in ws.merged_cells.ranges:
            if m.min_row == 15 and m.max_row == 32 and m.max_col - m.min_col >= 4:
                out[(15, m.min_col)] = "both"
    elif ws.title == "CARPET AREA":
        for m in ws.merged_cells.ranges:
            if m.min_row == 52 and m.max_row == 79:
                out[(52, m.min_col)] = "upload"
    return out


def photo_rects(lay):
    """Where the photos go inside a picture area, as fractions (x, y, w, h) of the area.
    lay = {"n": 1 | 2, "mode": "side" | "stack", "split": 20..80 (% given to the first photo)}"""
    lay = lay or {}
    if int(lay.get("n", 1)) < 2:
        return [(0.0, 0.0, 1.0, 1.0)]
    sp = max(20, min(80, int(lay.get("split", 50)))) / 100.0
    if lay.get("mode") == "stack":
        return [(0.0, 0.0, 1.0, sp), (0.0, sp, 1.0, 1 - sp)]
    return [(0.0, 0.0, sp, 1.0), (sp, 0.0, 1 - sp, 1.0)]


def _is_bar(cell):
    f = cell.fill
    return bool(f and f.fill_type == "solid" and f.fgColor is not None
                and f.fgColor.type == "rgb" and f.fgColor.rgb in BAR_COLORS)


def page_blocks(ws):
    """[(first_row, last_row, first_col, last_col)] - one per printed page of the sheet."""
    blocks = []
    for m in header_ranges(ws):
        end = None
        for r in range(m.min_row + 4, min(ws.max_row, m.min_row + 80) + 1):
            if _is_bar(ws.cell(r, m.min_col)):
                end = r
                break
        blocks.append((m.min_row, end or m.min_row + 38, m.min_col, m.max_col))
    return blocks


def _filled(v):
    return v is not None and str(v).strip() != ""


def _include(ws, blk, image_path):
    r0, r1, c0, c1 = blk
    if ws.title == "QUALITY CHECK":
        if (_filled(ws.cell(7, c0).value) or _filled(ws.cell(11, c0).value)
                or _filled(ws.cell(11, c0 + 4).value)):
            return True
        return os.path.exists(image_path(ws.title, 15, c0)) or os.path.exists(image_path(ws.title, 15, c0, 1))
    if ws.title == "DIMENSION MEASUREMENT":
        return any(_filled(ws.cell(r, c).value) for r in range(r0 + 15, r1) for c in range(5, 11))
    return True


# ---------------------------------------------------------------- dimension page symbols
# Simple line symbols drawn above the DOOR / WINDOW / WALL TILE headings, in a 100 x 100 box
# (y grows downwards).  Same spec is used for the PDF and for the on-screen page.
ICONS = {
    "door": [("rect", 24, 4, 52, 92), ("rect", 33, 13, 34, 32), ("rect", 33, 55, 34, 32),
             ("circle", 66, 52, 3.5)],
    "window": [("rect", 12, 8, 76, 84), ("line", 50, 8, 50, 92), ("line", 12, 50, 88, 50),
               ("rect", 5, 92, 90, 5)],
    "tile": [("rect", 8, 8, 84, 84), ("line", 36, 8, 36, 92), ("line", 64, 8, 64, 92),
             ("line", 8, 36, 92, 36), ("line", 8, 64, 92, 64)],
}


def dimension_icons(ws):
    """{(row, col) of a heading block: icon name} for the DIMENSION MEASUREMENT page."""
    out = {}
    if ws.title == "DIMENSION MEASUREMENT":
        for m in ws.merged_cells.ranges:
            if m.min_row == 12 and m.max_row == 17:
                name = {5: "door", 7: "window", 9: "tile"}.get(m.min_col)
                if name:
                    out[(12, m.min_col)] = name
    return out


def icon_svg(name):
    parts = []
    for shape in ICONS[name]:
        if shape[0] == "rect":
            parts.append('<rect x="%s" y="%s" width="%s" height="%s"/>' % shape[1:])
        elif shape[0] == "line":
            parts.append('<line x1="%s" y1="%s" x2="%s" y2="%s"/>' % shape[1:])
        else:
            parts.append('<circle cx="%s" cy="%s" r="%s"/>' % shape[1:])
    return ('<svg viewBox="0 0 100 100" fill="none" stroke="#fff" stroke-width="4" '
            'stroke-linejoin="round" stroke-linecap="round" aria-hidden="true">%s</svg>' % "".join(parts))


def _draw_icon(c, name, x, y, size):
    """Draw icon `name` in the square whose top-left corner is (x, y) (PDF points, y = top edge)."""
    k = size / 100.0
    c.saveState()
    c.setStrokeColor(HexColor("#FFFFFF"))
    c.setLineWidth(max(1.2, 4 * k))
    c.setLineJoin(1)
    c.setLineCap(1)
    for shape in ICONS[name]:
        if shape[0] == "rect":
            _, sx, sy, w, h = shape
            c.rect(x + sx * k, y - (sy + h) * k, w * k, h * k, stroke=1, fill=0)
        elif shape[0] == "line":
            _, x1, y1, x2, y2 = shape
            c.line(x + x1 * k, y - y1 * k, x + x2 * k, y - y2 * k)
        else:
            _, cx, cy, r = shape
            c.circle(x + cx * k, y - cy * k, r * k, stroke=1, fill=0)
    c.restoreState()


# ------------------------------------------------------------------------------- fonts
_fonts_ready = False


def _register_fonts():
    global _fonts_ready
    if _fonts_ready:
        return
    for name, fn in (("Carlito", "Carlito-Regular.ttf"), ("Carlito-Bold", "Carlito-Bold.ttf"),
                     ("Carlito-Italic", "Carlito-Italic.ttf"),
                     ("Carlito-BoldItalic", "Carlito-BoldItalic.ttf")):
        pdfmetrics.registerFont(TTFont(name, os.path.join(FONT_DIR, fn)))
    _fonts_ready = True


def _font(name, bold, italic, text):
    suffix = ("Bold" if bold else "") + ("Italic" if italic else "")
    if (name or "").lower() in ("arial", "helvetica") and all(ord(ch) < 256 for ch in text):
        return {"": "Helvetica", "Bold": "Helvetica-Bold", "Italic": "Helvetica-Oblique",
                "BoldItalic": "Helvetica-BoldOblique"}[suffix]
    return "Carlito-" + suffix if suffix else "Carlito"


def _wrap(text, font, size, max_w):
    lines = []
    for para in str(text).replace("\r", "").split("\n"):
        words, cur = para.split(" "), ""
        for w in words:
            trial = w if not cur else cur + " " + w
            if pdfmetrics.stringWidth(trial, font, size) <= max_w or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = w
            while pdfmetrics.stringWidth(cur, font, size) > max_w and len(cur) > 1:   # very long word
                k = len(cur)
                while k > 1 and pdfmetrics.stringWidth(cur[:k], font, size) > max_w:
                    k -= 1
                lines.append(cur[:k])
                cur = cur[k:]
        lines.append(cur)
    return lines


# ----------------------------------------------------------------------------- geometry
def _axes(ws, r0, r1, c0, c1):
    default_w = ws.sheet_format.defaultColWidth or 8.43
    widths, hidden = {c: default_w for c in range(c0, c1 + 1)}, set()
    for dim in ws.column_dimensions.values():
        lo = dim.min or column_index_from_string(dim.index)
        hi = dim.max or lo
        for c in range(max(lo, c0), min(hi, c1) + 1):
            if dim.width:
                widths[c] = dim.width
            if dim.hidden:
                hidden.add(c)
    xs = [0.0]
    for c in range(c0, c1 + 1):
        xs.append(xs[-1] + (0 if c in hidden else (widths[c] * 7 + 5) * 0.75))
    default_h = ws.sheet_format.defaultRowHeight or 15
    ys = [0.0]
    for r in range(r0, r1 + 1):
        d = ws.row_dimensions.get(r)
        h = 0 if (d and d.hidden) else ((d.height if d and d.height else default_h))
        ys.append(ys[-1] + h)
    return xs, ys


def _footer_parts(ws, page_no, total):
    f = ws.oddFooter
    out = []
    for part in (f.left, f.center, f.right):
        t = (part.text if part and part.text else "")
        t = t.replace("&P", str(page_no)).replace("&N", str(total))
        t = re.sub(r'&"[^"]*"', "", t)
        t = re.sub(r"&\d+", "", t)
        t = re.sub(r"&[A-Za-z]", "", t)
        out.append(t.strip())
    return out


# ------------------------------------------------------------------------------ drawing
def _draw_page(c, page, ctx, page_no, total):
    ws = page["ws"]
    r0, r1, c0, c1 = page["r0"], page["r1"], page["c0"], page["c1"]
    xs, ys = _axes(ws, r0, r1, c0, c1)
    bw, bh = xs[-1], ys[-1]
    s = min((PAGE_W - 2 * MARGIN_X) / bw, (PAGE_H - MARGIN_TOP - MARGIN_BOTTOM) / bh, 1.4)
    ox = (PAGE_W - bw * s) / 2
    X = lambda x: ox + x * s
    Y = lambda y: PAGE_H - (MARGIN_TOP + y * s)

    merged, interior = {}, set()
    for m in ws.merged_cells.ranges:
        if m.min_row >= r0 and m.max_row <= r1 and m.min_col >= c0 and m.max_col <= c1:
            merged[(m.min_row, m.min_col)] = m
            for rr in range(m.min_row, m.max_row + 1):
                for cc in range(m.min_col, m.max_col + 1):
                    if (rr, cc) != (m.min_row, m.min_col):
                        interior.add((rr, cc))
    cells = []
    for r in range(r0, r1 + 1):
        for cc in range(c0, c1 + 1):
            if (r, cc) in interior:
                continue
            m = merged.get((r, cc))
            rb, cb = (m.max_row, m.max_col) if m else (r, cc)
            x0, x1, y0, y1 = xs[cc - c0], xs[cb - c0 + 1], ys[r - r0], ys[rb - r0 + 1]
            if x1 - x0 > 0 and y1 - y0 > 0:
                cells.append((r, cc, rb, cb, x0, y0, x1, y1))

    slots = image_slots(ws)
    special = {(r0, c0)} | {k for k in slots if k in merged}

    # 1) fills
    for r, cc, rb, cb, x0, y0, x1, y1 in cells:
        cell = ws.cell(r, cc)
        if cell.fill and cell.fill.fill_type == "solid":
            col = ctx.color(cell.fill.fgColor)
            if col:
                c.setFillColor(HexColor(col))
                c.rect(X(x0), Y(y1), (x1 - x0) * s, (y1 - y0) * s, stroke=0, fill=1)

    # 2) page header: peach band, title, logo
    hx1, hy1 = xs[-1], ys[4]
    c.setFillColor(HexColor(PEACH))
    c.rect(X(0), Y(hy1), hx1 * s, hy1 * s, stroke=0, fill=1)
    c.setFillColor(black)
    c.setFont("Helvetica-Bold", 12.5 * s)
    c.drawString(X(0) + 8 * s, Y(hy1 / 2) - 4.4 * s, TITLE)
    if os.path.exists(LOGO_PATH):
        lh = hy1 - 6
        lw = lh * 309 / 167
        c.drawImage(LOGO_PATH, X(hx1 - lw - 3), Y(hy1 - 3), lw * s, lh * s)

    # 3) photos
    for (r, cc), _mode in slots.items():
        if (r, cc) not in merged:
            continue
        cell = next(t for t in cells if t[0] == r and t[1] == cc)
        _, _, _, _, x0, y0, x1, y1 = cell
        rects = photo_rects(ctx.layout(ws.title, r, cc) if _mode == "both" else None)
        for i, (fx, fy, fw, fh) in enumerate(rects):
            path = ctx.image_path(ws.title, r, cc, i)
            if not os.path.exists(path):
                continue
            pad = 6 if len(rects) == 1 else 4
            sx0, sy0 = x0 + (x1 - x0) * fx, y0 + (y1 - y0) * fy
            aw, ah = (x1 - x0) * fw - 2 * pad, (y1 - y0) * fh - 2 * pad
            img = ImageReader(path)
            iw, ih = img.getSize()
            k = min(aw / iw, ah / ih)
            dw, dh = iw * k, ih * k
            c.drawImage(img, X(sx0 + pad + (aw - dw) / 2), Y(sy0 + pad + (ah - dh) / 2 + dh),
                        dw * s, dh * s)

    # 4) text
    calc = ctx.calc(ws)
    for r, cc, rb, cb, x0, y0, x1, y1 in cells:
        if (r, cc) in special:
            continue
        cell = ws.cell(r, cc)
        raw = cell.value
        if isinstance(raw, str) and raw.startswith("="):
            text = calc.get((r, cc), "")
            numeric = bool(re.fullmatch(r"-?\d+(\.\d+)?", str(text).strip()))
        elif isinstance(raw, (int, float)) and not isinstance(raw, bool):
            fmt = cell.number_format or ""
            text, numeric = (ctx.fmt_number(raw, fmt) if "0." in fmt else str(raw)), True
        else:
            text, numeric = ("" if raw is None else str(raw)), False
        if not text.strip():
            continue
        f = cell.font
        size = (f.sz if f and f.sz else 11) * s
        font = _font(f.name if f else None, bool(f and f.b), bool(f and f.i), text)
        col = ctx.color(f.color) if f else None
        al = cell.alignment
        h = al.horizontal if al else None
        h = {"centerContinuous": "center"}.get(h, h)
        if h not in ("left", "center", "right"):
            h = "right" if numeric else "left"
        v = (al.vertical if al and al.vertical else "bottom")
        wrap = bool(al and al.wrap_text)
        pad = 2.2 * s
        cw, ch = (x1 - x0) * s, (y1 - y0) * s
        lines = _wrap(text, font, size, cw - 2 * pad) if wrap else [text.replace("\n", " ")]
        lh = size * LINE
        total_h = lh * len(lines)
        vpad = 0.6 * s
        if total_h > ch * 1.08:                      # does not fit: show the beginning of the text
            v = "top"
        top = (Y(y0) - vpad if v == "top" else
               Y(y0) - (ch - total_h) / 2 if v in ("center", "distributed", "justify") else
               Y(y1) + vpad + total_h)
        c.saveState()
        p = c.beginPath()
        p.rect(X(x0), Y(y1), cw, ch)
        if wrap or total_h > ch * 1.08:
            c.clipPath(p, stroke=0, fill=0)
        c.setFillColor(HexColor(col) if col else black)
        c.setFont(font, size)
        for i, line in enumerate(lines):
            base = top - i * lh - ASCENT * size
            if h == "center":
                c.drawCentredString(X(x0) + cw / 2, base, line)
            elif h == "right":
                c.drawRightString(X(x1) - pad, base, line)
            else:
                c.drawString(X(x0) + pad, base, line)
            if f and f.u:
                wdt = pdfmetrics.stringWidth(line, font, size)
                sx = (X(x0) + cw / 2 - wdt / 2 if h == "center" else
                      X(x1) - pad - wdt if h == "right" else X(x0) + pad)
                c.setStrokeColor(HexColor(col) if col else black)
                c.setLineWidth(0.6)
                c.line(sx, base - 1.5, sx + wdt, base - 1.5)
        c.restoreState()

    # 4b) symbols on the dimension page
    for (r, cc), name in dimension_icons(ws).items():
        cell = next((t for t in cells if t[0] == r and t[1] == cc), None)
        if cell:
            _, _, _, _, x0, y0, x1, y1 = cell
            size = min((x1 - x0) * 0.5, (y1 - y0) * 0.55)
            _draw_icon(c, name, X((x0 + x1) / 2 - size / 2), Y(y0 + (y1 - y0) * 0.08), size * s)

    # 5) borders (on top of fills)
    def side(border, x_a, y_a, x_b, y_b):
        if border and border.style:
            col = ctx.color(border.color) if border.color else None
            c.setStrokeColor(HexColor(col) if col else black)
            c.setLineWidth(BORDER_PT.get(border.style, 0.8) * min(s, 1.2))
            c.line(X(x_a), Y(y_a), X(x_b), Y(y_b))

    for r, cc, rb, cb, x0, y0, x1, y1 in cells:
        tl = ws.cell(r, cc)
        side(tl.border.top, x0, y0, x1, y0)
        side(ws.cell(rb, cc).border.bottom, x0, y1, x1, y1)
        side(tl.border.left, x0, y0, x0, y1)
        side(ws.cell(r, cb).border.right, x1, y0, x1, y1)

    # 6) footer line (url / page name / phone numbers)
    left, centre, right = _footer_parts(ws, page_no, total)
    c.setFillColor(black)
    c.setFont("Carlito", 10)
    fy = 30
    c.drawString(MARGIN_X, fy, left)
    c.drawCentredString(PAGE_W / 2, fy, centre)
    c.drawRightString(PAGE_W - MARGIN_X, fy, right)


class Ctx:
    def __init__(self, color, calc, fmt_number, image_path, layout=None):
        self.color, self._calc, self.fmt_number, self.image_path = color, calc, fmt_number, image_path
        self.layout = layout or (lambda sheet, r, c: {})
        self._cache = {}

    def calc(self, ws):
        if ws.title not in self._cache:
            self._cache[ws.title] = self._calc(ws)
        return self._cache[ws.title]


def _sort_by_area(ws, blks):
    """QUALITY CHECK pages are printed room by room (HALL, KITCHEN, BEDROOM 1, ...), not in the order the
    cards were filled. Cards of the same room keep their card order; typed-in (OTHER) rooms follow the
    listed ones in the order they first appear; cards with no room come last."""
    order = {a: i for i, a in enumerate(AREA_OPTIONS)}
    custom = {}

    def key(item):
        idx, blk = item
        v = ws.cell(7, blk[2]).value
        name = " ".join(str(v).split()).upper() if _filled(v) else ""
        if not name:
            return (2, 0, idx)
        if name in order:
            return (0, order[name], idx)
        custom.setdefault(name, len(custom))
        return (1, custom[name], idx)

    return [b for _, b in sorted(enumerate(blks), key=key)]


def build_pdf(sheets, color, calc, fmt_number, image_path, layout=None, title="Home Inspection Report"):
    """sheets: worksheets in print order.  Returns the PDF as bytes."""
    _register_fonts()
    ctx = Ctx(color, calc, fmt_number, image_path, layout)
    pages = []
    for ws in sheets:
        blks = [b for b in page_blocks(ws) if _include(ws, b, image_path)]
        if ws.title == "QUALITY CHECK":
            blks = _sort_by_area(ws, blks)
        for blk in blks:
            pages.append(dict(ws=ws, r0=blk[0], r1=blk[1], c0=blk[2], c1=blk[3]))
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H), pageCompression=1)
    c.setTitle(title)
    c.setAuthor("Civil Venture Home Inspection")
    for i, page in enumerate(pages, start=1):
        _draw_page(c, page, ctx, i, len(pages))
        c.showPage()
    if not pages:
        c.drawString(72, 700, "This report has no pages yet.")
        c.showPage()
    c.save()
    return buf.getvalue()
