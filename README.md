# ReporTrax

# Report Builder (Flask)

Sign up / sign in -> Dashboard -> **Report** -> your fixed Excel format, with every
sheet in a dropdown on the left and the selected sheet on the right. Every edit is
auto-saved to the signed-in user's own copy of the workbook (also when you switch pages).

## Run

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000

## Use YOUR Excel sheet

1. Put your workbook at `template/report_template.xlsx` (it is already your FINAL_REPORT_FORMAT file, cleaned of images).
2. Delete `data/user_files/` if you have already opened a report with the old one
   (each user gets a private copy of the template on first open).
3. Restart the app. Every visible sheet that contains data appears in the dropdown; empty sheets are ignored.

Notes
* Images/pictures already inside the workbook are intentionally not shown or kept (`sanitize.py`).
  Photos are added by the user instead (see below).
* Formula cells are shown read-only (cached values); all other cells are editable.
* Photos open in an editor (crop / fit to area, zoom, rotate, square-circle-line-arrow marks); a zoom slider sits bottom right.
* "Preview" shows the finished A4 report on screen; "Download PDF" (inside the preview) saves it.
* Data: `data/app.db` (users, hashed passwords) and `data/user_files/user_<id>.xlsx`.

## Structure
```
app.py                 Flask app (auth, report API, Excel read/write via openpyxl)
templates/             signup, signin, dashboard, report pages
static/css/style.css   styling
static/js/report.js    dropdown, grid, autosave
template/              <- put report_template.xlsx here
data/                  created at runtime
```

## Notes for this report format
* All 9 pages of FINAL_REPORT_FORMAT contain data and are listed in the dropdown (empty pages are skipped automatically).
* QUALITY CHECK is 200 cards laid out side by side in Excel; the app shows them one at a time with a **Section** selector.
* Calculated cells (e.g. AREA = LENGTH x WIDTH, totals with SUM) update live when you type.
* Images, logos and shapes are removed automatically (`sanitize.py`).

## Photos, PDF and credit (round 2)
* **Photos** - every QUALITY CHECK card has a picture area with **Take Photo** and **Upload**
  (phones open the camera, computers open a small camera window; if no camera is found a file picker opens).
  The CARPET AREA drawing area has **Upload Image** only. Photos are stored per user in
  `data/user_files/images_<id>/` and are drawn into the PDF.
* **Download PDF** (report page toolbar) - one A4 PDF of the whole report in the sample layout:
  header with logo, orange title bars, blue bottom bar, footer line, `Page X of N` on quality-check pages.
  - QUALITY CHECK cards with no room name, element, observation or photo are skipped.
  - DIMENSION MEASUREMENT is printed only when its table has entries.
  - Code: `pdf_export.py` (needs `reportlab` + `Pillow`, both in requirements.txt; font Carlito is bundled in `static/fonts`).
* The header (title + logo) is also shown at the top of every page in the editor.
* "Developed By Pooja Deore" is shown at the bottom centre of every screen (`templates/base.html`).

## Latest additions
* QUALITY CHECK: the orange bar is a drop-down of rooms / areas (choose OTHER to type your own).
* QUALITY CHECK photos: add 1 or 2 photos per card; side by side or top & bottom, with a size slider.
* "Download as" -> PDF or EXCEL (photos are placed in the Excel too).
* Everything is saved on the server per login (`data/`). Signing in again - on any phone or computer - shows all earlier
  work, and two people using the same login see each other's changes within a few seconds.
