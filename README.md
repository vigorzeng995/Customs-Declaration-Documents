# Manifold Tech — Customs Declaration Document Generator

A local web tool that turns a sales-order PDF into a complete, print-ready set of
customs declaration documents — one upload, six linked worksheets.

Built for the routine where every shipment shares the same buyer/seller and only the
line items, invoice number and contract number change.

The interface is in Chinese, matching the declaration paperwork it produces.
The page is plain white; the only image it ever shows is your own background, if you
upload one. `logo.png` is the header mark and can be swapped or deleted.

## What it does

1. **Upload the HS lookup table** (an Excel file mapping internal ERP codes → customs HS codes).
2. **Upload the sales order PDF** — customer, quote number, item code, quantity and unit price are parsed automatically.
3. **One click to generate** a single workbook containing:

| # | Sheet | Content |
|---|-------|---------|
| 1 | 报关单 Customs declaration | Item no. / HS code / description / qty / unit / unit price / origin / destination / exemption |
| 2 | 申报要素 Declaration elements | Sequence, description, HS, brand, model, brand type, preferential treatment |
| 3 | 发票 Invoice | Line items, unit price, amount, total |
| 4 | 箱单 Packing list | Cartons, quantity, unit, gross/net weight |
| 5 | 合同 Contract | Description, qty, unit, unit price, amount, total value |
| 6 | 不报关填写快递 Courier form | Placeholder sheet for non-declared shipments |

Everything is written directly into the XLSX package at the XML level, so the original
template's logos, borders, column widths and merged cells are preserved exactly —
nothing is re-saved through a spreadsheet library.

## Requirements

- Python 3.9+
- `pip install -r requirements.txt`
- A PDF text extractor, one of:
  - **poppler** (`pdftotext` on `PATH`) — recommended, works everywhere
  - the `workbuddy-pdf` skill directory, pointed to by the `PDF_SKILL_DIR` environment variable

If neither is available the tool still starts; only PDF import will report a clear error.

## Quick start

```bash
pip install -r requirements.txt
python app.py
```

Then open <http://127.0.0.1:8765>.

On Windows you can also double-click `启动.bat` (starts the server and opens the browser).

## Changing the background

The page is plain white by default — no built-in artwork. The **页面外观** card lets you
put your own picture behind the cards (JPG / PNG / WebP / GIF, up to 12 MB) and fade it
with the opacity slider, so text stays readable on a busy photo.

- The image is never sent anywhere: it is validated by its real file signature (not the
  extension), auto-rotated by EXIF, capped at 2560 px on the long side, flattened onto
  white if it has transparency, and stored as `custom_bg.webp` next to `app.py`.
- The picture and its opacity are remembered in `appearance.json` (both git-ignored,
  both local only).
- **恢复默认（白底）** deletes them and returns to the plain white page.
- Bad input (not an image, truncated file, >12 MB, >60 megapixels, SVG) is rejected with a
  readable message and leaves the previous background untouched.
- Pillow is optional: without it the file is stored as-is instead of being compressed.

## Data files

Two workbooks drive the logic. Real ones are **git-ignored** — this repo ships sanitized
samples instead, which the app copies into place on first run:

| File | Purpose | Sample in repo |
|------|---------|----------------|
| `master.xlsx` | ERP → HS / description / unit / price / weight lookup, used at generation time | `master.example.xlsx` |
| `HS对照表_待填.xlsx` | Blank table offered for download in the UI, filled in and uploaded back | `HS对照表_待填.example.xlsx` |

To use your own data, either replace those two files or upload your filled lookup table
through the UI (it overwrites `master.xlsx`).

`make_master.py` can rebuild the blank lookup table from a BOM dump placed at
`master_seed/bom_full.json` (not included).

## How the sheets stay in sync

The declaration sheet is the single source of truth; the other sheets reference it with
formulas. The generator writes only values and formulas and never touches styling, so
the template remains the authority on layout.

Two details worth knowing if you adapt the template:

- `fullCalcOnLoad` is set in `workbook.xml`, otherwise Excel shows cached totals from the
  previous workbook.
- Non-anchor cells inside a merged range are kept (empty, same style). Deleting them —
  even though they hold no value — breaks the merged area's border, since Excel draws
  it from each individual cell's border definition.

## Project layout

```
app.py                        Flask app: web UI, PDF parsing, XLSX writer
make_master.py                Builds the blank lookup table from a BOM dump
base_template_v2.xlsx         Document template (the 6 sheets)
base_template.xlsx            Legacy template, kept for the old code path
master.example.xlsx           Sanitized lookup sample (7 columns)
HS对照表_待填.example.xlsx      Sanitized lookup sample (9 columns)
logo.png                      Brand mark shown in the page header (swap for your own)
custom_bg.webp                User-uploaded background, created at runtime (git-ignored)
appearance.json               Background / opacity preference, created at runtime (git-ignored)
bg-sky.webp, robot.webp       Artwork from the earlier full-bleed theme, no longer used
启动.bat                       Windows one-click launcher
requirements.txt
```

## Notes

- Designed for a single user running it locally; no authentication, no database.
- Generated files land in `output/`; uploads and parsed caches in `uploads/`.
- If the target file is locked by Excel, the app falls back to `…(新).xlsx` and then to a
  timestamped name instead of failing.
