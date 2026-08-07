# WAE Label Scanner

Point a phone camera at a Beyond Gravity / WAE label, read the fields, watch the
table grow, export to Excel. A single-file proof of concept meant to grow into a
proper in-tenant tool later.

## What it does

One scan runs a fixed pipeline:

1. **Barcodes + datamatrix** are read first (ZXing). These give the *reliable*
   fields — Mat, CH-Nr, the 10-digit L-Char. Shown with a **green** dot.
2. **OCR** reads the remaining text (Tesseract.js) — Definition, Rev, Char,
   Auftr/Menge, Verfallsdatum. Shown with an **amber** dot: glance-check these.
3. A **review card** lets you fix anything before it lands in the table.
4. **Add row** pushes it into the live on-screen table.
5. **Export .xlsx** turns the whole table into a real Excel file (SheetJS).

The on-screen table *is* the live-incrementing sheet — it grows as you scan and
survives a page reload (rows are cached in the browser). The `.xlsx` file itself
is written only when you tap Export. A browser cannot write into an open Excel
file row-by-row; that's a backend job for the future version.

## Columns

`Mat · Definition · Rev · SerNr · Char · L-Char · Auftr · Menge · Einheit · Verfallsdatum · CH-Nr`

## Run it

### On a laptop (camera works via localhost)
Install the **Live Server** VS Code extension (per-user, no admin), then
right-click `index.html` -> *Open with Live Server*.

### On a phone (live camera)
The camera only works over **https** or localhost — not from a file sent over
WhatsApp/email (that opens as `file://` and the browser blocks the camera).
Host it on **GitHub Pages** (free, https, no admin):

1. Create a public repo, upload `index.html` + `README.md`.
2. Repo **Settings -> Pages -> Branch: main -> /(root) -> Save**.
3. Open the `https://<user>.github.io/<repo>/` link on your phone.

### Without hosting
The **Upload photo** button works everywhere, including from a file sent over
WhatsApp. No live camera, but the full read-and-export pipeline works.

## Extending it (the parts that matter later)

Everything lives in `index.html`, bottom `<script>`:

- **`COLS`** (top of the script) — the columns. Add/rename here and it flows
  through the review card, table, and Excel export. Single source of truth.
- **`applyBarcodes()`** — rules mapping decoded codes to fields. Reliable core.
- **`applyOcr()`** — regex rules pulling fields from OCR text. **This is what
  you tune** as you meet new label layouts (Helicalcoil vs. cable-harness vs.
  RUAG-relabeled all differ slightly).
- **Design tokens** — top of `<style>`, the `:root` block. Change `--signal`
  to reshuffle the accent colour; the whole identity keys off it.

## Libraries

Loaded from CDN for the POC:
- [ZXing](https://github.com/zxing-js/library) — barcode / datamatrix
- [Tesseract.js](https://tesseract.projectnaptha.com/) — OCR
- [SheetJS](https://sheetjs.com/) — Excel export

For a locked-down in-tenant deployment behind a strict firewall, vendor these
three files locally instead of using the CDN.

## Path to the real tool

The single HTML file is the throwaway packaging. The reusable core is the
`COLS` schema + the `applyOcr` / `applyBarcodes` logic — that ports straight
into a FastAPI backend (server-side OCR, shared table written to SharePoint)
with this page as the client. That is the Stage-2 version; this is the POC.
