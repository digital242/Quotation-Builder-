# THERMODRAIN Quotation Builder

Single-file web app (`index.html`) that turns a customer enquiry into a priced quotation
and downloads it as an Excel file in the Thermoset Poly Products QTN format.

Open `index.html` in a browser (or host it as a Claude artifact). No build step.

## Logos
Paste the base64 PNGs into `LOGO_TP` and `LOGO_TD` at the top of the `td-core` script in
`index.html` (with or without the `data:image/png;base64,` prefix). While they are empty the
header shows a text wordmark and the Excel file is generated without images.

## Tests
```
cd tests
npm install                      # exceljs 4.4.0 + playwright 1.56.1 (no browser download needed if Chromium is present)
npm test                         # core + Excel tests, then browser tests
```
- `run-tests.mjs`: parser, prices, totals, reference number, Excel layout/formulas. If LibreOffice
  (`soffice`) and `pdfinfo` are installed it also checks the file prints on one A4 page and that the
  H-column formulas recalculate after a qty edit.
- `fuzz-tests.mjs`: 60 random quotations (every product, frame, chamber, discounts, ₹ + GST / ₹ GST extra / US$ export).
  Checks list prices against the table, totals against an independent calculation, every Excel formula result,
  and (with LibreOffice) recalculation and one-page printing on a sample.
- `ui-tests.mjs`: Playwright (Chromium) tests of the page: typing, picker, download, 360 px width,
  blocked localStorage, dark-mode device setting. The ExcelJS CDN is served locally in these tests.
