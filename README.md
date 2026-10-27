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

## Google Drive client database
Every downloaded quotation can also be saved to Google Drive:
`THERMODRAIN Quotations / CITY / CLIENT / QTN_025_Client.xlsx`, plus one row per quotation in
`THERMODRAIN Quotations / Quotation register.csv` (opens in Google Sheets).
Each client folder also has a Google Sheet, `CLIENT - client data`, with one row per quoted product
(date, quotation ref, client, city, product, size, load class, qty, discount %, net rate, value, quote value).
Each new quotation adds its rows; downloading the same quotation again replaces its rows.
No extra API is needed: the sheet is kept up to date through the Google Drive API.

**Export client data (Excel)** under Saved customers downloads the same rows for every client,
from this browser's history (quotations downloaded in this browser).

The company Client ID is built into `GOOGLE_CLIENT_ID`; each browser still ticks
"Save every downloaded quotation to Google Drive" once and signs in.

One-time setup (company Google account):
1. console.cloud.google.com: create a project, enable **Google Drive API**.
2. OAuth consent screen: External, add scope `https://www.googleapis.com/auth/drive.file`, add staff as test users (or publish).
3. Credentials, OAuth client ID, Web application. Authorised JavaScript origins: your site, e.g. `https://your-site.netlify.app`.
4. Paste the Client ID into `GOOGLE_CLIENT_ID` near the top of `index.html` (or into the app under
   "Google Drive client database"), deploy, then press **Connect Google Drive** in the app.

The app only uses the `drive.file` permission, so it can see only the files and folders it creates.
Everyone who connects with the same Google account shares one database.
`tests/drive-tests.mjs` checks the folder and register logic against a stand-in Drive.
