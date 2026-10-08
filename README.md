# Floor Report Reader

Daily floor and production report tool for a dyeing / knit factory (Unit 1 and Unit 2).
One HTML file, no build step, no server.

## What it does
- Reads the **Main Production file** (.xlsm) or a daily **ACTUAL FLOOR REPORT** (.xlsx) in the browser. Nothing is uploaded anywhere.
- Pick a date, then get the CATEGORY-wise month report, machine-wise report, buyer-wise report (IN HOUSE / SUB CONTRACT), and a "Check entry" tab that lists problems row by row.
- Exports `ACTUAL FLOOR REPORT -DD MON YY (U-n).xlsx` and `FINAL PRODUCTION REPORT - DD MON (UNIT-n).xlsx`, plus buyer and machine pictures, or everything in one zip.
- "Buyer list" (top right) lets you add buyers, switch a buyer between IN HOUSE and SUB CONTRACT, and add keywords. Your changes are kept in the browser and the built-in list is never edited.

## Put it on GitHub Pages
1. Create a new repository and upload `index.html` (and the other files).
2. Settings, Pages, Source: **Deploy from a branch**, branch `main`, folder `/ (root)`.
3. Open `https://<your-user>.github.io/<repo>/`.

## What works where
| Feature | GitHub Pages / local file | claude.ai published page |
|---|---|---|
| Reading files, all reports, exports | yes | yes |
| Buyer list changes | saved in this browser only | shared |
| Online history (boss opens earlier days) | no | yes ("Update online log") |
| "Send to Google" button (monthly Google Sheet) | yes | no (button hidden) |

Files are saved with a normal browser download on GitHub Pages.

## Optional: Google Sheets copy
`apps-script/FloorReports.gs` is a Google Apps Script that copies exported ACTUAL files from a Drive folder into a monthly Google Sheet and emails you any mismatch. Setup steps are at the top of the file. It has not been run against a real Google account yet, so test it on two or three files first.

## Libraries (loaded from CDN)
SheetJS 0.18.5, xlsx-js-style 1.2.0, jsPDF 2.5.1, JSZip 3.10.1.

## Send to Google (GitHub Pages version)
1. Paste `apps-script/FloorReports.gs` into a new Apps Script project (script.google.com), add the **Drive API** service (v3), set `FOLDER_NAME`, `EMAIL` and your own `KEY` (a secret word) at the top.
2. Deploy, New deployment, type **Web app**, Execute as **Me**, access **Anyone**. Copy the `/exec` link.
3. Put your `/exec` link in `GOOGLE_DEFAULT_URL` near the end of `index.html` (the link alone is harmless, the script refuses requests without the KEY). Then on the page press **Google link**, type the KEY once and press **Copy setup link for other devices**. Open that setup link once on every other phone/PC and the link and KEY are saved there. Keep the setup link private, it contains the KEY.
4. Load the main file and press **Send to Google**. Every day of the file is written to the monthly Google Sheet (`FLOOR REPORT OCT 2026`) and the page lists anything that does not match what was stored before.
Do not put the KEY in the repository.
