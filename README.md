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
| Saved days (open earlier days without the main file) | yes, read back from the Google Sheet | yes ("Update online log") |
| "Send to Google" button (monthly Google Sheet) | yes | no (button hidden) |

Files are saved with a normal browser download on GitHub Pages.

## Optional: Google Sheets copy
`apps-script/FloorReports.gs` is a Google Apps Script that copies exported ACTUAL files from a Drive folder into a monthly Google Sheet and emails you any mismatch. Setup steps are at the top of the file. It has not been run against a real Google account yet, so test it on two or three files first.

## Libraries (loaded from CDN)
SheetJS 0.18.5, xlsx-js-style 1.2.0, jsPDF 2.5.1, JSZip 3.10.1.

## Send to Google (GitHub Pages version)
1. Paste `apps-script/FloorReports.gs` into a new Apps Script project (script.google.com), add the **Drive API** service (v3), set `FOLDER_NAME`, `EMAIL` at the top.
2. Deploy, New deployment, type **Web app**, Execute as **Me**, access **Anyone**. Copy the `/exec` link.
3. Put your `/exec` link in `GOOGLE_DEFAULT_URL` near the end of `index.html`. Nothing needs to be typed on the page or on other devices.
4. Load the main file and press **Send to Google**. Every day of the file is written to the monthly Google Sheet (`FLOOR REPORT OCT 2026`) and the page lists anything that does not match what was stored before.
The script has no password: anyone who has the `/exec` link can send data to your sheet, so do not share the link publicly.

## Saved days from Google (GitHub Pages version)
When the page opens it reads the saved days from your monthly Google Sheets (last 12 months) and lists them under the history tab / "Saved days". Press **Open** on a day to see its reports without loading the main file. **Refresh from Google** reloads the list. Load the main file and press **Send to Google** only when there is new or changed data; a day that is sent again replaces the old rows of that day. After pasting a new `FloorReports.gs`, redeploy it as a **New version**.

**Opening the page:** the last synced day is shown first, **view only** (no editing, no export, no sync), with a yellow bar asking you to open the latest file. Opening the latest file switches to the normal screen where you can check, edit and press **Send to Google**.

## Corrections and the corrected main file
- In the **Check entry** tab (after loading the main file) the Buyer, Order, Batch no, Category, Load and Unload cells can be typed over. Answers in **Needs a look** (category, unload time) count too.
- **Corrected main file (N changes)** makes a copy of your original file named `... - CORRECTED.xlsm`. Only the corrected cells are rewritten, so formats, formulas and macros stay as they are. Excel recalculates the formulas when it opens the copy. A list of every change is shown before it is saved. Check it, then replace the old file.

## Buyer list kept in Google
Buyers you add or correct in **Buyer list** (and the IN HOUSE / SUB answers) are saved to a small file `floor-report-settings.json` in the same Drive folder, so every phone and PC recognizes them. After pasting the new `FloorReports.gs`, redeploy it as a **New version**.

**If Excel cannot open the corrected copy:** press **Corrections list · .xlsx** instead. It lists each corrected cell (cell address, was, now) so you can type the changes into your own file.

## Logo, downloads, settings
- **Company logo:** put your logo file next to `index.html` and name it `logo.png`. It shows at the top left of the page (about 54 px high, smaller on phones). If the file is missing, nothing is shown.
- **Floating Downloads button** (bottom right): opens "Final reports · download options" with all the export buttons, Send to Google and the corrected main file.
- **Google link** is now a small "Settings · Google link" link at the very bottom of the page.

## Faster Send to Google
- Only days that are new or changed since the last send **from this device** are sent. If nothing changed, the page says so and offers "Send everything again".
- The script remembers the Drive folder / monthly file, reads and writes each sheet once per request (no more row-by-row deleting), and caches the saved-days list for 2 minutes. After pasting the new `FloorReports.gs`, redeploy it as a **New version**.

## Buyer not in the list
In the "not recognised" box (and the Problems window) the option **Wrong spelling: change my entries to…** puts the chosen buyer into the entries themselves (also into the corrected main file) and remembers that spelling for the next load. It no longer adds a keyword. You can undo it in **Buyer list → Spelling corrections kept**.

## Check entry
The `#` of a row turns green when it is a production batch (has a batch number and a quantity, not a sample / wash line, not in brackets).

## Reprocess
- Counted from batches whose status is **R/M, RM, R.M or Rematch** (not R/W or other statuses).
- Counted automatically **only for the newest day** (today). Older days are not recounted: they keep what was stored in Google (or 0 if nothing was stored), and sending never overwrites a stored value for them.
- A figure you type in the Month report (REPROCESS box) is **kept**: it is saved with your other fixes, shared between devices, and never recounted. Clear the box to go back to counting.
- After pasting the new `FloorReports.gs`, redeploy it as a **New version**.

## Added in this version
- **Actual file, Unit 1 + 2 combined · .xlsx**: one Sheet1 with every row of the day (UNIT column added), CATEGORY and BATCHMC tabs for each unit, and a TOTAL U1+2 tab. It is also inside the "Everything" zip.
- **Reprocess report · .xlsx**: every R/M batch of the loaded month up to the report day, both units (date, unit, MC, buyer, colour, batch, qty...), plus a BY DAY tab with kg per day. Also inside the "Everything" zip.
- The Category report · .xlsx export was removed (the .png stays).
- Every warning (Needs a look, Problems window, export stop, the Google mail) starts with **date, MC, buyer, colour, batch number**. The long-duration (18h+) warning now shows them too.
- **Check entry** fits one screen width: narrow fixed columns, the batch number column is just wide enough for 8 digits, Fabric and Check wrap. On a narrow phone it scrolls sideways.
- After pasting the new `FloorReports.gs`, redeploy it as a **New version** (the duplicate-batch message now names MC, buyer and colour).
