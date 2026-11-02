# QR Checker V4 – Shared Google Sheets + Monthly Reset

## Included
- `index.html`: GitHub Pages frontend; reads/writes only through Apps Script.
- `Code.gs`: Apps Script backend connected to the existing shared spreadsheet.
- `README_DEPLOY_TH.txt`: deployment steps.

## Shared database
Spreadsheet ID: `1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A`
Web app endpoint currently embedded in index.html:
`https://script.google.com/macros/s/AKfycbxRV7OShwPWWOYeZg1StYI-TRgpxj50NMFXBFCN3pdft4cQ9mnR24keqJg2Cfho9a4v3A/exec`

## New V4 behavior
1. Both users read and write the same `VEHICLES` and `SCAN_HISTORY` sheets.
2. Every scan is saved centrally with user, area, date, time, month key, and round.
3. Dashboard totals are calculated for the current month and current round.
4. ADMIN can click “รีเซตรอบตรวจเดือนนี้”. It increments the round for the current month. It DOES NOT delete old scan history. After reset, each vehicle can be counted again.
5. Admin can add/edit vehicle master data and bulk import; these changes are written into the same `VEHICLES` sheet.
6. A new `SCAN_CYCLES` sheet is created automatically to store monthly round/reset history. `SCAN_HISTORY` gets `MonthKey` and `Round` columns appended if missing.

## Deployment steps
1. Open the Apps Script project connected to the existing spreadsheet (or create one at script.google.com and use the existing sheet ID in Code.gs).
2. Replace all contents of `Code.gs` with the supplied `Code.gs`, save.
3. Deploy > Manage deployments > Edit > New version > Deploy.
4. Execute as: Me. Who has access: Anyone (if allowed by company Workspace policy).
5. Copy the deployed `/exec` URL. If different from the endpoint above, replace `API_URL` in `index.html`.
6. In GitHub repository `QRchecker-Vehicle-registrationbook`, replace the root `index.html` with the supplied `index.html`, commit changes, and wait for Pages deployment.
7. Open the site in two browsers/devices and hard refresh (Ctrl+F5).

## Test plan
- On device A scan a known barcode. Check a new row in `SCAN_HISTORY`.
- Device B should see it after reload or within 12 seconds.
- Scan same barcode on device B in same round: it should be marked duplicate.
- ADMIN presses monthly reset. The round number increases and old history remains.
- Scan same barcode again after reset: it should be counted as found in the new round.
- ADMIN adds a new barcode/plate. Confirm it appears in `VEHICLES` and can be scanned by both devices.

## Important notes
- Do not delete `SCAN_HISTORY` to reset monthly counts. Use the in-app reset button.
- This backend assumes `VEHICLES` has a header row including Barcode (or QRCode), ทะเบียน, เดือนจดทะเบียน; `SCAN_HISTORY` has the existing columns described in the project.
- The existing front-end demo login credentials are not secure server-side authentication. Restrict access appropriately and implement server-side identity/authorization before broad production use.
- This package has been prepared but not deployed to GitHub or Apps Script by this assistant. Deployment and two-device verification are still required.
