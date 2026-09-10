# Fleet Barcode Checker V2 — Google Sheets / GitHub Pages

## Connected endpoints
- GitHub Pages: https://nickparngong.github.io/QRchecker-Vehicle-registrationbook/
- Google Apps Script Web App: https://script.google.com/macros/s/AKfycbzAcQnJS1R5yg2WQXdn3SiSjPvN_zPR4_zakVyyNOBM6hU7B1qA9dysQ2GjLb1Gb2CSXw/exec
- Google Sheet: https://docs.google.com/spreadsheets/d/1rfd7D_NraAkv-0zhnKUfWYtGNg0OPsDH0B7QDcsUZ3A/edit

## Deploy to GitHub Pages
1. Extract this ZIP.
2. Replace the files in the GitHub repository with these files.
3. Commit / push to the branch used by GitHub Pages.
4. Open the GitHub Pages URL and hard-refresh (Ctrl+F5).

## Apps Script
The deployed Web App URL above is already placed in `config.js`.
The current Apps Script code uses the spreadsheet ID above and supports:
- getAll
- scan
- health

The current code has the optional APP_KEY disabled by leaving it at `CHANGE_THIS_KEY`. If you later enable a key in Code.gs, put the same key in config.js.

## Important
The Apps Script Web App must be deployed as a Web app and accessible to the intended users. Google documents that Web Apps are deployed from Deploy > New deployment and that the web app can execute as the deploying user. If access is restricted, GitHub Pages users will not be able to read/write the Sheet through the endpoint.
