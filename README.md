# RedirectValidator

Node.js tool using SheetJS to parse and export spreadsheets, as well as axios to validate redirect from / to, response status, etc.

## Install

Run `npm i` and replace `root url` and `someSpreadsheet.xlsx` with the root domain you want to test redirects for and a spreadsheet with `source` and `target` columns for redirects.

Use `JSON.safeStringify` for writing logs of responses for troubleshooting
