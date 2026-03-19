# Rebalance Dashboard

Single-page dashboard for **MSCI** and **S&P** rebalance net-flow workbooks (`.xlsx`). Parses in the browser with [SheetJS](https://sheetjs.com/).

## Run locally

Open `index.html` in a browser, or serve the folder (e.g. `npx serve .`).

## S&P: two files

You can load **one** workbook or **two** S&P files: a larger **net-flows** export and a **share-change** file with `CHANGE` (e.g. Addition, Deletion). The app merges events and CUSIP by ticker onto the larger list.

## License

Use at your own discretion for internal workflows.
