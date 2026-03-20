# Rebalance Dashboard

Single-page dashboard for **MSCI** and **S&P** rebalance net-flow workbooks (`.xlsx`). Parses in the browser with [SheetJS](https://sheetjs.com/).

## Run locally

Open `index.html` in a browser, or serve the folder (e.g. `npx serve .`).

## Multiple files (same format per upload)

Upload **one or many** `.xlsx` files in a single drop/select. All files must be **all MSCI** or **all S&P** in that batch.

- **S&P (US and S&P/TSX):** Rows are unioned by **ticker** (later file wins duplicates). **Share-change** layouts supply `CHANGE` / CUSIP / index to rows from **net-flows** layouts where tickers match. **S&P/TSX** client summaries (Ticker + Sedol + Name, no CUSIP) are treated as S&P, not MSCI. Summary uses the richest single workbook; Top Names are merged and deduped by ticker.
- **MSCI:** Rows are merged by **SEDOL** (later file wins). Summary totals use the **max** per headline metric; industry sections are **summed** by name. Comparison is merged by SEDOL. Top Names are merged and deduped.

## License

Use at your own discretion for internal workflows.
