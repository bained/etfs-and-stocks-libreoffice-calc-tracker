# LibreOffice ETF Data Tracker

[![YouTube: Setup & Usage Guide](https://img.shields.io/badge/YouTube-Setup%20%26%20Usage%20Guide-red)](https://www.youtube.com/watch?v=eK5rkf-VOP4)

Watch the video guide on how to use the pre-built `.ods` file:  
[**📺 LibreOffice ETF Tracker — Setup & Usage Guide**](https://www.youtube.com/watch?v=eK5rkf-VOP4)

---

This project provides an optimized system for tracking European ETFs in **LibreOffice Calc** using a **LibreOffice Basic** macro. Instead of relying on heavy real-time formulas that cause API rate limiting, data is fetched on demand via a button and stored in a local helper sheet (`etf_data`).

## File Structure

| File | Description |
|---|---|
| `ETFs_stat_v1.0.0.ods` | Pre-configured LibreOffice Calc spreadsheet — ready to use |
| `code.bas` | LibreOffice Basic macro source code |
| `README.md` | This documentation |

## Architecture

| Sheet | Purpose |
|---|---|
| `etf_data` | Local cache — row 1: tickers, row 2: current prices, rows 3+: historical close prices (newest → oldest) |
| `ETFs` | User interface with your portfolio, prices, and data |
| `settings` | Macro configuration |
| `etf_logs` | Auto-created log file (when logging is enabled) |

## Settings

The `settings` sheet controls macro behavior:

| Cell | Example | Description |
|---|---|---|
| **C2** | `Yes` / `No` | Enable or disable logging to `etf_logs` sheet |
| **C3** | `500` | Delay in milliseconds between Yahoo Finance requests |

## Data Flow

1. User clicks the **Update** button in the `ETFs` sheet
2. Macro reads tickers from row 1 of `etf_data`
3. For each ticker, it fetches 1 year of daily close prices from Yahoo Finance
4. Data is saved: row 2 = current price, rows 3+ = historical prices (newest at top)
5. Formulas in `ETFs` read from `etf_data` instantly — no network requests

## Requirements

- **LibreOffice 7.x+** (Windows)
- Active internet connection
- **MSXML2.XMLHTTP.6.0** (included with Windows)
- **Macros must be enabled** in LibreOffice settings (see below)

## Enabling Macros in LibreOffice

Before using the spreadsheet, you must allow macro execution:

1. Open the `.ods` file
2. Go to **Tools → Options → LibreOffice → Security**
3. Click **Macro Security** button
4. Set security level to **Medium** (recommended) or **Low**
5. Save and reopen the file
6. When prompted, click **Enable Macros**

> **Note:** On **Medium** security, you will be prompted each time you open the file. Select "Enable Macros" to allow the update button to work.

If macros are not enabled, clicking the **Update** button will do nothing.

## Notes

- Prices are reversed automatically — Yahoo returns oldest-first, the macro stores newest-first
- `GetSparklineData(ticker, period)` is available as a custom Calc function for charting (returns array of close prices per period: `"1w"`, `"1m"`, `"1y"`)
- Use with **Ctrl+Shift+Enter** (array formula) or `INDEX(GETSPARKLINEDATA(...); 1; n)`
- No caching — each button press fetches fresh data