# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

A collection of Jupyter notebooks for personal finance analysis: investment returns tracking, stock portfolio holdings, loan amortization, and Singapore brokerage data.

## Key Files

| File | Purpose |
|------|---------|
| `sg_invest.ipynb` | Main investment tracker: stock holdings, PnL, dividends, IRR vs S&P500 |
| `bls_data_points.ipynb` | BLS economic data analysis |
| `AAII_survey_analysis.ipynb` | AAII investor sentiment survey analysis |
| `bmnr analysis.ipynb` | BMNR-related analysis |
| `generate_loan_schedule.py` | Generates loan amortization Excel workbook via xlsxwriter |
| `common.py` | Shared utilities: `get_historical_px`, `get_historical_close_px`, `get_px_for_date` |

## sg_invest.ipynb Architecture

### Data Sources (cell `f31f407d`)
```python
INVEST_FILE = r'G:\My Drive\invest\investment_returns.xlsx'   # deposits/withdrawals by brokerage
STOCK_FILE  = r'G:\My Drive\invest\stock_purchase.xlsx'       # sheet="stocks": ticker transactions
```

### Cell Structure
- **Cell `f31f407d`** — constants: file paths, `SHEETS_TO_SOURCE_MAP`
- **Cell `5be20916`** — all function definitions (large cell, ~620 lines)
- **Cell `2c63f870`** — loads `stocks = pd.read_excel(STOCK_FILE, sheet_name="stocks")`
- **Cell `05d43ffd`** — calls `stock_returns = calculate_stock_holdings(stocks, "market_value")`

### Key Functions in cell `5be20916`

**`investment_summary()`** — reads `INVEST_FILE` across multiple brokerage sheets, computes IRR, compares vs S&P500 (SPY × USDSGD=X, dividends reinvested). Prints total invested in S$.

**`calculate_stock_holdings(stock_transactions, sort_col)`** — core holdings engine:
- Infers currency per ticker: `.SI` → SGD, `.F` → EUR, else USD
- Downloads historical FX series (yfinance `USDSGD=X`, `EURSGD=X`) for the transaction date range
- Converts each transaction's native price to SGD using historical FX before passing to `Holdings`
- `Holdings` stores all monetary values in SGD
- Native `market_px` and `cost_basis_per_share` kept in native currency for display
- Calls `add_dividends_to_holdings(...)` then recalculates returns using `investment` (gross buys) as denominator
- Returns `find_sum_of_columns(out)` which appends a Total row

**`Holdings(market_px_sgd)`** — tracks a single ticker position in SGD:
- `investment` = gross buy cash outflow (never reduced by sells)
- `cost_basis` = remaining open position cost (reduces on sells)
- `returns` = `pnl / investment * 100` (uses gross investment, not exposure)

**`add_dividends_to_holdings(df, stock_transactions, years, fx_series_map, ticker_currency)`**:
- Withholding: SGD=1.0 (none), USD=0.70 (30% US WHT), EUR=0.7375 (26.375% German WHT)
- Converts dividends to SGD using historical FX rate on ex-dividend date
- Adds columns `dividends_YYYY` to the dataframe

**`find_sum_of_columns(df)`** — appends Total row; returns denominator is `investment` sum (not `cost_basis`).

### FX / Currency Constants
```python
FX_TICKER_MAP      = {'USD': 'USDSGD=X', 'EUR': 'EURSGD=X'}
CURRENCY_SYMBOL_MAP = {'SGD': 'S$', 'USD': '$', 'EUR': '€'}
WITHHOLDING_TAX    = {'SGD': 1.0, 'USD': 0.70, 'EUR': 0.7375}
```

### Important Design Decisions

**Returns denominator = `investment` (gross buys), not `cost_basis` (exposure)**
- Using `cost_basis` inflates returns for partially-sold positions
- Total row in `find_sum_of_columns`: `pnl.sum() / investment.sum()`
- Per-row recalc after dividends: `pnl / investment`

**Historical FX, not today's spot rate**
- All transaction prices converted to SGD using the FX rate on the transaction date
- Dividend amounts converted using FX rate on ex-dividend date
- Current market value uses today's spot rate (correct — it's current valuation)

**`total_value` = `market_value + realised_pnl`**
- `realised_pnl` already includes dividends and closed-position gains
- Do NOT subtract `total_dividends` again (they are already in `realised_pnl`)

**`investment_summary()` vs `calculate_stock_holdings()` total invested will not match exactly**
- `investment_summary` reads actual SGD cash deposits/withdrawals from brokerage statements
- `calculate_stock_holdings` reconstructs SGD cost from native-currency prices × yfinance FX rates
- Gap sources: broker FX spread vs yfinance mid-market, idle cash, reinvested proceeds

### Editing Notebooks
- `.ipynb` files must be edited with `NotebookEdit` tool (not `Edit`)
- For targeted string replacements across a large cell, use a Python JSON manipulation script via `Bash`
- After editing `.ipynb` on disk via Python/Bash, **close and reopen the notebook + Restart Kernel and Run All** to pick up changes (VS Code kernel caches old definitions)

### Known Tickers
- SGD stocks end with `.SI` (Singapore Exchange)
- EUR stocks end with `.F` (Frankfurt, e.g. `FJ2P.F`)
- US stocks: no suffix (e.g. `AAPL`, `MSFT`)
- `FJ2P.F` dividend fetch raises a period warning from yfinance — this is harmless, market price fetch works fine

## generate_loan_schedule.py

Standalone script. Generates an Excel loan amortization schedule with:
- `Inputs` sheet: loan amount, start date, rate change periods, prepayment events
- `Schedule` sheet: 420 rows of Excel formulas (EDATE, XLOOKUP, PMT, etc.)
- Run: `python generate_loan_schedule.py` → creates `loan_schedule_YYYYMMDD-HHMMSS.xlsx`
- Requires `xlsxwriter` (`pip install xlsxwriter`)
