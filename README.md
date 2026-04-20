# Valus — DCF Stock Screener & Model Generator

Valus is a Python tool that generates institutional-quality discounted cash flow (DCF) models for any publicly traded stock. It fetches historical financial data from Yahoo Finance, calculates WACC from scratch, projects free cash flows under Base/Bull/Bear scenarios, and exports a professionally formatted 10-sheet Excel workbook with income statement, balance sheet, cash flow statement, DCF valuation with sensitivity analysis, PP&E and working capital schedules, WACC build-up, scenario comparison, comps placeholder, and a one-page summary. It also includes a quick-screen mode that runs a simple DCF across the top 50 S&P 500 stocks.

## Installation

```bash
git clone <repo-url> && cd Valus
pip install -r requirements.txt
```

## Usage

### Web App (Streamlit)

```bash
streamlit run app.py
```

### CLI — Full DCF Model

```bash
python3 dcf.py AAPL
```

Generates `DCF_AAPL_2026-04-20.xlsx` with interactive scenario prompts.

### CLI — Quick Screener

```bash
python3 main.py
```

Screens the top 50 S&P 500 stocks and exports `valus_output.csv`.

## Screenshot

<!-- Add screenshot here -->
