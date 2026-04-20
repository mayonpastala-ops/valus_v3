import yfinance as yf

# Top 50 S&P 500 by market cap (as of early 2026)
SP500_TOP50 = [
    "AAPL", "MSFT", "NVDA", "AMZN", "GOOGL", "META", "BRK-B", "LLY", "AVGO", "TSM",
    "JPM", "TSLA", "WMT", "UNH", "V", "XOM", "MA", "ORCL", "PG", "COST",
    "JNJ", "HD", "ABBV", "NFLX", "BAC", "CRM", "KO", "CVX", "MRK", "AMD",
    "PEP", "TMO", "LIN", "ACN", "ADBE", "MCD", "CSCO", "ABT", "WFC", "DHR",
    "TXN", "PM", "NEE", "AMGN", "QCOM", "INTU", "ISRG", "IBM", "GE", "CAT",
]


def fetch_financials(ticker):
    """Fetch all required financial data using yfinance."""
    stock = yf.Ticker(ticker)

    # Annual financials
    income = stock.financials
    cashflow = stock.cashflow
    balance = stock.balance_sheet

    if income.empty or cashflow.empty or balance.empty:
        raise ValueError(f"Missing financial data for {ticker}")

    # Most recent annual column
    inc = income.iloc[:, 0]
    cf = cashflow.iloc[:, 0]
    bal = balance.iloc[:, 0]

    revenue = _get_row(inc, ["Total Revenue"])
    net_income = _get_row(inc, ["Net Income"])
    fcf = _get_row(cf, ["Free Cash Flow"])
    total_debt = _get_row(bal, ["Total Debt"])
    cash = _get_row(bal, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"])
    shares = stock.info.get("sharesOutstanding", 0)

    # Calculate growth rate from revenue history
    growth_rate = _calc_growth(income)

    return {
        "revenue": revenue,
        "net_income": net_income,
        "free_cash_flow": fcf,
        "total_debt": total_debt or 0,
        "cash": cash or 0,
        "shares_outstanding": shares,
        "growth_rate": growth_rate,
    }


def fetch_price(ticker):
    """Fetch current share price."""
    stock = yf.Ticker(ticker)
    price = stock.info.get("currentPrice") or stock.info.get("regularMarketPrice")
    if not price:
        raise ValueError(f"No price data for {ticker}")
    return price


def _get_row(series, labels):
    """Try multiple row labels and return the first match."""
    for label in labels:
        if label in series.index:
            val = series[label]
            if val is not None:
                return float(val)
    return 0


def _calc_growth(income_df):
    """Calculate revenue CAGR from available annual data."""
    try:
        revenues = []
        for col in income_df.columns:
            for label in ["Total Revenue"]:
                if label in income_df.index:
                    val = income_df.loc[label, col]
                    if val and float(val) > 0:
                        revenues.append(float(val))

        if len(revenues) >= 2:
            # Columns are newest-first, so reverse
            oldest = revenues[-1]
            newest = revenues[0]
            years = len(revenues) - 1
            if oldest > 0:
                cagr = (newest / oldest) ** (1 / years) - 1
                return cagr
    except Exception:
        pass
    return 0.05  # Default 5%
