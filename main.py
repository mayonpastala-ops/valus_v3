#!/usr/bin/env python3
"""Valus — DCF Stock Screener for S&P 500"""

from data import SP500_TOP50, fetch_financials, fetch_price
from output import build_table, print_table, export_csv, verdict


def run_dcf(financials, discount_rate=0.10, terminal_growth=0.025, projection_years=5):
    """Simple DCF model for the screener. Returns intrinsic value per share."""
    fcf = financials["free_cash_flow"]
    growth = financials["growth_rate"]
    debt = financials["total_debt"]
    cash = financials["cash"]
    shares = financials["shares_outstanding"]

    if not fcf or fcf <= 0 or not shares or shares <= 0:
        return None

    growth = max(-0.10, min(0.30, growth))

    pv_fcfs = 0
    projected_fcf = fcf
    for year in range(1, projection_years + 1):
        projected_fcf *= (1 + growth)
        pv_fcfs += projected_fcf / (1 + discount_rate) ** year

    terminal_fcf = projected_fcf * (1 + terminal_growth)
    terminal_value = terminal_fcf / (discount_rate - terminal_growth)
    pv_terminal = terminal_value / (1 + discount_rate) ** projection_years

    enterprise_value = pv_fcfs + pv_terminal
    equity_value = enterprise_value - debt + cash
    return round(equity_value / shares, 2)


def run_screener(progress_callback=None):
    """
    Run the screener and return (DataFrame, errors_list).
    progress_callback: optional callable(i, total, ticker) for UI updates.
    """
    results = []
    errors = []

    for i, ticker in enumerate(SP500_TOP50, 1):
        if progress_callback:
            progress_callback(i, len(SP500_TOP50), ticker)
        try:
            financials = fetch_financials(ticker)
            price = fetch_price(ticker)
            intrinsic = run_dcf(financials)

            if intrinsic is None:
                errors.append((ticker, "Negative or zero FCF"))
                continue

            discount_pct = round(((intrinsic - price) / price) * 100, 1)
            results.append({
                "ticker": ticker,
                "price": round(price, 2),
                "intrinsic_value": intrinsic,
                "discount_pct": discount_pct,
                "verdict": verdict(discount_pct),
            })
        except Exception as e:
            errors.append((ticker, str(e)))

    df = build_table(results) if results else None
    return df, errors


def main():
    print("Valus — Fetching data for 50 stocks...\n")

    def cli_progress(i, total, ticker):
        print(f"[{i}/{total}] {ticker}...", end=" ", flush=True)

    df, errors = run_screener(progress_callback=cli_progress)

    if df is not None:
        print_table(df)
        export_csv(df)

    if errors:
        print(f"\n{len(errors)} ticker(s) skipped due to errors:")
        for t, err in errors:
            print(f"  {t}: {err}")


if __name__ == "__main__":
    main()
