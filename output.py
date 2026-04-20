import pandas as pd


def verdict(discount_pct):
    """Classify based on discount/premium percentage."""
    if discount_pct > 20:
        return "UNDERVALUED"
    elif discount_pct < -20:
        return "OVERVALUED"
    return "FAIR"


def build_table(results):
    """Build a DataFrame from results list."""
    df = pd.DataFrame(results)
    df = df.sort_values("discount_pct", ascending=False).reset_index(drop=True)
    df.index += 1  # 1-based ranking
    return df


def print_table(df):
    """Print a clean terminal table."""
    print("\n" + "=" * 78)
    print(f"{'VALUS — DCF Stock Screener':^78}")
    print("=" * 78)
    print(
        f"{'#':>3}  {'Ticker':<7} {'Price':>9} {'Intrinsic':>10} {'Disc/Prem':>10}  {'Verdict'}"
    )
    print("-" * 78)

    for idx, row in df.iterrows():
        sign = "+" if row["discount_pct"] > 0 else ""
        print(
            f"{idx:>3}  {row['ticker']:<7} "
            f"${row['price']:>8.2f} "
            f"${row['intrinsic_value']:>9.2f} "
            f"{sign}{row['discount_pct']:>8.1f}%  "
            f"{row['verdict']}"
        )

    print("=" * 78)
    print(f"  {len(df)} stocks analyzed\n")


def export_csv(df, path="valus_output.csv"):
    """Export results to CSV."""
    df.to_csv(path, index_label="rank")
    print(f"Results exported to {path}")
