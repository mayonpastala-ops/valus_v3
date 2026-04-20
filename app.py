"""Valus — Streamlit web app with terminal UI."""

import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="VALUS",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═════════════════════════════════════════════════════════════════════════════
# TERMINAL CSS
# ═════════════════════════════════════════════════════════════════════════════

TERMINAL_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&display=swap');

:root {
    --bg:        #0a0a0a;
    --bg-raised: #111111;
    --green:     #00ff41;
    --green-dim: #00cc33;
    --amber:     #ffb000;
    --red:       #ff3333;
    --grey:      #666666;
    --grey-lt:   #888888;
    --white:     #cccccc;
    --border:    #222222;
    --font:      'JetBrains Mono', 'Courier New', monospace;
}

/* ── Global ─────────────────────────────────────────────────────────── */
html, body, .stApp, [data-testid="stAppViewContainer"],
[data-testid="stHeader"], [data-testid="stToolbar"] {
    background-color: var(--bg) !important;
    color: var(--green) !important;
    font-family: var(--font) !important;
}

[data-testid="stHeader"] { display: none !important; }
[data-testid="stToolbar"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }

/* ── Sidebar ────────────────────────────────────────────────────────── */
[data-testid="stSidebar"],
[data-testid="stSidebar"] > div:first-child,
section[data-testid="stSidebar"] {
    background-color: #050505 !important;
    border-right: 1px solid var(--border) !important;
}

[data-testid="stSidebar"] * {
    font-family: var(--font) !important;
    color: var(--grey-lt) !important;
}

[data-testid="stSidebar"] .stRadio label span {
    color: var(--green) !important;
}

[data-testid="stSidebar"] hr {
    border-color: var(--border) !important;
}

/* ── All text ───────────────────────────────────────────────────────── */
h1, h2, h3, h4, h5, h6, p, span, label, li, div, td, th {
    font-family: var(--font) !important;
}

h1 { color: var(--green) !important; font-size: 1.4rem !important; }
h2 { color: var(--green-dim) !important; font-size: 1.1rem !important; }
h3 { color: var(--green-dim) !important; font-size: 1.0rem !important; }
p, span, label { color: var(--white) !important; }

/* ── Input fields ───────────────────────────────────────────────────── */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stSelectbox > div > div > div {
    background-color: var(--bg-raised) !important;
    color: var(--green) !important;
    border: 1px solid var(--border) !important;
    border-radius: 0 !important;
    font-family: var(--font) !important;
    font-size: 0.85rem !important;
}

.stTextInput > label, .stNumberInput > label, .stSelectbox > label {
    color: var(--grey-lt) !important;
    font-family: var(--font) !important;
    font-size: 0.75rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.05em !important;
}

/* Selectbox dropdown */
[data-baseweb="select"] { background-color: var(--bg-raised) !important; }
[data-baseweb="popover"] {
    background-color: var(--bg-raised) !important;
    border: 1px solid var(--border) !important;
}
[data-baseweb="popover"] li {
    color: var(--green) !important;
    background-color: var(--bg-raised) !important;
}
[data-baseweb="popover"] li:hover {
    background-color: var(--border) !important;
}

/* Number input arrows */
.stNumberInput button {
    background-color: var(--bg-raised) !important;
    color: var(--grey) !important;
    border: 1px solid var(--border) !important;
    border-radius: 0 !important;
}

/* ── Buttons ────────────────────────────────────────────────────────── */
.stButton > button, .stDownloadButton > button {
    background-color: var(--bg) !important;
    color: var(--green) !important;
    border: 1px solid var(--green-dim) !important;
    border-radius: 0 !important;
    font-family: var(--font) !important;
    font-size: 0.85rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    transition: all 0.15s !important;
}

.stButton > button:hover, .stDownloadButton > button:hover {
    background-color: var(--green) !important;
    color: var(--bg) !important;
}

/* ── Progress bar ───────────────────────────────────────────────────── */
.stProgress > div > div > div {
    background-color: var(--green-dim) !important;
    border-radius: 0 !important;
}
.stProgress > div > div {
    background-color: var(--border) !important;
    border-radius: 0 !important;
}

/* ── Alerts ─────────────────────────────────────────────────────────── */
.stAlert {
    background-color: var(--bg-raised) !important;
    border: 1px solid var(--border) !important;
    border-radius: 0 !important;
    color: var(--white) !important;
}

[data-testid="stNotification"] {
    background-color: var(--bg-raised) !important;
    border-radius: 0 !important;
}

/* ── Dataframe ──────────────────────────────────────────────────────── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--border) !important;
    border-radius: 0 !important;
}

/* ── Spinner ────────────────────────────────────────────────────────── */
.stSpinner > div { color: var(--green) !important; }

/* ── Expander ───────────────────────────────────────────────────────── */
.streamlit-expanderHeader {
    background-color: var(--bg-raised) !important;
    color: var(--grey-lt) !important;
    border: 1px solid var(--border) !important;
    border-radius: 0 !important;
}

/* ── Radio buttons ──────────────────────────────────────────────────── */
.stRadio > div {
    background-color: transparent !important;
}
.stRadio label {
    color: var(--green) !important;
}

/* ── Dividers ───────────────────────────────────────────────────────── */
hr {
    border-color: var(--border) !important;
}

/* ── Custom classes ─────────────────────────────────────────────────── */
.term-box {
    background: var(--bg-raised);
    border: 1px solid var(--border);
    padding: 16px 20px;
    font-family: var(--font);
    font-size: 0.85rem;
    line-height: 1.6;
    margin-bottom: 12px;
}

.term-green { color: var(--green); }
.term-amber { color: var(--amber); }
.term-red   { color: var(--red); }
.term-grey  { color: var(--grey-lt); }
.term-dim   { color: var(--grey); }
.term-white { color: var(--white); }
.term-bold  { font-weight: 700; }

.term-header {
    color: var(--green);
    font-family: var(--font);
    font-size: 0.85rem;
    border-bottom: 1px solid var(--border);
    padding-bottom: 8px;
    margin-bottom: 12px;
}

.term-verdict-under { color: #00ff41; font-weight: 700; }
.term-verdict-fair  { color: #ffb000; font-weight: 700; }
.term-verdict-over  { color: #ff3333; font-weight: 700; }
</style>
"""

st.markdown(TERMINAL_CSS, unsafe_allow_html=True)

# ═════════════════════════════════════════════════════════════════════════════
# EXCHANGE MAPPING
# ═════════════════════════════════════════════════════════════════════════════

EXCHANGES = {
    "US": "",
    "India - NSE": ".NS",
    "India - BSE": ".BO",
    "UK": ".L",
    "Germany": ".DE",
    "Hong Kong": ".HK",
    "Australia": ".AX",
    "Canada": ".TO",
}


def _verdict_class(v):
    if v == "UNDERVALUED":
        return "term-verdict-under"
    elif v == "OVERVALUED":
        return "term-verdict-over"
    return "term-verdict-fair"


def _sign(x):
    return "+" if x > 0 else ""


# ═════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═════════════════════════════════════════════════════════════════════════════

st.sidebar.markdown(
    '<span class="term-green term-bold" style="font-size:1.1rem;">VALUS v1.0</span>',
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    '<span class="term-dim">DCF Model Generator</span>',
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")

page = st.sidebar.radio("NAV", ["dcf-model", "screener"], format_func=lambda x: f"> {x}")

st.sidebar.markdown("---")
st.sidebar.markdown(
    '<div class="term-box">'
    '<span class="term-grey">Generates institutional-quality<br>'
    'DCF models with 10 Excel sheets.<br><br>'
    'Data: Yahoo Finance (yfinance)<br><br>'
    'SHEETS:<br>'
    '  1. Income Statement<br>'
    '  2. Balance Sheet<br>'
    '  3. Cash Flow Statement<br>'
    '  4. DCF Valuation<br>'
    '  5. PPE & DA Schedule<br>'
    '  6. Working Capital<br>'
    '  7. WACC<br>'
    '  8. Scenarios<br>'
    '  9. Comps<br>'
    ' 10. Summary</span>'
    '</div>',
    unsafe_allow_html=True,
)

# ═════════════════════════════════════════════════════════════════════════════
# DCF MODEL PAGE
# ═════════════════════════════════════════════════════════════════════════════

if page == "dcf-model":
    st.markdown(
        '<div class="term-header">'
        '&#9608;&#9608; VALUS &mdash; DCF MODEL GENERATOR'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── Ticker + Exchange input ───────────────────────────────────────────
    c1, c2 = st.columns([2, 3])
    with c1:
        ticker_raw = st.text_input("TICKER", value="AAPL", max_chars=20).strip().upper()
    with c2:
        exchange = st.selectbox("EXCHANGE", list(EXCHANGES.keys()), index=0)

    suffix = EXCHANGES[exchange]
    ticker = f"{ticker_raw}{suffix}"

    st.markdown(
        f'<span class="term-dim">  yfinance ticker: </span>'
        f'<span class="term-green term-bold">{ticker}</span>',
        unsafe_allow_html=True,
    )

    # ── Scenario inputs ──────────────────────────────────────────────────
    st.markdown(
        '<div class="term-header" style="margin-top:20px;">'
        '> SCENARIO ASSUMPTIONS'
        '</div>',
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('<span class="term-green term-bold">BASE CASE</span>', unsafe_allow_html=True)
        base_growth = st.number_input("GROWTH %", value=8.0, step=0.5, key="bg", format="%.1f")
        base_tgr = st.number_input("TERMINAL %", value=2.5, step=0.5, key="bt", format="%.1f")
        base_wacc = st.number_input("WACC %", value=10.0, step=0.5, key="bw", format="%.1f")

    with col2:
        st.markdown('<span class="term-green term-bold">BULL CASE</span>', unsafe_allow_html=True)
        bull_growth = st.number_input("GROWTH %", value=12.0, step=0.5, key="ug", format="%.1f")
        bull_tgr = st.number_input("TERMINAL %", value=3.0, step=0.5, key="ut", format="%.1f")
        bull_wacc = st.number_input("WACC %", value=9.0, step=0.5, key="uw", format="%.1f")

    with col3:
        st.markdown('<span class="term-green term-bold">BEAR CASE</span>', unsafe_allow_html=True)
        bear_growth = st.number_input("GROWTH %", value=3.0, step=0.5, key="eg", format="%.1f")
        bear_tgr = st.number_input("TERMINAL %", value=2.0, step=0.5, key="et", format="%.1f")
        bear_wacc = st.number_input("WACC %", value=11.0, step=0.5, key="ew", format="%.1f")

    st.markdown("")

    # ── Generate ─────────────────────────────────────────────────────────
    if st.button("[ GENERATE DCF MODEL ]", use_container_width=True):
        if not ticker_raw:
            st.markdown(
                '<span class="term-red">ERROR: No ticker entered.</span>',
                unsafe_allow_html=True,
            )
        else:
            scenarios = {
                "Base": dict(rev_growth=base_growth / 100, wacc=base_wacc / 100,
                             tgr=base_tgr / 100),
                "Bull": dict(rev_growth=bull_growth / 100, wacc=bull_wacc / 100,
                             tgr=bull_tgr / 100),
                "Bear": dict(rev_growth=bear_growth / 100, wacc=bear_wacc / 100,
                             tgr=bear_tgr / 100),
            }

            with st.spinner(f"Fetching data for {ticker}..."):
                try:
                    from dcf import generate_dcf
                    out = generate_dcf(ticker, scenarios)
                except Exception as e:
                    st.markdown(
                        f'<div class="term-box">'
                        f'<span class="term-red">ERROR: {e}</span>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    st.stop()

            res = out["results"]
            price = out["current_price"]
            cur = out.get("currency_symbol", "$")

            # ── Results block ────────────────────────────────────────────
            lines = [
                f'<span class="term-green term-bold">ANALYSIS COMPLETE &mdash; {ticker}</span>',
                f'<span class="term-dim">{"=" * 52}</span>',
                f'<span class="term-white">  CURRENT PRICE:     {cur}{price:>12,.2f}</span>',
                f'<span class="term-white">  CURRENCY:          {out.get("currency", "USD")}</span>',
                f'<span class="term-dim">{"─" * 52}</span>',
            ]

            for name in ["Base", "Bull", "Bear"]:
                r = res[name]
                vc = _verdict_class(r["verdict"])
                lines.append(
                    f'<span class="term-white">  {name.upper():4s} INTRINSIC:  '
                    f'{cur}{r["ivps"]:>12,.2f}  '
                    f'({_sign(r["disc_pct"])}{r["disc_pct"]:.1%})  </span>'
                    f'<span class="{vc}">{r["verdict"]}</span>'
                )

            lines.extend([
                f'<span class="term-dim">{"─" * 52}</span>',
                f'<span class="term-grey">  WACC (Base): {res["Base"]["wacc"]:.2%}  |  '
                f'TGR: {res["Base"]["tgr"]:.1%}</span>',
                f'<span class="term-grey">  FILE: {out["filename"]}</span>',
                f'<span class="term-dim">{"=" * 52}</span>',
            ])

            st.markdown(
                '<div class="term-box">' + '<br>'.join(lines) + '</div>',
                unsafe_allow_html=True,
            )

            st.download_button(
                label=f"[ DOWNLOAD {out['filename']} ]",
                data=out["excel_bytes"],
                file_name=out["filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ═════════════════════════════════════════════════════════════════════════════
# SCREENER PAGE
# ═════════════════════════════════════════════════════════════════════════════

elif page == "screener":
    st.markdown(
        '<div class="term-header">'
        '&#9608;&#9608; VALUS &mdash; S&P 500 SCREENER'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<span class="term-grey">Quick DCF screen of top 50 S&P 500 stocks by market cap.</span>',
        unsafe_allow_html=True,
    )

    if st.button("[ RUN SCREENER ]", use_container_width=True):
        progress_bar = st.progress(0, text="Initialising...")

        def update_progress(i, total, ticker):
            progress_bar.progress(i / total, text=f"[{i}/{total}] {ticker}")

        with st.spinner("Running..."):
            from main import run_screener
            df, errors = run_screener(progress_callback=update_progress)

        progress_bar.empty()

        if df is not None and not df.empty:
            st.markdown(
                f'<span class="term-green term-bold">'
                f'COMPLETE: {len(df)} stocks analysed</span>',
                unsafe_allow_html=True,
            )

            st.dataframe(df, use_container_width=True, height=600)

            csv = df.to_csv(index_label="rank")
            st.download_button(
                label="[ DOWNLOAD CSV ]",
                data=csv,
                file_name="valus_output.csv",
                mime="text/csv",
            )
        else:
            st.markdown(
                '<span class="term-amber">WARNING: No results returned.</span>',
                unsafe_allow_html=True,
            )

        if errors:
            with st.expander(f"{len(errors)} ticker(s) skipped"):
                err_lines = [f"  {t}: {err}" for t, err in errors]
                st.markdown(
                    '<div class="term-box"><span class="term-red">'
                    + '<br>'.join(err_lines)
                    + '</span></div>',
                    unsafe_allow_html=True,
                )
