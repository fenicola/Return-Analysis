"""
NVIDIA (NVDA) — Equity Analysis Dashboard
Carica i dati da nvda_analysis.xlsx (generato dal notebook NVIDIA.ipynb)

Run with: streamlit run nvda_dashboard.py
"""

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import statsmodels.api as sm
import streamlit as st
from plotly.subplots import make_subplots
from scipy import stats

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="NVDA · Equity Analysis",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── PALETTE ──────────────────────────────────────────────────────────────────
BG       = "#0b0f14"
SURFACE  = "#121821"
BORDER   = "#1f2a37"
TEXT     = "#e6edf3"
MUTED    = "#8b98a9"
FAINT    = "#5b6878"
NVDA_COL = "#76b900"   # NVIDIA green
MKT_COL  = "#8ea3c7"   # mercato: grigio-blu neutro
NEG      = "#ff6b6b"
AMBER    = "#f5b942"
GRID     = "rgba(139,152,169,0.10)"

TRADING_DAYS = 252
GITHUB_URL = "https://github.com/fenicola/Return-Analysis"

# ── CUSTOM CSS ────────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap');

html, body, [class*="css"], .stApp {{
    font-family: 'IBM Plex Sans', sans-serif;
    background: {BG};
    color: {TEXT};
}}
.block-container {{ padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1280px; }}
header[data-testid="stHeader"] {{ background: transparent; }}

/* Hero */
.hero {{
    display: flex; flex-wrap: wrap; justify-content: space-between; align-items: flex-end;
    gap: 20px; padding: 6px 0 22px; border-bottom: 1px solid {BORDER}; margin-bottom: 18px;
}}
.ticker {{
    font-family: 'IBM Plex Mono', monospace; font-size: 12px; letter-spacing: 2px;
    color: {NVDA_COL}; text-transform: uppercase; margin-bottom: 6px;
}}
.hero h1 {{
    font-size: clamp(30px, 5vw, 44px); font-weight: 700; line-height: 1.05;
    margin: 0; padding: 0; color: {TEXT}; letter-spacing: -0.5px;
}}
.hero .sub {{ color: {MUTED}; font-size: 14px; margin-top: 8px; }}
.hero-stat {{ text-align: right; }}
.hero-stat .v {{ font-family: 'IBM Plex Mono', monospace; font-size: 30px; color: {NVDA_COL}; }}
.hero-stat .l {{ color: {MUTED}; font-size: 12px; }}
@media (max-width: 900px) {{ .hero-stat {{ text-align: left; }} }}

/* Section headers */
.section {{ margin: 38px 0 6px; }}
.section .num {{
    font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: {NVDA_COL};
    letter-spacing: 2px;
}}
.section h2 {{ font-size: 21px; font-weight: 600; margin: 2px 0 2px; padding: 0; color: {TEXT}; }}
.section p {{ color: {MUTED}; font-size: 13.5px; margin: 0; max-width: 820px; }}

/* KPI grid */
.kpi-grid {{
    display: grid; grid-template-columns: repeat(auto-fit, minmax(165px, 1fr));
    gap: 12px; margin-top: 14px;
}}
.kpi-grid.four {{ grid-template-columns: repeat(4, 1fr); }}
@media (max-width: 900px) {{ .kpi-grid.four {{ grid-template-columns: repeat(2, 1fr); }} }}
.chart-title {{ font-size: 14px; font-weight: 600; color: {TEXT}; margin: 18px 0 -6px; }}
.kpi {{
    background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 10px;
    padding: 16px 18px;
}}
.kpi .l {{ font-size: 12px; color: {MUTED}; margin-bottom: 8px; }}
.kpi .v {{ font-family: 'IBM Plex Mono', monospace; font-size: 24px; color: {TEXT}; line-height: 1.1; }}
.kpi .m {{ font-family: 'IBM Plex Mono', monospace; font-size: 12px; color: {FAINT}; margin-top: 6px; }}
.kpi .m b {{ color: {MKT_COL}; font-weight: 500; }}

/* Legend chips */
.legend {{ display: flex; gap: 18px; font-size: 12.5px; color: {MUTED}; margin-top: 10px; }}
.dot {{ display: inline-block; width: 10px; height: 10px; border-radius: 2px; margin-right: 6px; vertical-align: -1px; }}

/* Regression table */
.reg-table {{ width: 100%; border-collapse: collapse; font-size: 13px; margin-bottom: 14px; }}
.reg-table th {{
    text-align: right; color: {MUTED}; font-weight: 500; font-size: 12px;
    padding: 8px 10px; border-bottom: 1px solid {BORDER};
}}
.reg-table th:first-child, .reg-table td:first-child {{ text-align: left; }}
.reg-table td {{
    text-align: right; padding: 9px 10px; border-bottom: 1px solid {BORDER};
    font-family: 'IBM Plex Mono', monospace; color: {TEXT};
}}
.reg-table td:first-child {{ font-family: 'IBM Plex Sans', sans-serif; color: {MUTED}; }}
.sig {{ color: {NVDA_COL}; }}

.note {{
    background: {SURFACE}; border: 1px solid {BORDER}; border-left: 3px solid {NVDA_COL};
    border-radius: 8px; padding: 14px 18px; font-size: 13.5px; color: {MUTED}; line-height: 1.65;
}}
.note b {{ color: {TEXT}; font-weight: 600; }}

.footer {{
    margin-top: 48px; padding-top: 18px; border-top: 1px solid {BORDER};
    color: {FAINT}; font-size: 12.5px; line-height: 1.7;
}}
.footer a {{ color: {MKT_COL}; }}
</style>
""", unsafe_allow_html=True)


# ── DATA LOADING DA EXCEL ─────────────────────────────────────────────────────
EXCEL_PATH = Path(__file__).parent / "nvda_analysis.xlsx"  # sempre nella cartella dello script


@st.cache_data(show_spinner=False)
def load_data(path: Path) -> pd.DataFrame:
    """
    Unisce ER_Analysis (excess return NVDA e mercato, in %) con il risk-free
    del foglio FamaFrench_3F, e ricostruisce i rendimenti totali giornalieri.
    """
    xls = pd.ExcelFile(path)
    er = pd.read_excel(xls, sheet_name="ER_Analysis", index_col=0, parse_dates=True)
    ff = pd.read_excel(xls, sheet_name="FamaFrench_3F", index_col=0, parse_dates=True)

    df = er.rename(columns={"Mkt-Rf": "MKT_ER", "ER": "NVDA_ER"})
    df = df.join(ff["RF"], how="left")
    df["RF"] = df["RF"].fillna(0.0)
    df = df.dropna(subset=["MKT_ER", "NVDA_ER"]).sort_index()

    # Rendimenti in decimale
    df["nvda_x"] = df["NVDA_ER"] / 100           # excess
    df["mkt_x"]  = df["MKT_ER"] / 100
    df["nvda"]   = (df["NVDA_ER"] + df["RF"]) / 100  # totale
    df["mkt"]    = (df["MKT_ER"] + df["RF"]) / 100
    return df


# ── METRICS ──────────────────────────────────────────────────────────────────
def drawdown(total: pd.Series) -> pd.Series:
    wealth = (1 + total).cumprod()
    return wealth / wealth.cummax() - 1


def risk_metrics(total: pd.Series, excess: pd.Series) -> dict:
    n = len(total)
    growth = (1 + total).prod()
    downside = np.sqrt((excess.clip(upper=0) ** 2).mean())
    var95 = np.percentile(total, 5)
    return dict(
        total_ret = growth - 1,
        cagr      = growth ** (TRADING_DAYS / n) - 1,
        vol       = total.std() * np.sqrt(TRADING_DAYS),
        sharpe    = excess.mean() / excess.std() * np.sqrt(TRADING_DAYS),
        sortino   = excess.mean() / downside * np.sqrt(TRADING_DAYS),
        max_dd    = drawdown(total).min(),
        var95     = var95,
        cvar95    = total[total <= var95].mean(),
    )


def run_capm(d: pd.DataFrame):
    x = sm.add_constant(d["MKT_ER"])
    return sm.OLS(d["NVDA_ER"], x).fit()


# ── PLOTLY THEME ──────────────────────────────────────────────────────────────
def style(fig: go.Figure, height: int = 380, **kwargs) -> go.Figure:
    layout = dict(
        height=height,
        plot_bgcolor=BG, paper_bgcolor=BG,
        font=dict(family="IBM Plex Sans, sans-serif", color=MUTED, size=12),
        margin=dict(l=8, r=8, t=40, b=8),
        hovermode="x unified",
        hoverlabel=dict(bgcolor=SURFACE, bordercolor=BORDER, font=dict(color=TEXT)),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, bgcolor="rgba(0,0,0,0)"),
    )
    fig.update_layout(**{**layout, **kwargs})
    fig.update_xaxes(gridcolor=GRID, zeroline=False, showline=False, automargin=True)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, showline=False, automargin=True)
    return fig


def show(fig: go.Figure):
    st.plotly_chart(fig, width="stretch", theme=None, config={"displayModeBar": False})


def section(num: str, title: str, text: str = ""):
    st.markdown(
        f'<div class="section"><div class="num">{num}</div><h2>{title}</h2>'
        + (f"<p>{text}</p>" if text else "") + "</div>",
        unsafe_allow_html=True,
    )


def pct(x: float, d: int = 1) -> str:
    return f"{x * 100:+,.{d}f}%" if x < 0 or x > 0 else f"{0:.{d}f}%"


def stars(p: float) -> str:
    return "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else ""


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN APP
# ═══════════════════════════════════════════════════════════════════════════════
try:
    full = load_data(EXCEL_PATH)
except FileNotFoundError:
    st.error(
        f"File **{EXCEL_PATH.name}** not found. Run `NVIDIA.ipynb` first and place "
        "`nvda_analysis.xlsx` next to this script."
    )
    st.stop()

if full.empty:
    st.error("The Excel file is empty or corrupted. Re-run the notebook.")
    st.stop()

# ── CONTROLS ─────────────────────────────────────────────────────────────────
PERIODS = {"1Y": 1, "3Y": 3, "5Y": 5, "Max": None}
WINDOWS = {"3M": 63, "6M": 126, "1Y": 252}

c_period, c_window, c_log = st.columns([2, 2, 1], vertical_alignment="bottom")
with c_period:
    period = st.segmented_control("Period", list(PERIODS), default="Max", required=True)
with c_window:
    win_label = st.segmented_control("Rolling window", list(WINDOWS), default="1Y", required=True)
with c_log:
    log_scale = st.toggle("Log scale", value=True)

years = PERIODS[period]
end = full.index.max()
start = full.index.min() if years is None else max(full.index.min(), end - pd.DateOffset(years=years))
df = full.loc[start:]
window = WINDOWS[win_label]

m_nvda = risk_metrics(df["nvda"], df["nvda_x"])
m_mkt  = risk_metrics(df["mkt"], df["mkt_x"])
model  = run_capm(df)
alpha, beta = model.params["const"], model.params["MKT_ER"]
p_alpha = model.pvalues["const"]
r2 = model.rsquared

# ── HERO ─────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="hero">
  <div>
    <div class="ticker">NASDAQ: NVDA · Equity analysis</div>
    <h1>NVIDIA Corporation</h1>
    <div class="sub">Performance, risk and CAPM regression · daily data
      {df.index.min():%b %Y} – {df.index.max():%b %Y} · market = Fama-French US market</div>
  </div>
  <div class="hero-stat">
    <div class="v">{pct(m_nvda['total_ret'], 0)}</div>
    <div class="l">total return over the period · market {pct(m_mkt['total_ret'], 0)}</div>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# 01 · PERFORMANCE OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
section("01", "Performance overview",
        "NVDA compared with the US equity market over the selected period. "
        "Returns include the risk-free rate; Sharpe and Sortino use excess returns.")

kpis = [
    ("CAGR",              pct(m_nvda["cagr"]),            pct(m_mkt["cagr"])),
    ("Volatility (ann.)", f"{m_nvda['vol']*100:.1f}%",    f"{m_mkt['vol']*100:.1f}%"),
    ("Sharpe ratio",      f"{m_nvda['sharpe']:.2f}",      f"{m_mkt['sharpe']:.2f}"),
    ("Sortino ratio",     f"{m_nvda['sortino']:.2f}",     f"{m_mkt['sortino']:.2f}"),
    ("Max drawdown",      pct(m_nvda["max_dd"]),          pct(m_mkt["max_dd"])),
    ("VaR 95% (1 day)",   pct(m_nvda["var95"], 2),        pct(m_mkt["var95"], 2)),
    ("CVaR 95% (1 day)",  pct(m_nvda["cvar95"], 2),       pct(m_mkt["cvar95"], 2)),
    ("CAPM beta",         f"{beta:.2f}",                  "1.00"),
]
st.markdown(
    '<div class="kpi-grid four">'
    + "".join(
        f'<div class="kpi"><div class="l">{l}</div><div class="v">{v}</div>'
        f'<div class="m">Market <b>{m}</b></div></div>'
        for l, v, m in kpis
    )
    + "</div>",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════════════════════════
# 02 · GROWTH & DRAWDOWN
# ══════════════════════════════════════════════════════════════════════════════
section("02", "Growth of $1 and drawdowns",
        "Value of $1 invested at the start of the period, and the percentage loss from the previous peak.")

wealth_n = (1 + df["nvda"]).cumprod()
wealth_m = (1 + df["mkt"]).cumprod()
dd_n, dd_m = drawdown(df["nvda"]) * 100, drawdown(df["mkt"]) * 100

fig = make_subplots(rows=2, cols=1, shared_xaxes=True, row_heights=[0.68, 0.32], vertical_spacing=0.04)
fig.add_trace(go.Scatter(x=wealth_n.index, y=wealth_n, name="NVDA",
                         line=dict(color=NVDA_COL, width=2),
                         hovertemplate="$%{y:,.2f}"), row=1, col=1)
fig.add_trace(go.Scatter(x=wealth_m.index, y=wealth_m, name="Market",
                         line=dict(color=MKT_COL, width=1.6),
                         hovertemplate="$%{y:,.2f}"), row=1, col=1)
fig.add_trace(go.Scatter(x=dd_n.index, y=dd_n, name="NVDA drawdown", showlegend=False,
                         line=dict(color=NEG, width=1), fill="tozeroy",
                         fillcolor="rgba(255,107,107,0.15)", hovertemplate="%{y:.1f}%"), row=2, col=1)
fig.add_trace(go.Scatter(x=dd_m.index, y=dd_m, name="Market drawdown", showlegend=False,
                         line=dict(color=MKT_COL, width=1, dash="dot"),
                         hovertemplate="%{y:.1f}%"), row=2, col=1)
style(fig, height=560)
fig.update_yaxes(type="log" if log_scale else "linear", tickprefix="$", row=1, col=1)
fig.update_yaxes(ticksuffix="%", row=2, col=1)
show(fig)

# ══════════════════════════════════════════════════════════════════════════════
# 03 · ANNUAL RETURNS
# ══════════════════════════════════════════════════════════════════════════════
section("03", "Calendar-year returns",
        "Compounded total return in each calendar year. Partial years are marked YTD or as a partial start.")

yr_n = (1 + df["nvda"]).groupby(df.index.year).prod() - 1
yr_m = (1 + df["mkt"]).groupby(df.index.year).prod() - 1
counts = df.groupby(df.index.year).size()
labels = []
for y in yr_n.index:
    if y == df.index.max().year and df.index.max().month < 12:
        labels.append(f"{y} YTD")
    elif y == df.index.min().year and df.index.min().month > 1:
        labels.append(f"{y}*")
    else:
        labels.append(str(y))

fig = go.Figure()
fig.add_trace(go.Bar(x=labels, y=yr_n * 100, name="NVDA", marker_color=NVDA_COL,
                     text=[f"{v*100:+.0f}%" for v in yr_n], textposition="outside",
                     textfont=dict(size=11, color=TEXT), hovertemplate="%{y:.1f}%"))
fig.add_trace(go.Bar(x=labels, y=yr_m * 100, name="Market", marker_color=MKT_COL,
                     hovertemplate="%{y:.1f}%"))
style(fig, height=380, barmode="group", bargap=0.25)
fig.update_xaxes(type="category")
fig.update_yaxes(ticksuffix="%")
show(fig)

# ══════════════════════════════════════════════════════════════════════════════
# 04 · ROLLING RISK
# ══════════════════════════════════════════════════════════════════════════════
section("04", f"Rolling risk · {win_label} window",
        "Annualised volatility and CAPM beta estimated on a moving window, to show how risk changes over time.")

# Calcolato su tutto il campione, poi tagliato: la finestra ha storia anche all'inizio del periodo
roll_vol_n = full["nvda"].rolling(window).std() * np.sqrt(TRADING_DAYS) * 100
roll_vol_m = full["mkt"].rolling(window).std() * np.sqrt(TRADING_DAYS) * 100
roll_beta  = full["NVDA_ER"].rolling(window).cov(full["MKT_ER"]) / full["MKT_ER"].rolling(window).var()

c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="chart-title">Annualised volatility</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_trace(go.Scatter(x=df.index, y=roll_vol_n.loc[start:], name="NVDA",
                             line=dict(color=NVDA_COL, width=1.8), hovertemplate="%{y:.1f}%"))
    fig.add_trace(go.Scatter(x=df.index, y=roll_vol_m.loc[start:], name="Market",
                             line=dict(color=MKT_COL, width=1.5), hovertemplate="%{y:.1f}%"))
    style(fig)
    fig.update_yaxes(ticksuffix="%")
    show(fig)
with c2:
    st.markdown('<div class="chart-title">CAPM beta</div>', unsafe_allow_html=True)
    fig = go.Figure()
    fig.add_hline(y=1, line=dict(color=FAINT, dash="dot", width=1))
    fig.add_hline(y=beta, line=dict(color=AMBER, dash="dash", width=1),
                  annotation_text=f"full-period β = {beta:.2f}", annotation_font_color=AMBER,
                  annotation_position="top left")
    fig.add_trace(go.Scatter(x=df.index, y=roll_beta.loc[start:], name="Rolling β",
                             line=dict(color=NVDA_COL, width=1.8), hovertemplate="%{y:.2f}"))
    style(fig, showlegend=False)
    show(fig)

# ══════════════════════════════════════════════════════════════════════════════
# 05 · DISTRIBUTION
# ══════════════════════════════════════════════════════════════════════════════
section("05", "Return distribution",
        "Daily excess returns compared with a normal distribution with the same mean and variance.")

er = df["NVDA_ER"].values
jb_stat, jb_p = stats.jarque_bera(er)
moments = [
    ("Mean (daily)",     f"{er.mean():+.3f}%"),
    ("Std dev (daily)",  f"{er.std(ddof=1):.3f}%"),
    ("Skewness",         f"{stats.skew(er):.2f}"),
    ("Excess kurtosis",  f"{stats.kurtosis(er):.2f}"),
    ("Jarque-Bera p",    "< 0.001" if jb_p < 0.001 else f"{jb_p:.3f}"),
]
st.markdown(
    '<div class="kpi-grid">'
    + "".join(f'<div class="kpi"><div class="l">{l}</div><div class="v">{v}</div></div>' for l, v in moments)
    + "</div>",
    unsafe_allow_html=True,
)

c3, c4 = st.columns(2)
with c3:
    st.markdown('<div class="chart-title">Histogram vs normal</div>', unsafe_allow_html=True)
    x_axis = np.linspace(er.min(), er.max(), 500)
    fig = go.Figure()
    fig.add_trace(go.Histogram(x=er, histnorm="probability density", nbinsx=120, name="Observed",
                               marker=dict(color="rgba(118,185,0,0.45)", line=dict(width=0))))
    fig.add_trace(go.Scatter(x=x_axis, y=stats.norm.pdf(x_axis, er.mean(), er.std()), name="Normal",
                             line=dict(color=MKT_COL, width=2, dash="dash")))
    fig.add_trace(go.Scatter(x=x_axis, y=stats.gaussian_kde(er)(x_axis), name="Kernel density",
                             line=dict(color=NVDA_COL, width=2)))
    style(fig, barmode="overlay", hovermode="closest")
    fig.update_xaxes(ticksuffix="%")
    show(fig)
with c4:
    st.markdown('<div class="chart-title">Q-Q plot</div>', unsafe_allow_html=True)
    (osm, osr), (slope, intercept, _) = stats.probplot(er, dist="norm")
    qq_x = np.array([osm.min(), osm.max()])
    fig = go.Figure()
    fig.add_trace(go.Scattergl(x=osm, y=osr, mode="markers", name="Observed quantiles",
                               marker=dict(color=NVDA_COL, size=4, opacity=0.6)))
    fig.add_trace(go.Scatter(x=qq_x, y=slope * qq_x + intercept, name="Normal line",
                             line=dict(color=MKT_COL, width=2, dash="dash")))
    style(fig, hovermode="closest")
    fig.update_xaxes(title="Theoretical quantiles")
    fig.update_yaxes(title="Observed quantiles (%)")
    show(fig)

tail_note = (
    "Points far from the line in both tails mean extreme days happen far more often than a normal "
    "distribution predicts (fat tails), so normal-based risk measures understate the real risk."
    if stats.kurtosis(er) > 1 else
    "Tails are close to normal over this period."
)
st.markdown(f'<div class="note">{tail_note}</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# 06 · CAPM REGRESSION
# ══════════════════════════════════════════════════════════════════════════════
section("06", "CAPM regression",
        "OLS of NVDA daily excess returns on market excess returns: "
        "R<sub>NVDA</sub> − R<sub>f</sub> = α + β (R<sub>M</sub> − R<sub>f</sub>) + ε")

col_sc, col_tab = st.columns([1.1, 1])
with col_sc:
    x_fit = np.linspace(df["MKT_ER"].min(), df["MKT_ER"].max(), 100)
    fig = go.Figure()
    fig.add_trace(go.Scattergl(x=df["MKT_ER"], y=df["NVDA_ER"], mode="markers", name="Daily observations",
                               marker=dict(color=NVDA_COL, size=4, opacity=0.35),
                               hovertemplate="Market %{x:.2f}%<br>NVDA %{y:.2f}%<extra></extra>"))
    fig.add_trace(go.Scatter(x=x_fit, y=alpha + beta * x_fit, name=f"Characteristic line (β = {beta:.2f})",
                             line=dict(color=AMBER, width=2.5), hoverinfo="skip"))
    style(fig, height=430, hovermode="closest")
    fig.update_xaxes(title="Market excess return (%)", zeroline=True, zerolinecolor=BORDER)
    fig.update_yaxes(title="NVDA excess return (%)", zeroline=True, zerolinecolor=BORDER)
    show(fig)

with col_tab:
    rows = [("Alpha (α, daily %)", "const"), ("Beta (β)", "MKT_ER")]
    body = "".join(
        f"<tr><td>{name}</td><td>{model.params[k]:.4f}</td><td>{model.bse[k]:.4f}</td>"
        f"<td>{model.tvalues[k]:.2f}</td><td>{model.pvalues[k]:.4f}</td>"
        f'<td class="sig">{stars(model.pvalues[k])}</td></tr>'
        for name, k in rows
    )
    alpha_ann = (1 + alpha / 100) ** TRADING_DAYS - 1
    st.markdown(f"""
    <table class="reg-table">
      <thead><tr><th>Parameter</th><th>Coef.</th><th>Std. err.</th><th>t</th><th>p-value</th><th></th></tr></thead>
      <tbody>{body}</tbody>
    </table>
    <table class="reg-table">
      <tbody>
        <tr><td>Alpha, annualised</td><td>{pct(alpha_ann)}</td></tr>
        <tr><td>R² / adjusted R²</td><td>{r2:.3f} / {model.rsquared_adj:.3f}</td></tr>
        <tr><td>Durbin-Watson</td><td>{sm.stats.stattools.durbin_watson(model.resid):.2f}</td></tr>
        <tr><td>Observations</td><td>{int(model.nobs):,}</td></tr>
      </tbody>
    </table>
    <div style="font-size:11.5px;color:{FAINT};margin:-4px 0 14px;">*** p&lt;0.001 · ** p&lt;0.01 · * p&lt;0.05</div>
    """, unsafe_allow_html=True)

    if p_alpha < 0.05:
        alpha_txt = (f"<b>Alpha</b> is {'positive' if alpha > 0 else 'negative'} and significant "
                     f"(p = {p_alpha:.4f}), about {pct(alpha_ann)} a year: NVDA "
                     f"{'beat' if alpha > 0 else 'lagged'} what its market exposure alone would explain.")
    else:
        alpha_txt = (f"<b>Alpha</b> is not statistically different from zero (p = {p_alpha:.2f}): "
                     "over this period NVDA's return is explained by its market exposure.")
    beta_txt = (f"<b>Beta</b> of {beta:.2f}: on average, a 1% market move goes with a "
                f"{beta:.1f}% move in NVDA, so the stock "
                f"{'amplifies' if beta > 1 else 'dampens'} market swings.")
    r2_txt = (f"<b>R²</b> of {r2*100:.0f}%: the market explains {r2*100:.0f}% of NVDA's daily variance; "
              f"the remaining {(1-r2)*100:.0f}% is company-specific risk.")
    st.markdown(f'<div class="note">{alpha_txt}<br><br>{beta_txt}<br><br>{r2_txt}</div>',
                unsafe_allow_html=True)

# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown(f"""
<div class="footer">
  <b>Data.</b> NVDA adjusted close prices from Yahoo Finance; market excess return (Mkt-RF) and risk-free rate
  from the Kenneth French Data Library. Pre-computed in <code>NVIDIA.ipynb</code> and loaded from
  <code>nvda_analysis.xlsx</code>.<br>
  <b>Definitions.</b> VaR and CVaR are historical (5th percentile of daily total returns and the mean beyond it).
  Rolling beta = rolling covariance / rolling market variance. Annualisation uses {TRADING_DAYS} trading days.<br>
  Code: <a href="{GITHUB_URL}" target="_blank">{GITHUB_URL.replace('https://', '')}</a>
</div>
""", unsafe_allow_html=True)
