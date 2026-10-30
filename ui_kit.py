"""Shared look and feel: light and dark palettes, CSS, number formats, Plotly template, grouped KPI cards, keyed widgets.

Call apply_mode(st.context.theme.type) once per run before drawing; it sets the palette for that mode."""
from __future__ import annotations
import math
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

PALETTES = {
    "light": dict(INK="#17202A", MUTED="#5B6675", RULE="#DCE1E8", PANEL="#F4F6F9", SURFACE="#FFFFFF", GRID="#EEF1F5",
                  OCM="#24527A", OCM_LIGHT="#4F7FA8", LA="#1A8C7E", LA_LIGHT="#62B3A4", SCHOOL="#C8962E", FEES="#7A8696",
                  DEBT="#8E2C3A", OK_C="#2E7D4F", WARN_C="#B7791F", BAD_C="#B3261E", HEAT0="#FFFFFF", HEAT1="#9CC3DE"),
    "dark": dict(INK="#E6EAF0", MUTED="#9AA4B2", RULE="#2C3440", PANEL="#161B22", SURFACE="#0F141A", GRID="#222A35",
                 OCM="#6FA8DC", OCM_LIGHT="#3E6E99", LA="#3CC2B0", LA_LIGHT="#2A8478", SCHOOL="#E3B355", FEES="#9AA6B6",
                 DEBT="#E07A88", OK_C="#5BC48A", WARN_C="#E2A84B", BAD_C="#F0706A", HEAT0="#0F141A", HEAT1="#2F5D86"),
}
MODE = "light"
globals().update(PALETTES["light"])
TIER_COLORS = []
CSS = ""


def apply_mode(mode: str | None):
    """Set the palette, CSS and Plotly template for 'light' or 'dark'."""
    global MODE, TIER_COLORS, CSS
    MODE = "dark" if mode == "dark" else "light"
    P = PALETTES[MODE]
    globals().update(P)
    TIER_COLORS = [P["OCM"], P["OCM_LIGHT"], P["LA"], P["LA_LIGHT"], P["SCHOOL"]]
    CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@400;500;600&display=swap');
html, body, [class*="css"], .stMarkdown, button, input, textarea {{
  font-family: 'Source Sans 3', system-ui, -apple-system, 'Segoe UI', sans-serif !important;
}}
.block-container {{ padding-top: 1.6rem; max-width: 1400px; }}
h1 {{ font-weight: 600; letter-spacing: -0.01em; font-size: 1.9rem; margin-bottom: 0.2rem; }}
h2 {{ font-weight: 600; font-size: 1.3rem; margin-top: 1.4rem; }}
h3 {{ font-weight: 600; font-size: 1.05rem; }}
.lede {{ color: {P['MUTED']}; font-size: 0.98rem; margin-bottom: 1rem; max-width: 78ch; }}
.rail {{ display:flex; border:1px solid {P['RULE']}; border-radius:6px; overflow:hidden; margin: 0.4rem 0 1.2rem 0; }}
.rail div {{ flex:1; padding:0.55rem 0.9rem; font-size:0.9rem; color:{P['INK']}; border-right:1px solid {P['RULE']}; background:{P['SURFACE']}; }}
.rail div:last-child {{ border-right:none; }}
.rail b {{ font-weight:600; }}
.dot {{ display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:7px; vertical-align:1px; }}
.kpi {{ border-top: 3px solid var(--c); padding: 0.55rem 0.1rem 0.4rem 0.1rem; }}
.kpi .v {{ font-size: 1.85rem; font-weight: 600; color: {P['INK']}; font-variant-numeric: tabular-nums; line-height:1.15; }}
.kpi .v.neg, .neg {{ color: {P['BAD_C']} !important; }}
.kpi .l {{ font-size: 0.86rem; color: {P['MUTED']}; margin-top: 0.1rem; }}
.kpi .s {{ font-size: 0.8rem; color: {P['MUTED']}; }}
.note {{ font-size: 0.85rem; color: {P['MUTED']}; }}
/* party cards: the one place with large numbers */
.party {{ border:1px solid {P['RULE']}; border-top:4px solid var(--c); border-radius:8px; padding:0.9rem 1.1rem 0.8rem 1.1rem;
          background:{P['SURFACE']}; height:100%; }}
.party .name {{ font-size:0.95rem; font-weight:600; color:{P['INK']}; }}
.party .share {{ font-size:0.8rem; color:{P['MUTED']}; margin-bottom:0.35rem; }}
.party .big {{ display:flex; gap:1.6rem; align-items:baseline; }}
.party .big div {{ font-variant-numeric: tabular-nums; }}
.party .big .n {{ font-size:2.1rem; font-weight:600; color:{P['INK']}; line-height:1.1; }}
.party .big .u {{ font-size:0.8rem; color:{P['MUTED']}; }}
.party .flow {{ font-size:0.85rem; color:{P['MUTED']}; margin-top:0.45rem; font-variant-numeric: tabular-nums; }}
/* grouped figures: label left, value right */
.group {{ border:1px solid {P['RULE']}; border-radius:8px; padding:0.7rem 1rem 0.5rem 1rem; background:{P['SURFACE']}; height:100%; }}
.group .t {{ font-size:0.95rem; font-weight:600; color:{P['INK']}; padding-bottom:0.45rem; border-bottom:2px solid var(--c); margin-bottom:0.2rem; }}
.group .row {{ display:flex; justify-content:space-between; gap:0.8rem; padding:0.38rem 0; border-bottom:1px solid {P['RULE']}; }}
.group .row:last-child {{ border-bottom:none; }}
.group .row .k {{ font-size:0.88rem; color:{P['INK']}; }}
.group .row .k small {{ display:block; font-size:0.76rem; color:{P['MUTED']}; }}
.group .row .val {{ font-size:0.98rem; font-weight:600; color:{P['INK']}; font-variant-numeric: tabular-nums; white-space:nowrap; }}
.status {{ font-weight:600; }}
[data-testid="stDataFrame"] {{ font-variant-numeric: tabular-nums; }}
</style>
"""
    pio.templates["fusion"] = go.layout.Template(layout=go.Layout(
        font=dict(family="Source Sans 3, system-ui, sans-serif", size=13, color=P["INK"]),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(gridcolor=P["GRID"], zerolinecolor=P["RULE"], linecolor=P["RULE"]),
        yaxis=dict(gridcolor=P["GRID"], zerolinecolor=P["MUTED"], linecolor=P["RULE"]),
        legend=dict(orientation="h", yanchor="top", y=-0.12, x=0, bgcolor="rgba(0,0,0,0)"),
        title=dict(x=0, xanchor="left", font=dict(size=15, color=P["INK"])),
        margin=dict(l=10, r=10, t=46, b=10), hovermode="x unified",
        colorway=[P["OCM"], P["LA"], P["SCHOOL"], P["FEES"], P["DEBT"], P["OCM_LIGHT"], P["LA_LIGHT"]],
    ))
    pio.templates.default = "fusion"


apply_mode("light")


# ---- formats
def _bad(x):
    return x is None or (isinstance(x, float) and (math.isnan(x) or math.isinf(x)))


def fx(x, d=2):
    return "n.a." if _bad(x) else f"{x:,.{d}f}x"


def fp(x, d=1):
    return "n.a." if _bad(x) else f"{x * 100:,.{d}f}%"


def fm(x, d=1):
    if _bad(x):
        return "n.a."
    a = abs(x)
    s = "-" if x < 0 else ""
    if a >= 1e9:
        return f"{s}${a / 1e9:,.{d}f}B"
    if a >= 1e6:
        return f"{s}${a / 1e6:,.{d}f}M"
    if a >= 1e3:
        return f"{s}${a / 1e3:,.{d}f}K"
    return f"{s}${a:,.0f}"


def fc(x, d=4):
    return "n.a." if _bad(x) else f"{x:,.{d}f}"


def kpi(value: str, label: str, sub: str = "", color: str | None = None, negative: bool = False):
    st.markdown(f"<div class='kpi' style='--c:{color or INK}'><div class='v{' neg' if negative else ''}'>{value}</div>"
                f"<div class='l'>{label}</div><div class='s'>{sub}</div></div>", unsafe_allow_html=True)


def party_card(name, share, moic, irr, flow, color, neg=False):
    """Headline card for one party: MOIC and IRR large, capital share and cash flow small."""
    cls = " neg" if neg else ""
    st.markdown(
        f"<div class='party' style='--c:{color}'><div class='name'>{name}</div><div class='share'>{share}</div>"
        f"<div class='big'><div><div class='n{cls}'>{moic}</div><div class='u'>MOIC</div></div>"
        f"<div><div class='n{cls}'>{irr}</div><div class='u'>IRR, annualised</div></div></div>"
        f"<div class='flow'>{flow}</div></div>", unsafe_allow_html=True)


def group_card(title, rows, color=None):
    """A titled group of figures. rows = [(label, value, sub, negative)]; sub and negative optional."""
    html = [f"<div class='group' style='--c:{color or INK}'><div class='t'>{title}</div>"]
    for r in rows:
        label, value = r[0], r[1]
        sub = r[2] if len(r) > 2 and r[2] else ""
        neg = r[3] if len(r) > 3 else False
        html.append(f"<div class='row'><div class='k'>{label}{f'<small>{sub}</small>' if sub else ''}</div>"
                    f"<div class='val{' neg' if neg else ''}'>{value}</div></div>")
    html.append("</div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def status_color(s):
    return {"OK": OK_C, "WARNING": WARN_C}.get(s, BAD_C)


def status_rail(res):
    d = res["derived"]
    chk = d["checks"]
    inp_s, mod_s = res["input_status"], res["model_status"]
    fails = ", ".join(chk["input_failures"]) or ("last draw month is before the purchase period ends"
                                                 if chk["last_draw_warning"] != "OK" else "all 9 input checks pass")
    inp = res["inputs"]
    sett = (f"{fp(d['tiers'][0][0], 0) if d['tiers'] else 'n.a.'} to investors, {fp(d['F71'])} advance, "
            f"{fm(inp.purchase_amount, 0)} over {inp.purchase_period} months, recycling {'on' if inp.recycling else 'off'}, "
            f"leverage {'on' if inp.leverage else 'off'}")
    st.markdown(
        f"<div class='rail'>"
        f"<div><span class='dot' style='background:{status_color(inp_s)}'></span><b>Inputs {inp_s}</b> &nbsp;{fails}</div>"
        f"<div><span class='dot' style='background:{status_color(mod_s)}'></span><b>Model {mod_s}</b> &nbsp;"
        f"{'every identity and tie-out holds' if mod_s == 'OK' else 'see the All checks page'}</div>"
        f"<div style='flex:1.6'>{sett}</div></div>", unsafe_allow_html=True)


def lede(text):
    st.markdown(f"<div class='lede'>{text}</div>", unsafe_allow_html=True)


def note(text):
    st.markdown(f"<div class='note'>{text}</div>", unsafe_allow_html=True)


def table(df: pd.DataFrame, formats: dict | None = None, height=None, row_formats: dict | None = None):
    show = df.copy()
    if row_formats:
        show = show.astype(object)
        for r, f in row_formats.items():
            if r in show.index:
                show.loc[r] = [f(v) if isinstance(v, (int, float, np.floating)) else v for v in df.loc[r]]
    if formats:
        show = show.astype(object)
        for c, f in formats.items():
            if c in show.columns:
                show[c] = [f(v) if isinstance(v, (int, float, np.floating)) else v for v in df[c]]
    if height:
        st.dataframe(show, width="stretch", height=height)
    else:
        st.dataframe(show, width="stretch")


# ---- keyed widgets that read from / write to the Inputs object
def _init(key, value):
    if key not in st.session_state:
        st.session_state[key] = value


def w_num(label, key, value, **kw):
    _init(key, value)
    return st.number_input(label, key=key, **kw)


def w_pct(label, key, value, step=0.5, **kw):
    _init(key, round(value * 100, 6))
    return st.number_input(label, key=key, step=step, format="%.2f", **kw) / 100


def w_toggle(label, key, value, **kw):
    _init(key, value)
    return st.toggle(label, key=key, **kw)


def w_select(label, key, value, options, **kw):
    _init(key, value if value in options else options[0])
    return st.selectbox(label, options, key=key, **kw)


def clear_widget_state():
    for k in list(st.session_state.keys()):
        if isinstance(k, str) and k.startswith("w_"):
            del st.session_state[k]


def line_fig(x, series: dict, title="", yfmt=None, height=340, colors=None, dash=None):
    fig = go.Figure()
    for name, y in series.items():
        fig.add_trace(go.Scatter(x=x, y=y, name=name, mode="lines",
                                 line=dict(width=2.2, color=(colors or {}).get(name), dash=(dash or {}).get(name))))
    fig.update_layout(title=dict(text=title, font=dict(size=15)), height=height)
    if yfmt:
        fig.update_yaxes(tickformat=yfmt)
    return fig


def bar_fig(x, series: dict, title="", yfmt=None, height=340, colors=None, barmode="relative"):
    fig = go.Figure()
    for name, y in series.items():
        fig.add_trace(go.Bar(x=x, y=y, name=name, marker_color=(colors or {}).get(name), marker_line_width=0))
    fig.update_layout(title=dict(text=title, font=dict(size=15)), height=height, barmode=barmode, bargap=0.15)
    if yfmt:
        fig.update_yaxes(tickformat=yfmt)
    return fig
