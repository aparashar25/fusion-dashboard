"""The 18 charts in the workbook (2 on F0, 16 on F6), same data, series, chart types, axes and month ranges."""
from __future__ import annotations
import numpy as np
import plotly.graph_objects as go
import ui_kit as K

H = 380


def _fig(title, height=H):
    f = go.Figure()
    f.update_layout(title=dict(text=title, font=dict(size=14)), height=height)
    return f


def _bar(f, x, y, name, color, secondary=False):
    f.add_trace(go.Bar(x=x, y=y, name=name, marker_color=color, marker_line_width=0, yaxis="y2" if secondary else "y"))


def _line(f, x, y, name, color, secondary=False, dash=None, width=2.2):
    f.add_trace(go.Scatter(x=x, y=y, name=name, mode="lines", line=dict(color=color, width=width, dash=dash),
                           yaxis="y2" if secondary else "y"))


def _y2(f, fmt=None):
    f.update_layout(yaxis2=dict(overlaying="y", side="right", showgrid=False, tickformat=fmt))


def _ax(axis):
    return "yaxis" if axis == "y" else "yaxis2"


# hover and tick formats: $ whole dollars, MOIC 3 decimals with x, percentages 2 decimals
def _money(f, axis="y"):
    f.update_layout(**{_ax(axis): dict(tickprefix="$", hoverformat="$,.0f")})


def _moic(f, axis="y"):
    f.update_layout(**{_ax(axis): dict(ticksuffix="x", hoverformat=".3f")})


def _pct(f, axis="y"):
    f.update_layout(**{_ax(axis): dict(tickformat=".0%", hoverformat=".2%")})


# ---------------------------------------------------------------- section 1: returns overview (F5 rows 35-49)
def running_moic(res, months=120, la_right=False, title="Running MOIC by party (cumulative distributions ÷ cumulative equity)"):
    mv = res["f5"]["monthly"].iloc[:months]
    x = mv.index
    f = _fig(title)
    _line(f, x, mv["OCM running MOIC"], "OCM running MOIC", K.OCM)
    if la_right:   # F0 chart 1: Loan America on its own axis
        _line(f, x, mv["Total investor running MOIC"], "Total investor running MOIC", K.INK)
        _line(f, x, mv["Loan America running MOIC"], "Loan America running MOIC", K.LA, secondary=True)
        _y2(f)
    else:          # F6 chart: one axis, order OCM, Loan America, total
        _line(f, x, mv["Loan America running MOIC"], "Loan America running MOIC", K.LA)
        _line(f, x, mv["Total investor running MOIC"], "Total investor running MOIC", K.INK)
    _moic(f)
    if la_right:
        _moic(f, 'y2')
    return f


def cash_investors_vs_school(res):
    mv = res["f5"]["monthly"]
    f = _fig("Monthly cash received: investors (distributions) vs school (advance + share of collections), stacked")
    _bar(f, mv.index, mv["Investor distributions"], "Investor distributions", K.OCM)
    _bar(f, mv.index, mv["School cash received"], "School cash received (advance + share of collections)", K.SCHOOL)
    f.update_layout(barmode="relative")
    _money(f)
    return f


def school_cumulative(res):
    mv = res["f5"]["monthly"]
    f = _fig("School cumulative cash received (% of program face)")
    _line(f, mv.index, mv["School cumulative cash (% of program face)"], "School cumulative cash (% of program face)", K.SCHOOL)
    f.update_yaxes(tickformat=".0%")
    _pct(f)
    return f


# ---------------------------------------------------------------- section 2: funding & recycling (F5 rows 99-107, months 1-60)
def funding_mix(res):
    rc = res["f5"]["recycling"].iloc[:60]
    x = rc.index
    f = _fig("How each month's purchases are funded: recycled collections vs debt vs new equity; line = % funded by collections")
    _bar(f, x, rc["Funded by recycled collections ($)"], "Funded by recycled collections ($)", K.LA)
    _bar(f, x, rc["Funded by debt draws ($)"], "Funded by debt draws ($)", K.DEBT)
    _bar(f, x, rc["Funded by new equity ($)"], "Funded by new equity ($)", K.OCM)
    _line(f, x, rc["Share funded by recycled collections (%)"], "Share of purchases funded by recycled collections (%)", K.INK, secondary=True)
    f.update_layout(barmode="relative"); _y2(f, ".0%")
    _money(f)
    _pct(f, 'y2')
    return f


def cum_equity_vs_purchases(res):
    rc = res["f5"]["recycling"].iloc[:60]
    f = _fig("Cumulative purchases vs cumulative equity called (the gap is funded by recycled collections and debt)")
    _line(f, rc.index, rc["Cumulative equity called ($)"], "Cumulative equity called ($)", K.OCM)
    _line(f, rc.index, rc["Cumulative purchases ($)"], "Cumulative purchases ($)", K.SCHOOL)
    _money(f)
    return f


# ---------------------------------------------------------------- section 3: leverage (F5 rows 143-153, months 1-60)
def debt_facility(res):
    lv = res["f5"]["leverage"].iloc[:60]
    x = lv.index
    f = _fig("Debt facility: monthly draws and repayments (columns); debt balance vs borrowing capacity (lines)")
    _bar(f, x, lv["Draws ($)"], "Draws ($)", K.OCM)
    _bar(f, x, lv["Repayments ($)"], "Repayments ($)", K.DEBT)
    _line(f, x, lv["Ending debt ($)"], "Ending debt ($)", K.INK)
    _line(f, x, lv["Borrowing capacity ($)"], "Borrowing capacity ($)", K.FEES, dash="dot")
    f.update_layout(barmode="overlay")
    _money(f)
    return f


def leverage_investor_cash(res):
    lv = res["f5"]["leverage"].iloc[:60]
    f = _fig("Investors' cash: equity called vs distributions (no distributions during the debt paydown)")
    _bar(f, lv.index, lv["Equity called ($)"], "Equity called ($)", K.DEBT)
    _bar(f, lv.index, lv["Distributions to investors ($)"], "Distributions to investors ($)", K.OCM)
    f.update_layout(barmode="group")
    _money(f)
    return f


# ---------------------------------------------------------------- section 4: equity vs distributions by party (F4 rows 108-118, months 1-96)
def party_equity_vs_dist(res, party, months=96, with_line=True, title=None):
    w = res["spv"]["monthly"].iloc[:months]
    x = w.index
    if party == "SPV":
        f = _fig(title or "SPV (all investors): equity called vs distributions; line = cumulative net cash")
        _bar(f, x, -w.equity, "SPV (total investors): equity called (shown negative)", K.DEBT)
        _bar(f, x, w.dist, "SPV: distributions", K.OCM)
        if with_line:
            _line(f, x, w.spv_cum_net, "SPV: cumulative net cash flow", K.INK, secondary=True)
    elif party == "OCM":
        f = _fig("OCM: equity called vs distributions; line = cumulative net cash")
        _bar(f, x, -w.ocm_equity, "OCM: equity called (shown negative)", K.DEBT)
        _bar(f, x, w.ocm, "OCM: distributions", K.OCM)
        _line(f, x, w.ocm_cum_net, "OCM: cumulative net cash flow", K.INK, secondary=True)
    else:
        f = _fig("Loan America (sponsor): equity called vs distributions (pro rata + promote); line = cumulative net cash")
        _bar(f, x, -w.la_equity, "Loan America: equity called (shown negative)", K.DEBT)
        _bar(f, x, w.la_pro_rata, "Loan America: distributions, pro rata share", K.LA_LIGHT)
        _bar(f, x, w.la_promote, "Loan America: distributions, promote", K.LA)
        _line(f, x, w.la_cum_net, "Loan America: cumulative net cash flow", K.INK, secondary=True)
    f.update_layout(barmode="relative")
    if with_line:
        _y2(f)
    _money(f)
    if with_line:
        _money(f, 'y2')
    return f


# ---------------------------------------------------------------- section 5: promote (F4 rows 98-105, months 1-96)
def promote_monthly(res):
    w = res["spv"]["monthly"].iloc[:96]
    x = w.index
    f = _fig("Monthly distributions: OCM vs Loan America (pro rata + promote); line = Loan America's share of each month")
    _bar(f, x, w.ocm, "Distributions to OCM", K.OCM)
    _bar(f, x, w.la_pro_rata, "Loan America: pro rata share of distributions", K.LA_LIGHT)
    _bar(f, x, w.la_promote, "Loan America: promote above pro rata", K.LA)
    _line(f, x, w.la_share_month, "Loan America share of the month's distributions (%)", K.INK, secondary=True)
    f.update_layout(barmode="relative"); _y2(f, ".0%")
    _money(f)
    _pct(f, 'y2')
    return f


def promote_cumulative(res):
    w = res["spv"]["monthly"].iloc[:96]
    f = _fig("Cumulative distributions: OCM (left axis) vs Loan America (right axis)")
    _line(f, w.index, w.ocm_cum_dist, "Cumulative distributions to OCM", K.OCM)
    _line(f, w.index, w.la_cum_dist, "Cumulative distributions to Loan America", K.LA, secondary=True)
    _y2(f)
    _money(f)
    _money(f, 'y2')
    return f


# ---------------------------------------------------------------- section 6: advance-rate sensitivity (F2b)
def adv_sens(res, what="MOIC"):
    a = res["adv_sens"]
    x = a.loc["Blended advance"].values
    names = [c for c in res["ue"]["summary"].columns]
    f = _fig(f"Investor {what} per $1 of face vs advance rate (by term and blended)")
    cols = [K.OCM_LIGHT, K.LA_LIGHT, K.SCHOOL, K.DEBT, K.INK]
    for n, c in zip(names, cols):
        _line(f, x, a.loc[f"{what} {n}"].values, n if n != "Blended" else "Blended (mix-weighted)", c, width=3 if n == "Blended" else 2)
    f.update_xaxes(tickformat=".0%", title="Advance rate (blended)")
    f.update_layout(legend=dict(y=-0.22))
    if what == "IRR":
        f.update_yaxes(tickformat=".0%")
    (_moic(f) if what == 'MOIC' else _pct(f))
    f.update_xaxes(hoverformat='.2%')
    return f


# ---------------------------------------------------------------- section 7: split sensitivity (F4a)
def split_sens(res, what="MOIC"):
    s = res["split_sens"]
    x = s.loc["Investor %"].values
    f = _fig(f"{what} vs investor split: SPV and OCM (left axis), Loan America (right axis)")
    _line(f, x, s.loc[f"Total investors {what}"].values, f"SPV {what}", K.INK)
    _line(f, x, s.loc[f"OCM {what}"].values, f"OCM {what}", K.OCM)
    _line(f, x, s.loc[f"Loan America {what}"].values, f"Loan America {what}", K.LA, secondary=True)
    _y2(f, ".0%" if what == "IRR" else None)
    f.update_xaxes(tickformat=".0%", title="Investor % of net collections")
    f.update_layout(legend=dict(y=-0.22))
    if what == "IRR":
        f.update_yaxes(tickformat=".0%")
    (_moic(f) if what == 'MOIC' else _pct(f))
    (_moic(f, 'y2') if what == 'MOIC' else _pct(f, 'y2'))
    f.update_xaxes(hoverformat='.2%')
    return f
