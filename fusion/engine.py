"""Fusion deal engine. Each function mirrors one tab of LoanAmerica_Model v4.

Row numbers in comments refer to the Excel workbook so every line can be traced:
F1 = Fusion Assumptions, F2 = Unit Economics, F2a = Asset Returns,
F2b / F4a = sensitivities, F3 = Pool Cash Flows, F4 = SPV Summary.
"""
from __future__ import annotations
import datetime as dt
import math
import numpy as np
import pandas as pd

from .inputs import Inputs, AGES, MONTHS, PLAN_MONTHS
from .irr import irr_annual, npv_excel

BIG = 1e12        # Excel's stand-in for "never completes" when a tier's investor % is 0
TOL = 1e-9


# =============================================================================== F1
def derive(inp: Inputs) -> dict:
    """Formula cells on F1: payments, effective curves, fees, tiers, purchase plan, checks."""
    d: dict = {}
    T = [int(t) for t in inp.terms]
    d["T"] = T
    d["r"] = [a / 12 for a in inp.apr]
    d["pmt"] = [_pmt(a / 12, t) for a, t in zip(inp.apr, T)]                     # F1 row 33
    mix = np.array(inp.mix, float)
    d["mix"] = mix
    d["wa_apr"] = float(np.dot(inp.apr, mix))                                    # F29
    d["wa_term"] = float(np.dot(T, mix))                                         # F34
    sizes = np.array(inp.loan_size, float)
    d["avg_loan"] = float(1 / np.sum(mix / sizes)) if np.all(sizes > 0) else 0.0  # F36
    lm = inp.loss_mult if inp.losses else 0.0                                    # B44
    pm = inp.prepay_mult if inp.prepay else 0.0                                  # B45
    d["lm_eff"], d["pm_eff"] = lm, pm
    fam_idx = [inp.curve_names.index(f) if f in inp.curve_names else 0 for f in inp.curve_family]  # row 32
    d["fam_idx"] = fam_idx
    loss, smm = [], []
    for k in range(4):
        base_l = np.array(inp.loss_curves[fam_idx[k]], float)
        base_s = np.array(inp.smm_curves[fam_idx[k]], float)
        ages = np.arange(AGES)
        loss.append(np.minimum(1.0, lm * base_l[np.minimum(ages, T[k])]))       # rows 60-63
        smm.append(np.where(ages <= T[k], pm * base_s, 0.0))                     # rows 64-67
    d["loss"], d["smm"] = loss, smm
    # payment-stop share rows: short row for terms <= 36, long row otherwise
    d["stop_share"] = [inp.stop_share_short if T[k] <= 36 else inp.stop_share_long for k in range(4)]

    # ---- fee table (rows 78-97)
    fee_ok = []
    for f in inp.fees:
        if not f.active:
            fee_ok.append("")
            continue
        type_ok = f.type in ("Upfront", "Across life")
        pct_ok = (f.method == "%" and isinstance(f.rate, (int, float)) and f.rate >= 0 and
                  (f.basis in ("Net paid", "Face") if f.type == "Upfront"
                   else f.basis in ("Collections", "Beginning balance", "Ending balance", "Average balance")))
        fix_ok = f.method == "Fixed $" and isinstance(f.fixed, (int, float)) and f.fixed >= 0
        fee_ok.append("OK" if type_ok and (pct_ok or fix_ok) else "CHECK")
    d["fee_ok"] = fee_ok
    adv = np.array(inp.advance, float)
    up = np.zeros(4)
    for f, ok in zip(inp.fees, fee_ok):
        if ok == "OK" and f.type == "Upfront":
            if f.method == "%":
                up += f.rate * (1.0 if f.basis == "Face" else adv)
            elif f.method == "Fixed $":
                up += f.fixed / sizes
    d["adv"], d["upfront"] = adv, up                                            # rows 71, 72
    d["outlay"] = adv + up                                                      # row 73
    d["target_outlay"] = adv if inp.outlay_basis == "Price only" else adv + up  # row 74
    d["F71"], d["F72"] = float(np.dot(adv, mix)), float(np.dot(up, mix))
    d["F73"] = d["F71"] + d["F72"]

    def coef(basis=None, fixed=False):
        s = 0.0
        for f, ok in zip(inp.fees, fee_ok):
            if ok == "OK" and f.type == "Across life":
                if fixed and f.method == "Fixed $":
                    s += f.fixed
                elif not fixed and f.method == "%" and f.basis == basis:
                    s += f.rate
        return s
    d["c_coll"], d["c_beg"], d["c_end"], d["c_avg"] = (coef("Collections"), coef("Beginning balance"),
                                                         coef("Ending balance"), coef("Average balance"))
    d["c_fix"] = coef(fixed=True)                                              # rows 93-97

    # ---- target tiers (rows 101-114)
    tier_ok = []
    for t in inp.tiers:
        if not t.active:
            tier_ok.append("")
            continue
        tgt = t.moic_target if inp.target_mode == "MOIC" else t.irr_target
        ok = (isinstance(t.investor_pct, (int, float)) and 0 <= t.investor_pct <= 1 and
              isinstance(tgt, (int, float)) and tgt > 0)
        tier_ok.append("OK" if ok else "CHECK")
    active = [(t.investor_pct, t.moic_target if inp.target_mode == "MOIC" else t.irr_target)
              for t, ok in zip(inp.tiers, tier_ok) if ok == "OK"]
    d["tier_ok"], d["tiers"] = tier_ok, active
    n = len(active)
    d["n_tiers"] = n
    d["after_pct"] = 0.0 if n == 0 else (active[-1][0] if inp.after_last_tier == "Continue split" else 0.0)  # B112

    # ---- purchase plan (rows 117-126)
    P, R = int(inp.purchase_period), int(inp.ramp_months)
    ramp = np.array(inp.ramp_split[:PLAN_MONTHS], float)
    ramp_sum = float(ramp[:R].sum()) if R > 0 else 0.0
    split = np.zeros(PLAN_MONTHS)
    for m in range(1, PLAN_MONTHS + 1):
        if m <= R:
            split[m - 1] = ramp[m - 1]
        elif m <= P and P > R:
            split[m - 1] = (1 - ramp_sum) / (P - R)
    cash = inp.purchase_amount * split
    face = cash / d["F71"] if d["F71"] else np.zeros(PLAN_MONTHS)
    d["plan_split"], d["plan_cash"], d["plan_face"] = split, cash, face
    d["total_face"] = float(face.sum())                                         # B120
    d["loans_by_term"] = [d["total_face"] * mix[k] / sizes[k] if sizes[k] else 0 for k in range(4)]  # row 37
    d["facility_max"] = inp.ltv * inp.purchase_amount                          # B139
    fm = dt.date.fromisoformat(str(inp.first_month)[:10])
    d["dates"] = [_eomonth(fm, m - 1) for m in range(1, MONTHS + 1)]

    # ---- F1 checks
    chk = {}
    chk["term_mix"] = "OK" if abs(mix.sum() - 1) < 1e-6 else "CHECK"                       # B35
    curves = np.vstack(loss)
    chk["loss_curves"] = "OK" if not np.any(curves[:, 1:] < curves[:, :-1]) else "CHECK"   # B68
    chk["fee_table"] = "OK" if (sizes.min() > 0 and "CHECK" not in fee_ok) else "CHECK"     # B90
    chk["target_tiers"] = "OK" if (n > 0 and "CHECK" not in tier_ok) else "CHECK"           # B114
    chk["purchase_split"] = "OK" if (abs(split.sum() - 1) < 1e-6 and 5 <= P <= 60) else "CHECK"  # BL122
    chk["purchase_cash"] = "OK" if abs(cash.sum() - inp.purchase_amount) < 1 else "CHECK"   # BL123
    chk["ramp_up"] = "OK" if (R >= 0 and R < P and ramp_sum <= 1) else "CHECK"               # BL126
    chk["leverage_months"] = "OK" if inp.last_draw >= inp.facility_start else "CHECK"       # C141
    warn = (not inp.leverage) or inp.last_draw >= P                                          # D141
    chk["last_draw_warning"] = "OK" if warn else "WARNING"
    labels = {"term_mix": "term mix ≠ 100%", "loss_curves": "loss curves", "fee_table": "fee table",
              "target_tiers": "target tiers", "purchase_split": "purchase split",
              "purchase_cash": "purchase cash", "ramp_up": "ramp-up", "leverage_months": "leverage months"}
    fails = [labels[k] for k in labels if chk[k] != "OK"]
    chk["input_status"] = "CHECK" if fails else ("OK" if warn else "WARNING")               # B12
    chk["input_failures"] = fails
    d["checks"] = chk
    return d


def _pmt(r, n):
    if n <= 0:
        return 0.0
    return 1.0 / n if r == 0 else r / (1 - (1 + r) ** -n)


def _eomonth(start: dt.date, k: int) -> dt.date:
    y, m = start.year + (start.month - 1 + k) // 12, (start.month - 1 + k) % 12 + 1
    nxt = dt.date(y + (m // 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


# =============================================================================== F2
def term_block(inp: Inputs, d: dict, k: int) -> dict:
    """One term block of F2 (rows 84-162 for 24m; identical layout for 36/48/60m). Per $1 of face."""
    T, r, pmt = d["T"][k], d["r"][k], d["pmt"][k]
    loss, smm = d["loss"][k], d["smm"][k]
    stop = inp.payment_stop
    share = d["stop_share"][k]
    N = AGES
    t_ = np.arange(N)
    # rows 89-90: scheduled balance and principal factor
    S = np.zeros(N)
    for t in range(N):
        if t <= T:
            S[t] = 1 - t / T if r == 0 else max(0.0, (1 + r) ** t - pmt * ((1 + r) ** t - 1) / r)
    fac = np.zeros(N)
    for t in range(1, N):
        if t <= T and S[t - 1] > 0:
            fac[t] = (S[t - 1] - S[t]) / S[t - 1]
    # row 105: balance of one defaulting loan at charge-off (per $1 defaulted face)
    defbal = np.zeros(N)
    if stop:
        for t in range(N):
            if t <= T:
                v = S[max(0, t - 5)]
                for j in range(1, 5):
                    if t - j >= 1:
                        v -= share[j - 1] * (pmt - r * S[t - j - 1])
                defbal[t] = v
    opn, co, surv, intr, sch, pre, end = (np.zeros(N) for _ in range(7))
    ppsurv, s104, defface, active = (np.zeros(N) for _ in range(4))
    for t in range(N):
        opn[t] = 1.0 if t == 0 else end[t - 1]
        # row 94 charge-offs (curve increment, capped)
        if t == 0:
            co[t] = min(opn[t], loss[0])
        else:
            cap = (S[t - 1] * s104[t - 1] * ppsurv[t - 1] if t <= T else 0.0) if stop else opn[t]
            co[t] = min(max(0.0, loss[t] - loss[t - 1]), cap)
        # row 103 prepayment survival
        ppsurv[t] = 1 - smm[0] if t == 0 else (ppsurv[t - 1] * (1 - smm[t]) if t <= T else ppsurv[t - 1])
        # row 104 survivors as if payment stop OFF
        if not stop:
            s104[t] = 0.0
        elif t == 0:
            s104[t] = 1 - co[0] / opn[0] if opn[0] > 0 else 0.0
        elif t <= T:
            base = S[t - 1] * s104[t - 1] * ppsurv[t - 1]
            s104[t] = s104[t - 1] * (1 - co[t] / base) if base > 0 else 0.0
        else:
            s104[t] = s104[t - 1]
        # row 106 defaulted face
        defface[t] = co[t] / defbal[t] if (stop and t <= T and defbal[t] > 0) else 0.0
        # row 95 surviving balance
        if t == 0:
            surv[t] = opn[t] - (defface[t] if stop else co[t])
        else:
            surv[t] = opn[t] - (defface[t] * S[t - 1] if stop else co[t])
        intr[t] = 0.0 if t == 0 else surv[t] * r                       # row 96
        sch[t] = 0.0 if t == 0 else surv[t] * fac[t]                   # row 97
        pre[t] = smm[t] * (surv[t] - sch[t])                           # row 98
        end[t] = surv[t] - sch[t] - pre[t]                             # row 99
        active[t] = 1.0 if t == 0 else (opn[t] / S[t - 1] if S[t - 1] > 0 else 0.0)  # row 100
    lag = int(inp.recovery_lag)
    rec = np.array([inp.recovery_rate * co[t - lag] if t - lag >= 0 else 0.0 for t in range(N)])  # row 101
    missed = np.zeros(N)                                                # row 102
    if stop:
        for t in range(1, N):
            if t < T:
                s = 0.0
                for j in range(1, 5):
                    if t + j <= T and t + j < N:
                        s += (1 - share[j - 1]) * defface[t + j]
                missed[t] = pmt * s
    removed = opn - surv                                                # row 107
    gross = intr + sch + pre + rec - missed                             # row 108
    size = inp.loan_size[k]
    fees = np.zeros(N)                                                  # row 109
    for t in range(N):
        v = d["c_coll"] * gross[t]
        if t >= 1:
            v += (d["c_beg"] * opn[t] + d["c_end"] * end[t] + d["c_avg"] * (opn[t] + end[t]) / 2) / 12
            v += (d["c_fix"] / size * active[t]) if size else 0.0
        fees[t] = min(gross[t], v)
    net = gross - fees                                                  # row 110
    sp = split_engine(net, d["target_outlay"][k], d["tiers"], d["after_pct"], inp.target_mode)

    outlay = d["outlay"][k]
    inv_cf = sp["investor"].copy(); inv_cf[0] -= outlay                 # row 153
    cum = np.cumsum(sp["investor"])                                     # row 154
    moic_run = cum / outlay if outlay else np.zeros(N)                  # row 155
    irr_run = np.array([np.nan] + [irr_annual(inv_cf[:t + 1]) for t in range(1, N)])  # row 156
    school_cf = sp["school"].copy(); school_cf[0] += d["adv"][k]        # row 157
    df = pd.DataFrame({
        "age": t_, "sched_balance": S, "sched_factor": fac, "loss_curve": loss, "smm": smm,
        "opening": opn, "charge_offs": co, "surviving": surv, "interest": intr, "sched_principal": sch,
        "prepaid": pre, "ending": end, "active_loans": active, "recoveries": rec, "missed": missed,
        "prepay_survival": ppsurv, "survivors_stop_off": s104, "defaulting_balance": defbal,
        "defaulted_face": defface, "removed": removed, "gross": gross, "fees": fees, "net": net,
        "investor": sp["investor"], "school": sp["school"], "tiers_met": sp["met"],
        "tier_in_force": sp["tier_in_force"], "investor_pct_in_force": sp["pct_in_force"],
        "need_next": sp["need_next"], "investor_cf": inv_cf, "investor_cum": cum, "moic_running": moic_run,
        "irr_running": irr_run, "school_cf": school_cf, "school_cum": np.cumsum(school_cf),
        "upfront_fees_paid": np.r_[d["upfront"][k], np.zeros(N - 1)], "across_life_fees_paid": fees,
    }).set_index("age")
    n = d["n_tiers"]
    for i in range(10):
        df[f"need_tier_{i + 1}"] = sp["need_all"][i]                       # rows 112-121
    for i in range(11):
        df[f"breakpoint_{i}"] = sp["bp_all"][i]                            # rows 122-132
    for i in range(10):
        df[f"alloc_tier_{i + 1}"] = sp["alloc_all"][i]                     # rows 133-142
    df["alloc_after_last"] = sp["after"]                                   # row 143
    df["tier_in_force_label"] = ["After" if x > n else f"Tier {x}" for x in sp["tier_in_force"]]   # row 148
    met = sp["met"]
    df["target_hit_flag"] = [(f"Tier {met[0]} hit" if met[0] > 0 else "") if t == 0 else
                             (f"Tier {met[t]} hit" if met[t] > met[t - 1] else "") for t in range(N)]  # row 150
    return {"df": df, "split": sp, "outlay": outlay}


def split_engine(net, outlay, tiers, after_pct, mode):
    """F2 rows 112-151: investor / school split of net collections through up to 10 tiers.
    MOIC mode: need = target x outlay - cash received. IRR mode: need grows at the target IRR monthly.
    A tier crossed mid-month splits that month at the breakpoint (no overshoot)."""
    N = len(net)
    n = len(tiers)
    need = np.zeros((10, N))
    bp = np.zeros((11, N))
    alloc = np.zeros((10, N))
    after = np.zeros(N)
    inv = np.zeros(N)
    for t in range(N):
        for k in range(10):
            if k >= n:
                need[k, t] = 0.0
            elif t == 0:
                need[k, t] = tiers[k][1] * outlay if mode == "MOIC" else outlay
            elif mode == "MOIC":
                need[k, t] = need[k, t - 1] - inv[t - 1]
            else:
                need[k, t] = (need[k, t - 1] - inv[t - 1]) * (1 + tiers[k][1]) ** (1 / 12)
        for k in range(10):
            if k >= n:
                bp[k + 1, t] = bp[k, t]
                continue
            prev = max(need[k - 1, t], 0) if k > 0 else 0.0
            inc = max(0.0, max(need[k, t], 0) - prev) if k > 0 else max(0.0, need[k, t])
            pct = tiers[k][0] or 0.0
            bp[k + 1, t] = bp[k, t] + (0.0 if inc == 0 else (BIG if pct == 0 else inc / pct))
        for k in range(10):
            alloc[k, t] = max(0.0, min(net[t], bp[k + 1, t]) - bp[k, t])
        after[t] = max(0.0, net[t] - bp[10, t])
        inv[t] = sum((tiers[k][0] if k < n else 0.0) * alloc[k, t] for k in range(10)) + after_pct * after[t]
    school = net - inv
    met = np.array([sum(1 for k in range(n) if need[k, t] - inv[t] <= TOL) for t in range(N)])
    tif = np.array([min(n + 1, 1) if t == 0 else min(n + 1, 1 + met[t - 1]) for t in range(N)])
    pct_in_force = np.array([after_pct if x > n else tiers[x - 1][0] for x in tif]) if n else np.full(N, after_pct)
    need_next = np.array([0.0 if met[t] >= n else max(0.0, need[met[t], t] - inv[t]) for t in range(N)])
    return {"investor": inv, "school": school, "alloc": alloc[:n], "after": after, "need": need[:n],
            "need_all": need, "bp_all": bp, "alloc_all": alloc,
            "met": met, "tier_in_force": tif, "pct_in_force": pct_in_force, "need_next": need_next}


def unit_economics(inp: Inputs, d: dict) -> dict:
    blocks = [term_block(inp, d, k) for k in range(4)]
    mix = d["mix"]
    cols = ["opening", "charge_offs", "interest", "sched_principal", "prepaid", "ending", "recoveries",
            "missed", "removed", "gross", "fees", "net", "investor", "school", "investor_cf"]
    bl = pd.DataFrame({c: sum(mix[k] * blocks[k]["df"][c].values for k in range(4)) for c in cols},
                      index=blocks[0]["df"].index)                                  # rows 58-72
    bl["investor_cum"] = np.cumsum(bl["investor"].values)
    F10 = d["F73"]
    bl["moic_running"] = bl["investor_cum"] / F10 if F10 else 0.0
    bl["irr_running"] = [np.nan] + [irr_annual(bl["investor_cf"].values[:t + 1]) for t in range(1, AGES)]
    sch = bl["school"].values.copy(); sch[0] += d["F71"]
    bl["school_cf"] = sch                                                            # row 76
    bl["school_cum"] = np.cumsum(sch)
    bl["cost_basis"] = sum(mix[k] * d["adv"][k] * blocks[k]["df"]["ending"].values for k in range(4))  # row 80

    pv_m = (1 + inp.school_pv_rate) ** (1 / 12) - 1
    rows = {}
    for k in range(5):
        if k < 4:
            df = blocks[k]["df"]
            adv, up = d["adv"][k], d["upfront"][k]
            gross, fees, net = df.gross.sum(), df.fees.sum(), df.net.sum()
            co, missed, inv = df.charge_offs.sum(), df.missed.sum(), df.investor.sum()
            school_share = df.school.sum()
            irr = irr_annual(df.investor_cf.values)
            pv = df.school_cf.values[0] + npv_excel(pv_m, df.school_cf.values[1:])
            met = int(df.tiers_met.values[-1])
            hit = int(np.argmax(df.tiers_met.values >= 1)) if (df.tiers_met.values >= 1).any() else None
            name = f"{d['T'][k]}m"
            mixw, apr = d["mix"][k], inp.apr[k]
            tiers_txt = f"{met} of {d['n_tiers']}"
        else:
            adv, up = d["F71"], d["F72"]
            gross, fees, net = bl.gross.sum(), bl.fees.sum(), bl.net.sum()
            co, missed, inv = bl.charge_offs.sum(), bl.missed.sum(), bl.investor.sum()
            school_share = bl.school.sum()
            irr = irr_annual(bl.investor_cf.values)
            pv = sch[0] + npv_excel(pv_m, sch[1:])
            hit, tiers_txt, name = None, "n.a.", "Blended"
            mixw, apr = float(d["mix"].sum()), d["wa_apr"]
        outlay = adv + up
        rows[name] = {
            "Term (months)": d["T"][k] if k < 4 else d["wa_term"], "Mix weight": mixw, "APR": apr,
            "Advance paid to school": adv, "Upfront fees": up, "Investor outlay": outlay,
            "Gross collections": gross, "Across-life fees": fees, "Net collections": net,
            "Charge-offs": co, "Missed payments": missed, "Investor cash received": inv,
            "Investor MOIC": inv / outlay if outlay else np.nan, "Investor IRR": irr,
            "Tiers hit": tiers_txt, "Month tier 1 hit": hit if hit is not None else ("n.a." if k == 4 else "Not hit"),
            "School day-0 advance": adv, "School collections": school_share,
            "School total (undiscounted)": adv + school_share, "School PV of collections": pv - adv,
            "School total (PV)": pv, "Effective discount (undiscounted)": 1 - (adv + school_share),
            "Effective discount (PV)": 1 - pv,
        }
    summary = pd.DataFrame(rows)
    # tier summary (rows 41-52)
    tier_rows = []
    for i, (pct, tgt) in enumerate(d["tiers"]):
        rec = {"Tier": f"Tier {i + 1}", "Investor %": pct,
               "Target": f"{tgt:.2f}x MOIC" if inp.target_mode == "MOIC" else f"{tgt:.1%} IRR"}
        for k in range(4):
            df = blocks[k]["df"]; sp = blocks[k]["split"]
            met = df.tiers_met.values
            age = int(np.argmax(met >= i + 1)) if (met >= i + 1).any() else None
            rec[f"{d['T'][k]}m age hit"] = age if age is not None else "Not hit"
            rec[f"{d['T'][k]}m investor $"] = pct * sp["alloc"][i].sum()
            rec[f"{d['T'][k]}m school $"] = (1 - pct) * sp["alloc"][i].sum()
            rec[f"{d['T'][k]}m MOIC at hit"] = df.moic_running.values[age] if age is not None else None
            rec[f"{d['T'][k]}m IRR at hit"] = df.irr_running.values[age] if age is not None else None
        tier_rows.append(rec)
    checks = {}
    for k in range(4):
        df = blocks[k]["df"]
        c1 = np.abs(np.round(df.investor + df.school - df.net, 9)).sum()
        c2 = round(df.removed.sum() + df.sched_principal.sum() + df.prepaid.sum() + df.ending.values[-1] - 1, 9)
        checks[f"{d['T'][k]}m"] = "OK" if (c1 < 1e-6 and abs(c2) < 1e-6) else "CHECK"
    cb = np.abs(np.round(bl.investor + bl.school - bl.net, 9)).sum()
    checks["Blended"] = "OK" if (cb < 1e-6 and all(v == "OK" for v in checks.values())) else "CHECK"
    return {"blocks": blocks, "blended": bl, "summary": summary, "tier_summary": pd.DataFrame(tier_rows),
            "checks": checks}


# =============================================================================== F2a
def asset_returns(inp: Inputs, d: dict, ue: dict) -> dict:
    """Gross (no fees) vs net. The tier split is re-run on gross collections with targets on price only."""
    mix = d["mix"]
    gross_inv_cf, out = [], {}
    for k in range(4):
        df = ue["blocks"][k]["df"]
        adv = d["adv"][k]
        sp = split_engine(df.gross.values, adv, d["tiers"], d["after_pct"], inp.target_mode)
        cf = sp["investor"].copy(); cf[0] -= adv
        gross_inv_cf.append(cf)
        out[f"{d['T'][k]}m"] = {
            "Mix weight": mix[k], "Price paid (advance)": adv, "Gross collections": df.gross.sum(),
            "Investor cash, gross": sp["investor"].sum(), "Investor MOIC, gross": sp["investor"].sum() / adv if adv else 0,
            "Investor IRR, gross": irr_annual(cf),
            "Investor MOIC, net": ue["summary"].loc["Investor MOIC", f"{d['T'][k]}m"],
            "Investor IRR, net": ue["summary"].loc["Investor IRR", f"{d['T'][k]}m"],
            "School total, gross": adv + sp["school"].sum(),
            "School total, net": ue["summary"].loc["School total (undiscounted)", f"{d['T'][k]}m"],
            "_check": round(sp["investor"].sum() + sp["school"].sum() - df.gross.sum(), 9),
        }
    names = [f"{t}m" for t in d["T"]]
    bl = {}
    for key in ["Price paid (advance)", "Gross collections", "Investor cash, gross", "School total, gross"]:
        bl[key] = sum(mix[k] * out[names[k]][key] for k in range(4))
    bl["Mix weight"] = float(mix.sum())
    bl["Investor MOIC, gross"] = bl["Investor cash, gross"] / bl["Price paid (advance)"] if bl["Price paid (advance)"] else 0
    bcf = sum(mix[k] * gross_inv_cf[k] for k in range(4))
    bl["Investor IRR, gross"] = irr_annual(bcf)
    bl["Investor MOIC, net"] = ue["summary"].loc["Investor MOIC", "Blended"]
    bl["Investor IRR, net"] = ue["summary"].loc["Investor IRR", "Blended"]
    bl["School total, net"] = ue["summary"].loc["School total (undiscounted)", "Blended"]
    bl["_check"] = sum(out[n]["_check"] for n in names)
    out["Blended"] = bl
    df = pd.DataFrame(out)
    df.loc["Fee drag on MOIC"] = df.loc["Investor MOIC, gross"] - df.loc["Investor MOIC, net"]
    check = "OK" if abs(bl["_check"]) < 1e-6 else "CHECK"
    return {"table": df.drop(index="_check"), "check": check, "gross_cf_blended": bcf}


# =============================================================================== F3
def pool(inp: Inputs, d: dict, ue: dict) -> dict:
    M = MONTHS
    m_ = np.arange(1, M + 1)
    face = np.array([d["plan_face"][m - 1] if m <= inp.purchase_period else 0.0 for m in m_])  # row 7
    cash = face * d["F71"]; upf = face * d["F72"]; outlay = cash + upf                          # rows 8-10
    bl = ue["blended"]

    def stack(series):  # SUMPRODUCT of face by month with the per-$1 pattern by age (helpers rows 233-247)
        pat = np.zeros(M); pat[:AGES] = series.values[:min(AGES, M)]
        return np.array([np.dot(face[:m], pat[m - 1::-1][:m]) for m in m_])

    p = pd.DataFrame(index=pd.Index(m_, name="month"))
    p["date"] = d["dates"]
    p["face"], p["cash_to_school"], p["upfront_fees"], p["outlay"] = face, cash, upf, outlay
    for col in ["opening", "charge_offs", "interest", "sched_principal", "prepaid", "recoveries", "missed",
                "gross", "fees", "net", "investor", "school", "ending", "cost_basis", "removed"]:
        p[col] = stack(bl[col])
    p["school_total"] = p.cash_to_school + p.school                                              # row 33
    p["school_cum"] = p.school_total.cumsum()
    p["unlev_cf"] = p.investor - p.outlay                                                        # row 39
    p["active"] = ((p.face > 0) | (p.ending > 0.5) | (p.net > 0.5)).astype(int)                 # row 43
    # leverage + fees + SPV cash (rows 44-73), month by month
    cols = {c: np.zeros(M) for c in ["mode", "eligible", "capacity", "sofr", "rate", "debt_open", "interest_debt",
                                     "draws", "repay", "debt_end", "other_fees", "net_spv", "operating",
                                     "equity", "dist", "recycled"]}
    mode = []
    debt = 0.0
    for i, m in enumerate(m_):
        md = "Off" if not inp.leverage else ("Pre-raise" if m < inp.facility_start else
                                              ("Origination" if m <= inp.last_draw else "Amortization"))
        mode.append(md)
        elig = p.cost_basis.values[i] * (1 - inp.ineligibles)
        cap = min(inp.ltv * elig, d["facility_max"]) if inp.leverage else 0.0
        sofr = inp.sofr[i] if i < len(inp.sofr) else 0.0
        rate = sofr + inp.spread
        opn = debt
        if md == "Origination":
            draw = max(0.0, cap - opn) if inp.recycling else min(max(0.0, cap - opn), outlay[i])
        else:
            draw = 0.0
        intr = ((opn + opn + draw - max(0.0, opn - cap)) / 2 if md == "Origination" else opn) * rate / 12
        inv = p.investor.values[i]
        annual = inp.fee_high if math.ceil(m / 12) <= inp.fee_years_high else inp.fee_low
        fee = 0.0 if inv <= 0 else -(p.active.values[i] * min(annual / 12, max(0.0, inv - intr)))
        if md == "Origination":
            rep = max(0.0, opn - cap)
        elif md == "Amortization":
            rep = min(opn, max(0.0, opn - cap, inp.sweep * (inv + fee - intr)))
        else:
            rep = 0.0
        debt = opn + draw - rep
        net_spv = inv + draw - outlay[i] + fee - intr - rep                                     # row 67
        operating = inv + fee - intr - rep                                                     # row 68
        if inp.recycling:
            eq, dist = max(0.0, -net_spv), max(0.0, net_spv)
            recyc = min(max(0.0, operating), max(0.0, -(-outlay[i] + draw)))
        else:
            eq = max(0.0, -(-outlay[i] + draw)) + max(0.0, -operating)
            dist = max(0.0, operating) + max(0.0, draw - outlay[i])
            recyc = 0.0
        for c_, v in zip(["eligible", "capacity", "sofr", "rate", "debt_open", "interest_debt", "draws", "repay",
                          "debt_end", "other_fees", "net_spv", "operating", "equity", "dist", "recycled"],
                         [elig, cap, sofr, rate, opn, intr, draw, rep, debt, fee, net_spv, operating, eq, dist, recyc]):
            cols[c_][i] = v
    p["mode"] = mode
    for c_ in cols:
        if c_ != "mode":
            p[c_] = cols[c_]
    p["net_to_investors"] = p.dist - p.equity                                                    # row 73
    # F5 section 5: how each month's purchases are funded
    p["fund_recycled"] = p.recycled
    p["fund_debt"] = np.minimum(np.maximum(0, p.draws), np.maximum(0, p.outlay - p.recycled))
    p["fund_equity"] = np.maximum(0, p.outlay - p.fund_recycled - p.fund_debt)
    # vintage grids (sections 9 and 10)
    nv = min(inp.purchase_period, 60)
    inv_pat, sch_pat = bl["investor"].values, bl["school_cf"].values
    vin_inv = np.zeros((nv, M)); vin_sch = np.zeros((nv, M))
    for v in range(1, nv + 1):
        for m in range(v, M + 1):
            a = m - v
            if a <= 60:
                vin_inv[v - 1, m - 1] = face[v - 1] * inv_pat[a]
                vin_sch[v - 1, m - 1] = face[v - 1] * sch_pat[a]
    S = {}
    S["Total face purchased"] = face.sum()
    S["Cash paid to schools"] = cash.sum()
    S["Upfront fees paid by investor"] = upf.sum()
    S["Gross collections"] = p.gross.sum()
    S["Across-life fees"] = p.fees.sum()
    S["Net collections"] = p.net.sum()
    S["Investor share of net collections"] = p.investor.sum()
    S["School share of net collections"] = p.school.sum()
    S["Charge-offs"] = p.charge_offs.sum()
    S["Charge-offs (% of face)"] = p.charge_offs.sum() / face.sum() if face.sum() else 0
    S["Unlevered investor MOIC"] = p.investor.sum() / outlay.sum() if outlay.sum() else 0
    S["Unlevered investor IRR"] = irr_annual(p.unlev_cf.values)
    S["School effective advance (undiscounted)"] = p.school_total.sum() / face.sum() if face.sum() else 0
    pv_m = (1 + inp.school_pv_rate) ** (1 / 12) - 1
    st_ = p.school_total.values
    S["School effective advance (PV)"] = (st_[0] + npv_excel(pv_m, st_[1:])) / face.sum() if face.sum() else 0  # F5 B21
    S["Peak debt"] = p.debt_end.max()
    S["Total debt draws"] = p.draws.sum()
    S["Total interest"] = p.interest_debt.sum()
    S["Other SPV fees"] = -p.other_fees.sum()
    S["Equity called"] = p.equity.sum()
    S["Distributions"] = p.dist.sum()
    S["Collections recycled into purchases"] = p.recycled.sum()
    S["Equity MOIC"] = p.dist.sum() / p.equity.sum() if p.equity.sum() else 0
    S["Equity IRR"] = irr_annual(p.net_to_investors.values)
    # checks (rows 26-28, 72, 88, 90, 164, 228 -> B99)
    c26 = np.abs(np.round(p.gross - (p.interest + p.sched_principal + p.prepaid + p.recoveries - p.missed), 4)).sum()
    c27 = np.abs(np.round(p.investor + p.school - p.net, 4)).sum()
    prev_end = np.concatenate([[0.0], p.ending.values[:-1]])
    c28 = np.abs(np.round(p.ending.values - (prev_end + face - p.removed.values - p.sched_principal.values
                                              - p.prepaid.values), 4)).sum()
    c72 = np.abs(np.round(p.dist - p.equity - p.net_spv, 4)).sum()
    ue_bl = ue["summary"]["Blended"]
    t88 = round(S["Unlevered investor MOIC"] - ue_bl["Investor MOIC"], 6) + round(S["Unlevered investor IRR"] - ue_bl["Investor IRR"], 6)
    t90 = round(S["School effective advance (undiscounted)"] - ue_bl["School total (undiscounted)"], 6)
    v_inv = np.abs(np.round(vin_inv.sum(0) - p.investor.values, 4)).sum()
    v_sch = np.abs(np.round(vin_sch.sum(0) - p.school_total.values, 4)).sum()
    ok = (c26 + c27 + c28 + c72 == 0 and abs(p.debt_end.values[-1]) < 1 and t88 == 0 and t90 == 0
          and v_inv < 1 and v_sch < 1)
    checks = {"pool_identities": c26 + c27 + c28 + c72, "debt_repaid": p.debt_end.values[-1],
              "tie_to_F2": t88 + t90, "vintage_ties": v_inv + v_sch, "status": "OK" if ok else "CHECK"}
    return {"monthly": p, "summary": S, "vintage_investor": vin_inv, "vintage_school": vin_sch, "checks": checks}


# =============================================================================== F4
def spv(inp: Inputs, d: dict, pl: dict, light: bool = False) -> dict:
    p = pl["monthly"]
    M = len(p)
    eq_tot = p.equity.sum()
    if inp.capital_basis == "OCM %":
        la = 1 - inp.ocm_pct
    else:
        la = min(1.0, inp.la_fixed / eq_tot) if eq_tot else 0.0                     # B8
    oc = 1 - la
    hm = [(1 + h) ** (1 / 12) - 1 for h in inp.hurdles]                                 # B11:D11
    s = [x + (1 - x) * la if inp.promote_method == "On top of pro rata" else x for x in inp.la_tier_shares]  # B12:D12
    eq, dist = p.equity.values, p.dist.values
    out = {c: np.zeros(M) for c in ["t1", "t2", "t3", "t4", "t5", "ocm", "la", "hb2", "hb3", "hb4"]}
    cum_eq = cum_t1 = 0.0
    hb2 = hb3 = hb4 = 0.0
    tier_paid = []
    for i in range(M):
        E, D = eq[i], dist[i]
        cum_eq += E
        t1 = min(D, max(0.0, cum_eq - cum_t1)); cum_t1 += t1                           # rows 38-39
        hb2_b = E if i == 0 else hb2 * (1 + hm[0]) + E                                   # row 40
        t12 = max(t1, min(D, max(0.0, hb2_b)))                                           # row 41
        hb2 = hb2_b - D                                                                  # row 43
        ocm12 = t12 * oc                                                                 # row 44
        hb3_b = E * oc if i == 0 else hb3 * (1 + hm[1]) + E * oc                         # row 45
        hb4_b = E * oc if i == 0 else hb4 * (1 + hm[2]) + E * oc                         # row 47
        if inp.promote:
            t3 = min(D - t12, max(0.0, hb3_b - ocm12) / (1 - s[0])) if s[0] < 1 else D - t12
            t4 = min(D - t12 - t3, max(0.0, hb4_b - ocm12 - t3 * (1 - s[0])) / (1 - s[1])) if s[1] < 1 else D - t12 - t3
        else:
            t3 = t4 = 0.0
        t5 = D - t12 - t3 - t4                                                           # row 49
        o = (ocm12 + t3 * (1 - s[0]) + t4 * (1 - s[1]) + t5 * (1 - s[2])) if inp.promote else D * oc  # row 50
        hb3 = hb3_b - o; hb4 = hb4_b - o                                                 # rows 52-53
        for c_, v in zip(["t1", "t2", "t3", "t4", "t5", "ocm", "la", "hb2", "hb3", "hb4"],
                         [t1, t12 - t1, t3, t4, t5, o, D - o, hb2_b, hb3_b, hb4_b]):
            out[c_][i] = v
        tier_paid.append(0 if D <= 0 else (5 if t5 > 0.005 else 4 if t4 > 0.005 else 3 if t3 > 0.005 else 2 if t12 - t1 > 0.005 else 1))
    w = pd.DataFrame(out, index=p.index)
    w["date"] = p["date"].values
    w["equity"], w["dist"] = eq, dist
    w["ocm_equity"], w["la_equity"] = eq * oc, eq * la
    w["highest_tier"] = tier_paid
    w["la_t12"] = (w.t1 + w.t2) * la
    w["la_t3"] = w.t3 * s[0]; w["la_t4"] = w.t4 * s[1]
    w["la_t5"] = w.t5 * (s[2] if inp.promote else la)
    w["ocm_net"] = w.ocm - w.ocm_equity
    w["la_net"] = w.la - w.la_equity
    w["total_net"] = p.net_to_investors.values
    w["ocm_moic_running"] = np.where(w.ocm_equity.cumsum() > 0, w.ocm.cumsum() / w.ocm_equity.cumsum().replace(0, np.nan), 0)
    w["la_moic_running"] = np.where(w.la_equity.cumsum() > 0, w.la.cumsum() / w.la_equity.cumsum().replace(0, np.nan), 0)
    w["total_moic_running"] = np.where(w.equity.cumsum() > 0, w.dist.cumsum() / w.equity.cumsum().replace(0, np.nan), 0)
    w["debt_recap"] = [0.0 if not inp.recycling else (min(max(dist[i], 0), max(p.draws.values[i], 0))
                       if p.index[i] == inp.facility_start else 0.0) for i in range(M)]   # row 73
    w["la_promote"] = np.maximum(0, (w.la_t3 + w.la_t4 + w.la_t5) - (w.t3 + w.t4 + w.t5) * la)   # F4 row 101
    w["la_pro_rata"] = w.la - w.la_promote                                                 # F4 row 100
    w["la_share_month"] = np.where(w.dist > 0.5, w.la / w.dist.where(w.dist > 0.5, 1), np.nan)  # F4 row 102
    w["ocm_cum_dist"] = w.ocm.cumsum(); w["la_cum_dist"] = w.la.cumsum()                   # rows 104-105
    w["spv_cum_net"] = (w.dist - w.equity).cumsum()                                         # row 111
    w["ocm_cum_net"] = (w.ocm - w.ocm_equity).cumsum()                                      # row 114
    w["la_cum_net"] = (w.la - w.la_equity).cumsum()                                         # row 118
    if light:
        w["ocm_irr_running"] = w["la_irr_running"] = w["total_irr_running"] = np.nan
    else:
        w["ocm_irr_running"] = [irr_annual(w.ocm_net.values[:i + 1]) for i in range(M)]       # row 70
        w["la_irr_running"] = [irr_annual(w.la_net.values[:i + 1]) for i in range(M)]         # row 71
        w["total_irr_running"] = [irr_annual(w.total_net.values[:i + 1]) for i in range(M)]   # F5 row 48
    recap = w.debt_recap.sum()

    def party(eqc, dc, netc):
        e, dd = w[eqc].sum(), w[dc].sum()
        return {"Equity called": e, "Distributions": dd, "Net profit": dd - e,
                "MOIC": dd / e if e else np.nan, "IRR": irr_annual(w[netc].values)}
    R = pd.DataFrame({"Total investors": party("equity", "dist", "total_net"),
                      "OCM": party("ocm_equity", "ocm", "ocm_net"),
                      "Loan America": party("la_equity", "la", "la_net")})
    R.loc["Share of capital"] = [1.0, oc, la]
    R.loc["Share of distributions"] = [1.0, R.loc["Distributions", "OCM"] / R.loc["Distributions", "Total investors"]
                                       if R.loc["Distributions", "Total investors"] else 0,
                                       R.loc["Distributions", "Loan America"] / R.loc["Distributions", "Total investors"]
                                       if R.loc["Distributions", "Total investors"] else 0]
    R.loc["Debt recap"] = [recap, recap * oc, recap * la]
    R.loc["Adjusted MOIC"] = [(R.loc["Distributions", c] - R.loc["Debt recap", c]) / (R.loc["Equity called", c] - R.loc["Debt recap", c])
                              if (R.loc["Equity called", c] - R.loc["Debt recap", c]) else np.nan for c in R.columns]
    tiers = pd.DataFrame({
        "Total investors": [w.t1.sum(), w.t2.sum(), w.t3.sum(), w.t4.sum(), w.t5.sum()],
        "OCM": [w.t1.sum() * oc, w.t2.sum() * oc, w.t3.sum() * (1 - s[0]), w.t4.sum() * (1 - s[1]),
                w.t5.sum() * ((1 - s[2]) if inp.promote else oc)],
        "Loan America": [w.t1.sum() * la, w.t2.sum() * la, w.la_t3.sum(), w.la_t4.sum(), w.la_t5.sum()],
    }, index=["1. Return of capital", "2. Preferred return", "3. Promote tier (to OCM hurdle 1)",
              "4. Promote tier (to OCM hurdle 2)", "5. Residual"])
    c61 = abs(round((w.la_t12 + w.la_t3 + w.la_t4 + w.la_t5 - w.la).abs().sum(), 4))
    c67 = abs(round((w.ocm_net + w.la_net - w.total_net).abs().sum(), 4))
    ok = (c61 < 1e-3 and c67 < 1e-3 and abs(tiers["Total investors"].sum() - R.loc["Distributions", "Total investors"]) < 1
          and abs(R.loc["Distributions", "OCM"] + R.loc["Distributions", "Loan America"] - R.loc["Distributions", "Total investors"]) < 1
          and abs(R.loc["Equity called", "OCM"] + R.loc["Equity called", "Loan America"] - R.loc["Equity called", "Total investors"]) < 1)
    shares_ok = all(0 <= x <= 1 for x in s)
    # yearly view (section 7)
    w["la_promote_yearly_basis"] = (w.la_t3 + w.la_t4 + w.la_t5) - (w.t3 + w.t4 + w.t5) * la   # row 91 (no floor)
    g = w.assign(year=[dd.year for dd in w.date]).groupby("year")
    first_year = w.date.iloc[0].year
    years = list(range(first_year, first_year + 10))                                      # F4 C77:L77
    agg = g[["equity", "dist", "ocm_equity", "ocm", "la_equity", "la", "la_promote_yearly_basis"]].sum().reindex(years, fill_value=0.0)
    yr = pd.DataFrame(index=years)
    yr["Total: equity called"] = -agg.equity; yr["Total: distributions"] = agg.dist
    yr["Total: net cash flow"] = yr["Total: equity called"] + yr["Total: distributions"]
    yr["Total: cumulative net"] = yr["Total: net cash flow"].cumsum()
    yr["OCM: equity called"] = -agg.ocm_equity; yr["OCM: distributions"] = agg.ocm
    yr["OCM: net cash flow"] = yr["OCM: equity called"] + yr["OCM: distributions"]
    yr["OCM: cumulative net"] = yr["OCM: net cash flow"].cumsum()
    yr["LA: equity called"] = -agg.la_equity; yr["LA: distributions"] = agg.la
    yr["LA: of which promote above pro rata"] = agg.la_promote_yearly_basis
    yr["LA: net cash flow"] = yr["LA: equity called"] + yr["LA: distributions"]
    yr["LA: cumulative net"] = yr["LA: net cash flow"].cumsum()
    yr["LA share of distributions"] = np.where(yr["Total: distributions"] != 0, yr["LA: distributions"] / yr["Total: distributions"].replace(0, np.nan), 0)
    yr["Check: OCM + LA - total (0)"] = (yr["OCM: distributions"] + yr["LA: distributions"] - yr["Total: distributions"]).round(2)
    return {"monthly": w, "returns": R, "tiers": tiers, "la_share": la, "effective_la_shares": s,
            "monthly_hurdles": hm, "yearly": yr,
            "checks": {"waterfall": "OK" if ok else "CHECK", "promote_shares": "OK" if shares_ok else "CHECK"}}


# =============================================================================== F2b / F4a
def advance_sensitivity(inp: Inputs, steps: int = 5) -> pd.DataFrame:
    rows = {}
    for i in range(-steps, steps + 1):
        x = inp.copy()
        x.advance = [a + i * inp.advance_step for a in inp.advance]
        dd = derive(x)
        ue = unit_economics(x, dd)
        s = ue["summary"]
        col = {"Change (pp)": i * inp.advance_step, "Blended advance": dd["F71"], "Day-0 discount": 1 - dd["F71"]}
        # like F2b: a term whose shifted advance is <= 0 shows n.a.; blended is n.a. if any term is
        bad = [a <= 0 for a in x.advance] + [min(x.advance) <= 0]
        for j, name in enumerate(s.columns):
            col[f"MOIC {name}"] = np.nan if bad[j] else s.loc["Investor MOIC", name]
        for j, name in enumerate(s.columns):
            col[f"IRR {name}"] = np.nan if bad[j] else s.loc["Investor IRR", name]
        col["School total (undiscounted)"] = s.loc["School total (undiscounted)", "Blended"]
        col["School total (PV)"] = s.loc["School total (PV)", "Blended"]
        col["School effective discount (PV)"] = s.loc["Effective discount (PV)", "Blended"]
        rows[i] = col
    return pd.DataFrame(rows)


def split_sensitivity(inp: Inputs, steps: int = 5) -> pd.DataFrame:
    rows = {}
    d0 = derive(inp)
    base = d0["tiers"][0][0] if d0["tiers"] else 0.0
    for i in range(-steps, steps + 1):
        x = inp.copy()
        new = base + i * inp.split_step
        for t in x.tiers:
            if t.active:
                t.investor_pct = new
        res = run_model(x, sensitivities=False, light=True)
        R = res["spv"]["returns"]
        col = {"Change (pp)": i * inp.split_step, "Investor %": new, "School %": 1 - new}
        for party in R.columns:
            for k in ["Equity called", "Distributions", "MOIC", "IRR"]:
                col[f"{party} {k}"] = R.loc[k, party]
        col["LA share of distributions"] = R.loc["Share of distributions", "Loan America"]
        rows[i] = col
    return pd.DataFrame(rows)


def single_split(inp: Inputs, d: dict) -> bool:
    """F2b / F4a require one constant investor % (all active tiers and the after-last % equal)."""
    pcts = {round(p, 12) for p, _ in d["tiers"]} | {round(d["after_pct"], 12)}
    return len(pcts) == 1


# =============================================================================== run all
def run_model(inp: Inputs, sensitivities: bool = True, light: bool = False) -> dict:
    d = derive(inp)
    ue = unit_economics(inp, d)
    ar = asset_returns(inp, d, ue)
    pl = pool(inp, d, ue)
    sp = spv(inp, d, pl, light=light)
    res = {"inputs": inp, "derived": d, "ue": ue, "asset": ar, "pool": pl, "spv": sp}
    calc_ok = (all(v == "OK" for v in ue["checks"].values()) and ar["check"] == "OK" and
               pl["checks"]["status"] == "OK" and sp["checks"]["waterfall"] == "OK" and
               sp["checks"]["promote_shares"] == "OK")
    res["model_status"] = "OK" if calc_ok else "CHECK"                                  # F1 B13
    res["input_status"] = d["checks"]["input_status"]                                  # F1 B12
    if not light:
        res["f5"] = dashboard_views(res)
    if sensitivities:
        ok = single_split(inp, d)
        res["single_split"] = ok
        res["adv_sens"] = advance_sensitivity(inp) if ok else None
        res["split_sens"] = split_sensitivity(inp) if ok else None
        res["sens_checks"] = sensitivity_checks(res)
    return res


# =============================================================================== F5 dashboard views
def dashboard_views(res: dict) -> dict:
    """F5 sections 1-6: returns by party, school effective advance, investor per $1, monthly view,
    recycling funding view and leverage view, with their key points."""
    inp, d = res["inputs"], res["derived"]
    p, w, R = res["pool"]["monthly"], res["spv"]["monthly"], res["spv"]["returns"]
    S, PS = res["ue"]["summary"], res["pool"]["summary"]
    face = PS["Total face purchased"]
    na = float("nan")
    parties = pd.DataFrame({
        "OCM": [R.loc["Equity called", "OCM"], R.loc["Distributions", "OCM"], R.loc["Net profit", "OCM"], R.loc["MOIC", "OCM"], R.loc["IRR", "OCM"], R.loc["Share of capital", "OCM"]],
        "Loan America": [R.loc["Equity called", "Loan America"], R.loc["Distributions", "Loan America"], R.loc["Net profit", "Loan America"], R.loc["MOIC", "Loan America"], R.loc["IRR", "Loan America"], R.loc["Share of capital", "Loan America"]],
        "Total investors": [R.loc["Equity called", "Total investors"], R.loc["Distributions", "Total investors"], R.loc["Net profit", "Total investors"], R.loc["MOIC", "Total investors"], R.loc["IRR", "Total investors"], 1.0],
        "School": [na, p.school_total.sum(), na, na, na, na],
        "Goal (servicer)": [na, p.fees.sum(), na, na, na, na],
    }, index=["Cash put in ($)", "Cash received ($)", "Net profit ($)", "MOIC", "IRR (annualised)", "Share of capital"])
    tiers_txt = "; ".join(f"T{i + 1} {p_:.0%} to {t:.2f}x" if inp.target_mode == "MOIC" else f"T{i + 1} {p_:.0%} to {t:.0%}"
                          for i, (p_, t) in enumerate(d["tiers"]))
    settings = (f"Targets: {inp.target_mode} | after last tier: {inp.after_last_tier} | capital: {inp.capital_basis} | "
                f"promote: {'ON' if inp.promote else 'OFF'} | recycling: {'ON' if inp.recycling else 'OFF'} | "
                f"leverage: {'ON' if inp.leverage else 'OFF'} | tiers: {tiers_txt}")
    names = list(S.columns)
    school = pd.DataFrame({"Program": [face, PS["Cash paid to schools"] / face if face else 0, p.school.sum() / face if face else 0,
                                       PS["School effective advance (undiscounted)"], PS["School effective advance (PV)"], inp.school_pv_rate]},
                          index=["Face of loans sold ($ / per $1)", "Upfront advance at month 0 (% of face)",
                                 "Share of collections, undiscounted (% of face)", "Effective advance rate, undiscounted (% of face)",
                                 "Effective advance rate, PV (% of face)", "Discount rate (annual)"])
    for n in names:
        school[n if n != "Blended" else "Blended per $1"] = [1.0, S.loc["School day-0 advance", n], S.loc["School collections", n],
                                                             S.loc["School total (undiscounted)", n], S.loc["School total (PV)", n], inp.school_pv_rate]
    per1 = pd.DataFrame({n: [S.loc["Investor outlay", n], S.loc["Investor cash received", n], S.loc["Investor MOIC", n],
                             S.loc["Investor IRR", n], S.loc["Month tier 1 hit", n]] for n in names},
                        index=["Investor outlay per $1 (price + upfront fees)", "Investor cash received per $1", "MOIC",
                               "IRR (annualised)", "Month tier 1 hit (loan age)"])
    mv = pd.DataFrame({"OCM net cash flow": w.ocm_net, "Loan America net cash flow": w.la_net,
                       "Total investor net cash flow": w.total_net, "Investor distributions": w.dist,
                       "School cash received": p.school_total, "Goal / across-life fees": p.fees,
                       "OCM running MOIC": w.ocm_moic_running, "Loan America running MOIC": w.la_moic_running,
                       "Total investor running MOIC": w.total_moic_running, "OCM running IRR": w.ocm_irr_running,
                       "Loan America running IRR": w.la_irr_running, "Total investor running IRR": w.total_irr_running,
                       "School cumulative cash (% of program face)": p.school_total.cumsum() / face if face else 0}, index=p.index)
    rec = pd.DataFrame({"Loan purchases ($)": p.outlay, "Funded by recycled collections ($)": p.fund_recycled,
                        "Funded by debt draws ($)": p.fund_debt, "Funded by new equity ($)": p.fund_equity,
                        "Share funded by recycled collections (%)": np.where(p.outlay > 0, p.fund_recycled / p.outlay.where(p.outlay > 0, 1), np.nan),
                        "Equity called in the month ($)": p.equity, "Cumulative equity called ($)": p.equity.cumsum(),
                        "Cumulative purchases ($)": p.outlay.cumsum(),
                        "Check: collections + debt + equity - purchases (0)": (p.fund_recycled + p.fund_debt + p.fund_equity - p.outlay).round(2)},
                       index=p.index)
    sh = rec["Share funded by recycled collections (%)"].values
    first50 = next((int(m) for m, s_ in zip(p.index, sh) if np.isfinite(s_) and s_ >= 0.5), "never")
    first100 = next((int(m) for m, o, e_ in zip(p.index, p.outlay, p.fund_equity) if o > 0 and e_ < 1), "never")
    rec_kp = {"First month collections fund ≥ 50% of that month's purchases": first50,
              "First month collections fund 100% of that month's purchases (no new equity for purchases)": first100,
              "Total equity called ÷ total purchases": p.equity.sum() / p.outlay.sum() if p.outlay.sum() else 0}
    lev = pd.DataFrame({"Facility mode": p["mode"], "Investor cost basis ($)": p.cost_basis, "Borrowing capacity ($)": p.capacity,
                        "Opening debt ($)": p.debt_open, "Draws ($)": p.draws, "Repayments ($)": -p.repay,
                        "Interest paid ($)": p.interest_debt, "Ending debt ($)": p.debt_end, "Equity called ($)": p.equity,
                        "Distributions to investors ($)": w.dist,
                        "Check: ending = opening + draws + repayments (0)": (p.debt_end - (p.debt_open + p.draws - p.repay)).round(2)},
                       index=p.index)
    peak = float(p.debt_end.max())
    peak_m = int(p.index[p.debt_end.values.argmax()]) if peak > 0 else ""
    first_draw = next((int(m) for m, x in zip(p.index, p.draws) if x > 0.5), "none")
    repaid = "n.a."
    if peak > 0:
        repaid = next((int(m) for m, x in zip(p.index, p.debt_end) if x < 0.5 and m > peak_m), "not repaid")
    lev_kp = {"Leverage switch": "ON" if inp.leverage else "OFF", "First draw month": first_draw, "Peak debt ($)": peak,
              "Peak debt month": peak_m, "Debt fully repaid by month": repaid, "Total interest paid ($)": p.interest_debt.sum(),
              "Debt recap: start-month distributions funded by the first draw ($)": R.loc["Debt recap", "Total investors"]}
    return {"parties": parties, "la_orig_fees": PS["Upfront fees paid by investor"], "settings": settings, "school": school,
            "per1": per1, "monthly": mv, "recycling": rec, "recycling_kp": rec_kp, "leverage": lev, "leverage_kp": lev_kp}


def sensitivity_checks(res: dict) -> dict:
    """F2b rows 32-35 and F4a rows 32-34: the middle column must equal the live model."""
    out = {}
    a, s = res.get("adv_sens"), res.get("split_sens")
    S, R = res["ue"]["summary"], res["spv"]["returns"]
    if a is not None:
        c32 = (abs(a.loc["MOIC Blended", 0] - S.loc["Investor MOIC", "Blended"]) + abs(a.loc["IRR Blended", 0] - S.loc["Investor IRR", "Blended"])
               + abs(a.loc["School total (undiscounted)", 0] - S.loc["School total (undiscounted)", "Blended"]))
        c33 = sum(abs(a.loc[f"MOIC {n}", 0] - S.loc["Investor MOIC", n]) for n in S.columns[:4])
        c34 = sum(abs(a.loc[f"IRR {n}", 0] - S.loc["Investor IRR", n]) for n in S.columns[:4])
        out["F2b"] = {"middle = blended F2": round(c32, 9), "middle MOIC by term": round(c33, 9), "middle IRR by term": round(c34, 9),
                      "status": "OK" if max(c32, c33, c34) < 1e-6 else "CHECK"}
    if s is not None:
        c32 = sum(abs(s.loc[f"{p_} {k}", 0] - R.loc[k, p_]) for p_ in R.columns for k in ["Equity called", "Distributions", "MOIC", "IRR"]
                  if np.isfinite(R.loc[k, p_]))
        c33 = float(np.nansum(np.abs(s.loc["OCM Equity called"] + s.loc["Loan America Equity called"] - s.loc["Total investors Equity called"])
                              + np.abs(s.loc["OCM Distributions"] + s.loc["Loan America Distributions"] - s.loc["Total investors Distributions"])))
        out["F4a"] = {"middle = F4": round(c32, 4), "OCM + LA = SPV": round(c33, 2), "status": "OK" if c32 < 1e-3 and c33 < 1 else "CHECK"}
    return out
