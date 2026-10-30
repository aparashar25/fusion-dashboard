"""Build the investor workbook for a scenario, recalculate it with LibreOffice, and tie every formula back to the Python engine.
Usage: python tests/excel_roundtrip.py [inputs.json | model.xlsx ...]"""
import sys, os, math, shutil, subprocess, tempfile, warnings
import numpy as np, openpyxl
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from fusion import Inputs, load_inputs_from_excel, run_model
from fusion.excel_export import build_book, AS, LE, GN, PL, SP, PD, SA, SS, CK, OV, ac, mc
warnings.filterwarnings("ignore")
RECALC = "/mnt/skills/public/xlsx/scripts/recalc.py"


def close(a, b, rel=1e-9, abs_=1e-7):
    if isinstance(b, str) or b is None:
        return (a is None or (isinstance(a, float) and math.isnan(a))) and b in (None, "", "n.a.")
    if a is None or (isinstance(a, float) and math.isnan(a)):
        return False
    return abs(a - b) <= max(abs_, rel * max(abs(a), abs(b)))


class T:
    def __init__(s, name): s.name, s.n, s.f = name, 0, []
    def cmp(s, lab, py, xl, rel=1e-9, abs_=1e-7):
        s.n += 1
        if not close(py, xl, rel, abs_): s.f.append((lab, py, xl))
    def eq(s, lab, py, xl):
        s.n += 1
        if py != xl: s.f.append((lab, py, xl))
    def rep(s):
        print(f"  {s.name}: {s.n} cells, {len(s.f)} mismatches -> {'PASS' if not s.f else 'FAIL'}")
        for x in s.f[:6]: print("     ", x)
        return not s.f


def roundtrip(inp, tag="scenario"):
    print(f"\n=== {tag}")
    B = build_book(inp)
    d = tempfile.mkdtemp()
    p = os.path.join(d, "wb.xlsx")
    B.wb.save(p)
    out = subprocess.run([sys.executable, RECALC, p, "400"], capture_output=True, text=True).stdout
    if '"total_errors": 0' not in out:
        print("  recalc:", out[:400])
    wb = openpyxl.load_workbook(p, data_only=True)
    res = run_model(inp, sensitivities=True)
    ok = True
    v = lambda sh, r, c: wb[sh].cell(r, c).value
    # ---------------- Loan economics
    t = T("Loan economics")
    blocks = B.ref["LE"]["blocks"]
    pairs = {"S": "sched_balance", "fac": "sched_factor", "loss": "loss_curve", "smm": "smm", "open": "opening", "co": "charge_offs",
             "surv": "surviving", "int": "interest", "sch": "sched_principal", "pre": "prepaid", "end": "ending", "rec": "recoveries",
             "missed": "missed", "pps": "prepay_survival", "s104": "survivors_stop_off", "defbal": "defaulting_balance",
             "defface": "defaulted_face", "removed": "removed", "gross": "gross", "fees": "fees", "net": "net",
             "sp_inv": "investor", "sp_school": "school", "sp_met": "tiers_met", "sp_tif": "tier_in_force",
             "sp_pct": "investor_pct_in_force", "cf": "investor_cf", "cum": "investor_cum", "moic": "moic_running",
             "scf": "school_cf", "scum": "school_cum"}
    for k in range(4):
        df = res["ue"]["blocks"][k]["df"]
        for xr, col in pairs.items():
            for a in range(61):
                t.cmp(f"{col} t{k} a{a}", float(df[col].values[a]), v(LE, blocks[k][xr], 3 + a) or 0.0, abs_=1e-10)
        for a in range(2, 61):
            xl = v(LE, blocks[k]["irr"], 3 + a)
            if isinstance(xl, (int, float)) and xl > -0.9:
                t.cmp(f"irr t{k} a{a}", float(df.irr_running.values[a]), xl, abs_=1e-8)
        for i in range(10):
            for a in range(61):
                t.cmp(f"need{i} t{k} a{a}", float(df[f"need_tier_{i+1}"].values[a]), v(LE, blocks[k]["sp_need"] + i, 3 + a) or 0.0, abs_=1e-10)
                t.cmp(f"alloc{i} t{k} a{a}", float(df[f"alloc_tier_{i+1}"].values[a]), v(LE, blocks[k]["sp_alloc"] + i, 3 + a) or 0.0, abs_=1e-10)
    bl, bld = B.ref["LE"]["bl"], res["ue"]["blended"]
    for xr, col in {"open": "opening", "co": "charge_offs", "net": "net", "inv": "investor", "school": "school", "cf": "investor_cf",
                    "moic": "moic_running", "scf": "school_cf", "cost": "cost_basis"}.items():
        for a in range(61):
            t.cmp(f"blended {col} a{a}", float(bld[col].values[a]), v(LE, bl[xr], 3 + a) or 0.0, abs_=1e-10)
    S, Sx = res["ue"]["summary"], B.ref["LE_sum"]
    for j, nm in enumerate(S.columns):
        for lab, xr in {"Investor MOIC": "moic", "Investor IRR": "irr", "Net collections": "net", "Charge-offs": "co",
                        "Investor cash received": "inv", "School total (PV)": "s_pv", "School total (undiscounted)": "s_tot",
                        "Effective discount (PV)": "s_discpv"}.items():
            t.cmp(f"summary {lab} {nm}", float(S.loc[lab, nm]), v(LE, Sx[xr], 2 + j), abs_=1e-9)
        t.eq(f"summary checks {nm}", "OK", v(LE, Sx["chk"], 2 + j))
    ok &= t.rep()
    # ---------------- Gross vs net
    t = T("Gross vs net")
    A = res["asset"]["table"]
    for j, nm in enumerate(A.columns):
        for i, lab in enumerate(["Mix weight", "Price paid (advance)", "Gross collections", "Investor cash, gross", "Investor MOIC, gross",
                                 "Investor IRR, gross", "Investor MOIC, net", "Investor IRR, net", "Fee drag on MOIC", "School total, gross",
                                 "School total, net"]):
            t.cmp(f"{lab} {nm}", float(A.loc[lab, nm]), v(GN, 6 + i, 2 + j), abs_=1e-9)
    t.eq("check", "OK", v(GN, 19, 2))
    ok &= t.rep()
    # ---------------- assumptions checks
    t = T("Assumptions checks")
    t.eq("input status", res["input_status"], v(AS, 151, 2))
    ok &= t.rep()
    ok &= phase2(B, wb, res, v)
    ok &= phase3(B, wb, res, v)
    extra = globals().get("EXTRA")
    if extra:
        ok &= extra(B, wb, res, v)
    shutil.rmtree(d, ignore_errors=True)
    print("  RESULT:", "ALL TIE" if ok else "MISMATCHES")
    return ok


def phase2(B, wb, res, v):
    ok = True
    p = res["pool"]["monthly"]
    t = T("Pool cash flows")
    rows = {8: "face", 9: "cash_to_school", 10: "upfront_fees", 11: "outlay", 14: "opening", 15: "charge_offs", 16: "interest",
            17: "sched_principal", 18: "prepaid", 19: "recoveries", 20: "missed", 21: "gross", 22: "fees", 23: "net", 24: "investor",
            25: "school", 26: "ending", 27: "cost_basis", 28: "removed", 36: "school_total", 42: "unlev_cf", 46: "active", 47: "other_fees",
            53: "capacity", 55: "rate", 56: "debt_open", 57: "interest_debt", 58: "draws", 59: "repay", 60: "debt_end", 69: "net_spv",
            70: "operating", 71: "equity", 72: "dist", 73: "recycled", 75: "net_to_investors", 76: "fund_recycled", 77: "fund_debt", 78: "fund_equity"}
    for r, col in rows.items():
        for m in range(120):
            t.cmp(f"{col} m{m+1}", float(p[col].values[m]), v(PL, r, 3 + m) or 0.0, abs_=1e-4)
    for m in range(120):
        t.eq(f"mode m{m+1}", p["mode"].values[m], v(PL, 50, 3 + m))
    PS = res["pool"]["summary"]
    for lab, r in {"Unlevered investor MOIC": 91, "Unlevered investor IRR": 92, "Peak debt": 96, "Equity called": 99, "Distributions": 100,
                   "Equity MOIC": 102, "Equity IRR": 103, "School effective advance (PV)": 95, "Charge-offs (% of face)": 90}.items():
        t.cmp(lab, float(PS[lab]), v(PL, r, 2), abs_=1e-4)
    t.eq("pool check", res["pool"]["checks"]["status"], v(PL, 104, 2))
    vi, vs = res["pool"]["vintage_investor"], res["pool"]["vintage_school"]
    for vv in range(vi.shape[0]):
        for m in range(120):
            t.cmp(f"vin inv {vv} {m}", vi[vv, m], v(PL, 110 + vv, 3 + m) or 0.0, abs_=1e-4)
            t.cmp(f"vin sch {vv} {m}", vs[vv, m], v(PL, 175 + vv, 3 + m) or 0.0, abs_=1e-4)
    ok &= t.rep()
    t = T("SPV and waterfall")
    w = res["spv"]["monthly"]
    rows = {40: "equity", 41: "ocm_equity", 42: "la_equity", 45: "dist", 46: "t1", 48: "hb2", 51: "t2", 53: "hb3", 54: "t3", 55: "hb4",
            56: "t4", 57: "t5", 58: "ocm", 59: "la", 64: "la_t12", 65: "la_t3", 66: "la_t4", 67: "la_t5", 70: "ocm_net", 71: "la_net",
            72: "total_net", 74: "ocm_moic_running", 75: "la_moic_running", 76: "total_moic_running", 80: "debt_recap",
            84: "la_pro_rata", 85: "la_promote", 87: "ocm_cum_dist", 88: "la_cum_dist", 93: "spv_cum_net", 96: "ocm_cum_net", 100: "la_cum_net"}
    for r, col in rows.items():
        for m in range(120):
            t.cmp(f"{col} m{m+1}", float(w[col].values[m]), v(SP, r, 3 + m) or 0.0, abs_=1e-4)
    for m in range(120):
        t.eq(f"tier m{m+1}", int(w.highest_tier.values[m]), int(v(SP, 62, 3 + m)))
        for col, r in [("ocm_irr_running", 77), ("la_irr_running", 78), ("total_irr_running", 79)]:
            xl = v(SP, r, 3 + m)
            if isinstance(xl, (int, float)) and xl > -0.9:
                t.cmp(f"{col} m{m+1}", float(w[col].values[m]), xl, abs_=1e-8)
        ls, xl = w.la_share_month.values[m], v(SP, 86, 3 + m)
        if isinstance(xl, (int, float)):
            t.cmp(f"la share m{m+1}", float(ls), xl)
        else:
            t.n += 1
            if np.isfinite(ls): t.f.append((f"la share m{m+1}", ls, xl))
    R = res["spv"]["returns"]
    for j, party in enumerate(R.columns):
        for lab, r in {"Equity called": 17, "Distributions": 18, "Net profit": 19, "MOIC": 20, "IRR": 21, "Share of capital": 22,
                       "Share of distributions": 23, "Debt recap": 24, "Adjusted MOIC": 25}.items():
            t.cmp(f"{lab} {party}", float(R.loc[lab, party]), v(SP, r, 2 + j), abs_=1e-4)
    Tt = res["spv"]["tiers"]
    for j in range(3):
        for i in range(5):
            t.cmp(f"tier {i} {j}", float(Tt.iloc[i, j]), v(SP, 30 + i, 2 + j), abs_=1e-3)
    Y = res["spv"]["yearly"]
    ymap = {"Total: equity called": 104, "Total: distributions": 105, "Total: net cash flow": 106, "Total: cumulative net": 107,
            "OCM: equity called": 108, "OCM: distributions": 109, "OCM: net cash flow": 110, "OCM: cumulative net": 111,
            "LA: equity called": 112, "LA: distributions": 113, "LA: of which promote above pro rata": 114, "LA: net cash flow": 115,
            "LA: cumulative net": 116, "LA share of distributions": 117}
    for col, r in ymap.items():
        for j in range(10):
            t.cmp(f"yearly {col} {j}", float(Y[col].values[j]), v(SP, r, 3 + j) or 0.0, abs_=1e-3)
    t.eq("waterfall check", res["spv"]["checks"]["waterfall"], v(SP, 13, 2))
    ok &= t.rep()
    t = T("Program detail")
    f5 = res["f5"]
    P = f5["parties"]
    for j, party in enumerate(["OCM", "Loan America", "Total investors"]):
        for i, r in enumerate([6, 7, 8, 9, 10]):
            t.cmp(f"party {party} {i}", float(P[party].values[i]), v(PD, r, 2 + j), abs_=1e-4)
    t.cmp("school cash", float(P["School"].values[1]), v(PD, 7, 5), abs_=1e-4)
    t.cmp("goal", float(P["Goal (servicer)"].values[1]), v(PD, 7, 6), abs_=1e-4)
    t.eq("settings", f5["settings"], v(PD, 13, 2))
    Sch = f5["school"]
    for i, r in enumerate(range(17, 23)):
        for j, col in enumerate(Sch.columns):
            t.cmp(f"school {r} {col}", float(Sch[col].values[i]), v(PD, r, 2 + j), abs_=1e-6)
    mv = f5["monthly"]
    for j, col in enumerate(mv.columns):
        for m in range(120):
            xl = v(PD, 34 + j, 3 + m)
            if "IRR" in col:
                if isinstance(xl, (int, float)) and xl > -0.9:
                    t.cmp(f"{col} {m}", float(mv[col].values[m]), xl, abs_=1e-8)
                continue
            t.cmp(f"{col} {m}", float(mv[col].values[m]), xl or 0.0, abs_=1e-4)
    rc = f5["recycling"]
    for j, col in enumerate(rc.columns):
        for m in range(60):
            xl = v(PD, 50 + j, 3 + m); py = rc[col].values[m]
            if isinstance(xl, str) or xl is None:
                t.n += 1
                if col.startswith("Share") and np.isfinite(py): t.f.append((col, m, py, xl))
                continue
            t.cmp(f"{col} {m}", float(py), xl, abs_=1e-4)
    kp = f5["recycling_kp"]
    vals = list(kp.values())
    t.eq("first50", vals[0], v(PD, 62, 2)); t.eq("first100", vals[1], v(PD, 63, 2)); t.cmp("eq/purch", vals[2], v(PD, 64, 2))
    lv = f5["leverage_kp"]
    t.eq("first draw", lv["First draw month"], v(PD, 82, 2)); t.cmp("peak", lv["Peak debt ($)"], v(PD, 83, 2), abs_=1e-3)
    t.eq("repaid", lv["Debt fully repaid by month"], v(PD, 85, 2)); t.cmp("interest", lv["Total interest paid ($)"], v(PD, 86, 2), abs_=1e-3)
    ok &= t.rep()
    return ok


def phase3(B, wb, res, v):
    ok = True
    t = T("Sensitivities, checks, overview")
    if res.get("single_split"):
        a = res["adv_sens"]
        names = list(res["ue"]["summary"].columns)
        for j, i in enumerate(range(-5, 6)):
            t.cmp(f"adv {i}", a.loc["Blended advance", i], v(SA, 10, 3 + j))
            for n_, nm in enumerate(names):
                xm, xi = v(SA, 14 + n_, 3 + j), v(SA, 20 + n_, 3 + j)
                pm, pi = a.loc[f"MOIC {nm}", i], a.loc[f"IRR {nm}", i]
                if isinstance(xm, str) or xm is None:
                    t.n += 1
                    if not (isinstance(pm, float) and np.isnan(pm)): t.f.append((f"moic {nm} {i}", pm, xm))
                else:
                    t.cmp(f"moic {nm} {i}", pm, xm, rel=1e-8)
                if isinstance(xi, str) or xi is None:
                    t.n += 1
                    if not (isinstance(pi, float) and np.isnan(pi)): t.f.append((f"irr {nm} {i}", pi, xi))
                else:
                    t.cmp(f"irr {nm} {i}", pi, xi, rel=1e-7, abs_=1e-7)
            t.cmp(f"school {i}", a.loc["School total (undiscounted)", i], v(SA, 26, 3 + j))
            t.cmp(f"school pv {i}", a.loc["School total (PV)", i], v(SA, 27, 3 + j))
        s = res["split_sens"]
        for j, i in enumerate(range(-5, 6)):
            for party, r0 in [("Total investors", 13), ("OCM", 19), ("Loan America", 25)]:
                for off, k in enumerate(["Equity called", "Distributions", "MOIC", "IRR"]):
                    xl = v(SS, r0 + off, 3 + j)
                    if isinstance(xl, str) or xl is None:
                        t.n += 1; continue
                    t.cmp(f"split {party} {k} {i}", float(s.loc[f"{party} {k}", i]), xl, rel=1e-7, abs_=1e-4)
    t.eq("model status", res["model_status"], next(v(CK, r, 2) for r in range(5, 60) if str(v(CK, r, 1)).startswith("Model status")))
    t.eq("input status", res["input_status"], next(v(CK, r, 2) for r in range(5, 60) if str(v(CK, r, 1)).startswith("Input status")))
    R = res["spv"]["returns"]
    for i, party in enumerate(R.columns):
        t.cmp(f"overview moic {party}", float(R.loc["MOIC", party]), v(OV, 7 + i, 2))
        t.cmp(f"overview irr {party}", float(R.loc["IRR", party]), v(OV, 7 + i, 3))
    ok &= t.rep()
    return ok


if __name__ == "__main__":
    args = sys.argv[1:] or [os.path.join(os.path.dirname(__file__), "..", "scenarios", "base_case_v4.json")]
    allok = True
    for a in args:
        inp = Inputs.from_json(open(a).read()) if a.endswith(".json") else load_inputs_from_excel(a)
        allok &= roundtrip(inp, os.path.basename(a))
    print("\nOVERALL:", "PASS" if allok else "FAIL")
    sys.exit(0 if allok else 1)
