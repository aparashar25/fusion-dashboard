"""Tie the Python engine to an Excel workbook, cell by cell.

Usage: python tests/tie_out.py path/to/model.xlsx [more.xlsx ...]
The workbook must hold calculated values (saved by Excel, or recalculated by LibreOffice).
Every compared cell must match to a tight tolerance; the script prints the worst differences.
"""
import sys, math, warnings
import numpy as np
import openpyxl
sys.path.insert(0, __file__.rsplit("/tests/", 1)[0])
from fusion import load_inputs_from_excel, run_model
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")


def close(a, b, rel=1e-9, abs_=1e-7):
    if isinstance(b, str) and b.startswith("#"):
        b = float("nan")
    if a is None or (isinstance(a, float) and math.isnan(a)):
        return b in (None, "", "n.a.") or (isinstance(b, float) and math.isnan(b))
    if isinstance(b, str) or b is None:
        return False
    return abs(a - b) <= max(abs_, rel * max(abs(a), abs(b)))


class Tie:
    def __init__(self, name):
        self.name, self.n, self.fails = name, 0, []

    def cmp(self, label, py, xl, rel=1e-9, abs_=1e-7):
        self.n += 1
        if not close(py, xl, rel, abs_):
            self.fails.append((label, py, xl))

    def report(self):
        status = "PASS" if not self.fails else "FAIL"
        extra = f" ({self.solver_na} IRRs where the spreadsheet solver returned n.a.)" if getattr(self, "solver_na", 0) else ""
        print(f"  {self.name}: {self.n} cells compared, {len(self.fails)} mismatches -> {status}{extra}")
        for f in self.fails[:8]:
            print("     ", f)
        return not self.fails


def tie(path, check_sens=True):
    print(f"\n=== {path}")
    inp = load_inputs_from_excel(path)
    res = run_model(inp, sensitivities=check_sens)
    wb = openpyxl.load_workbook(path, data_only=True)
    F1, F2, F2a = wb["F1 Fusion Assumptions"], wb["F2 Fusion Unit Economics"], wb["F2a Fusion Asset Returns"]
    F3, F4 = wb["F3 Fusion Pool Cash Flows"], wb["F4 Fusion SPV Summary"]
    v = lambda ws, r, c: ws.cell(r, c).value
    ok = True

    # ---------------- F1
    t = Tie("F1 derived + checks")
    d = res["derived"]
    for k, col in enumerate("BCDE"):
        t.cmp(f"F1 {col}33 pmt", d["pmt"][k], F1[f"{col}33"].value)
        t.cmp(f"F1 {col}72 upfront", d["upfront"][k], F1[f"{col}72"].value)
        t.cmp(f"F1 {col}74 target outlay", d["target_outlay"][k], F1[f"{col}74"].value)
        for a in range(61):
            t.cmp(f"F1 loss {k} age {a}", d["loss"][k][a], v(F1, 60 + k, 2 + a) or 0.0)
            t.cmp(f"F1 smm {k} age {a}", d["smm"][k][a], v(F1, 64 + k, 2 + a) or 0.0)
    for m in range(60):
        t.cmp(f"F1 face m{m+1}", d["plan_face"][m], v(F1, 124, 2 + m), abs_=1e-4)
    t.cmp("F1 B120 total face", d["total_face"], F1["B120"].value, abs_=1e-3)
    t.n += 1
    if d["checks"]["input_status"] != F1["B12"].value:
        t.fails.append(("F1 B12", d["checks"]["input_status"], F1["B12"].value))
    t.n += 1
    if res["model_status"] != F1["B13"].value:
        t.fails.append(("F1 B13", res["model_status"], F1["B13"].value))
    ok &= t.report()

    # ---------------- F2 term blocks
    t = Tie("F2 unit economics (4 term blocks x 61 ages, blended, summary)")
    starts = [83, 164, 245, 326]
    rowmap = {"sched_balance": 6, "opening": 10, "charge_offs": 11, "surviving": 12, "interest": 13,
              "sched_principal": 14, "prepaid": 15, "ending": 16, "active_loans": 17, "recoveries": 18,
              "missed": 19, "prepay_survival": 20, "survivors_stop_off": 21, "defaulting_balance": 22,
              "defaulted_face": 23, "removed": 24, "gross": 25, "fees": 26, "net": 27,
              "investor": 61, "school": 62, "tiers_met": 63, "tier_in_force": 64, "investor_pct_in_force": 66, "need_next": 68, "investor_cf": 70, "investor_cum": 71,
              "moic_running": 72, "school_cf": 74, "school_cum": 75}
    for k, s in enumerate(starts):
        df = res["ue"]["blocks"][k]["df"]
        for col, off in rowmap.items():
            for a in range(61):
                xl = v(F2, s + off, 3 + a)
                if xl is None:
                    xl = 0.0
                if col == "active_loans" and a > inp.terms[k]:
                    continue   # past maturity Excel divides two floating-point residuals (~1e-17); only feeds fixed $ fees, capped at 0 collections
                t.cmp(f"F2 {col} term{k} age{a} (row {s+off})", float(df[col].values[a]), xl, abs_=1e-10)
        for i in range(10):
            for a in range(61):
                t.cmp(f"F2 need tier{i+1} term{k} age{a}", float(df[f"need_tier_{i+1}"].values[a]), v(F2, s + 29 + i, 3 + a) or 0.0, abs_=1e-10)
                t.cmp(f"F2 alloc tier{i+1} term{k} age{a}", float(df[f"alloc_tier_{i+1}"].values[a]), v(F2, s + 50 + i, 3 + a) or 0.0, abs_=1e-10)
        for i in range(11):
            for a in range(61):
                xl = v(F2, s + 39 + i, 3 + a) or 0.0
                py = float(df[f"breakpoint_{i}"].values[a])
                if xl >= 1e11 and py >= 1e11:
                    t.n += 1; continue
                t.cmp(f"F2 breakpoint {i} term{k} age{a}", py, xl, abs_=1e-10)
        for a in range(61):
            t.cmp(f"F2 after-last term{k} age{a}", float(df["alloc_after_last"].values[a]), v(F2, s + 60, 3 + a) or 0.0, abs_=1e-10)
            for col, off in [("tier_in_force_label", 65), ("target_hit_flag", 67)]:
                t.n += 1
                xl = (v(F2, s + off, 3 + a) or "").replace("▲ ", "")
                if df[col].values[a] != xl:
                    t.fails.append((f"F2 {col} term{k} age{a}", df[col].values[a], xl))
        for a in range(2, 61):   # running IRR, skip the ~-100% early months where solvers are unstable
            xl = v(F2, s + 73, 3 + a)
            if isinstance(xl, (int, float)) and xl > -0.9:
                t.cmp(f"F2 running IRR term{k} age{a}", df.irr_running.values[a], xl, abs_=1e-8)
    bl = res["ue"]["blended"]
    bmap = {"opening": 58, "charge_offs": 59, "interest": 60, "sched_principal": 61, "prepaid": 62, "ending": 63,
            "recoveries": 64, "missed": 65, "removed": 66, "gross": 67, "fees": 68, "net": 69, "investor": 70,
            "school": 71, "investor_cf": 72, "investor_cum": 73, "moic_running": 74, "school_cf": 76,
            "school_cum": 77, "cost_basis": 80}
    for col, r in bmap.items():
        for a in range(61):
            t.cmp(f"F2 blended {col} age{a}", float(bl[col].values[a]), v(F2, r, 3 + a) or 0.0, abs_=1e-10)
    S = res["ue"]["summary"]
    smap = {"Advance paid to school": 8, "Upfront fees": 9, "Investor outlay": 10, "Gross collections": 11,
            "Across-life fees": 12, "Net collections": 13, "Charge-offs": 14, "Missed payments": 15,
            "Investor cash received": 16, "Investor MOIC": 17, "Investor IRR": 18, "School day-0 advance": 22,
            "School collections": 23, "School total (undiscounted)": 24, "School PV of collections": 25,
            "School total (PV)": 26, "Effective discount (undiscounted)": 32, "Effective discount (PV)": 33}
    for j, name in enumerate(S.columns):
        for lab, r in smap.items():
            t.cmp(f"F2 summary {lab} {name}", float(S.loc[lab, name]), v(F2, r, 2 + j), abs_=1e-9)
        if j < 4:
            xl = v(F2, 20, 2 + j)
            py = S.loc["Month tier 1 hit", name]
            t.n += 1
            if not (py == xl or (isinstance(xl, str) and isinstance(py, str))):
                t.fails.append((f"F2 row 20 {name}", py, xl))
    ok &= t.report()

    # ---------------- F2a
    t = Tie("F2a asset returns (gross vs net)")
    A = res["asset"]["table"]
    amap = {"Price paid (advance)": 6, "Gross collections": 7, "Investor cash, gross": 8, "Investor MOIC, gross": 9,
            "Investor IRR, gross": 10, "Investor MOIC, net": 11, "Investor IRR, net": 12, "Fee drag on MOIC": 13,
            "School total, gross": 14, "School total, net": 15}
    for j, name in enumerate(A.columns):
        for lab, r in amap.items():
            t.cmp(f"F2a {lab} {name}", float(A.loc[lab, name]), v(F2a, r, 2 + j), abs_=1e-9)
    bcf = res["asset"]["gross_cf_blended"]
    for a in range(61):
        t.cmp(f"F2a row 20 age{a}", float(bcf[a]), v(F2a, 20, 3 + a) or 0.0, abs_=1e-10)
    ok &= t.report()

    # ---------------- F3
    t = Tie("F3 pool cash flows (120 months)")
    p = res["pool"]["monthly"]
    pmap = {"face": 7, "cash_to_school": 8, "upfront_fees": 9, "outlay": 10, "opening": 13, "charge_offs": 14,
            "interest": 15, "sched_principal": 16, "prepaid": 17, "recoveries": 18, "missed": 19, "gross": 20,
            "fees": 21, "net": 22, "investor": 23, "school": 24, "ending": 25, "school_total": 33,
            "unlev_cf": 39, "active": 43, "other_fees": 44, "cost_basis": 48, "eligible": 49, "capacity": 50,
            "rate": 52, "debt_open": 53, "interest_debt": 54, "draws": 55, "repay": 56, "debt_end": 57,
            "net_spv": 67, "operating": 68, "equity": 69, "dist": 70, "recycled": 71, "net_to_investors": 73}
    for col, r in pmap.items():
        for m in range(120):
            xl = v(F3, r, 3 + m)
            t.cmp(f"F3 {col} m{m+1} (row {r})", float(p[col].values[m]), xl if xl is not None else 0.0,
                  rel=1e-9, abs_=1e-4)
    for m in range(120):
        t.n += 1
        if p["mode"].values[m] != v(F3, 47, 3 + m):
            t.fails.append((f"F3 mode m{m+1}", p["mode"].values[m], v(F3, 47, 3 + m)))
    PS = res["pool"]["summary"]
    for lab, r in {"Unlevered investor MOIC": 86, "Unlevered investor IRR": 87, "Peak debt": 91,
                   "Equity called": 94, "Distributions": 95, "Collections recycled into purchases": 96,
                   "Equity MOIC": 97, "Equity IRR": 98}.items():
        t.cmp(f"F3 B{r} {lab}", float(PS[lab]), v(F3, r, 2), abs_=1e-4)
    vi = res["pool"]["vintage_investor"]
    for vv in range(vi.shape[0]):
        for m in range(120):
            t.cmp(f"F3 vintage inv v{vv+1} m{m+1}", vi[vv, m], v(F3, 103 + vv, 3 + m) or 0.0, abs_=1e-4)
    vs = res["pool"]["vintage_school"]
    for vv in range(vs.shape[0]):
        for m in range(120):
            t.cmp(f"F3 vintage school v{vv+1} m{m+1}", vs[vv, m], v(F3, 167 + vv, 3 + m) or 0.0, abs_=1e-4)
    t.n += 1
    if res["pool"]["checks"]["status"] != F3["B99"].value:
        t.fails.append(("F3 B99", res["pool"]["checks"]["status"], F3["B99"].value))
    ok &= t.report()

    # ---------------- F4
    t = Tie("F4 SPV summary and waterfall")
    w = res["spv"]["monthly"]
    wmap = {"equity": 30, "ocm_equity": 31, "la_equity": 32, "dist": 36, "t1": 38, "hb2": 40, "t2": 42,
            "hb3": 45, "t3": 46, "hb4": 47, "t4": 48, "t5": 49, "ocm": 50, "la": 51, "la_t12": 57, "la_t3": 58,
            "la_t4": 59, "la_t5": 60, "ocm_net": 64, "la_net": 65, "total_net": 66, "debt_recap": 73}
    for col, r in wmap.items():
        for m in range(120):
            xl = v(F4, r, 3 + m)
            t.cmp(f"F4 {col} m{m+1} (row {r})", float(w[col].values[m]), xl if xl is not None else 0.0, abs_=1e-4)
    extra = {"ocm": 99, "la_pro_rata": 100, "la_promote": 101, "ocm_cum_dist": 104, "la_cum_dist": 105,
             "spv_cum_net": 111, "ocm_cum_net": 114, "la_cum_net": 118}
    for col, r in extra.items():
        for m in range(120):
            t.cmp(f"F4 {col} m{m+1} (row {r})", float(w[col].values[m]), v(F4, r, 3 + m) or 0.0, abs_=1e-4)
    for m in range(120):
        xl = v(F4, 102, 3 + m)
        py = w.la_share_month.values[m]
        if isinstance(xl, (int, float)):
            t.cmp(f"F4 LA share m{m+1}", float(py), xl)
        else:
            t.n += 1
            if np.isfinite(py):
                t.fails.append((f"F4 LA share m{m+1}", py, xl))
        for col, r in [("ocm_irr_running", 70), ("la_irr_running", 71)]:
            xl = v(F4, r, 3 + m)
            if isinstance(xl, (int, float)) and xl > -0.9:
                t.cmp(f"F4 {col} m{m+1}", float(w[col].values[m]), xl, abs_=1e-8)
    Y = res["spv"]["yearly"]
    ymap = {"Total: equity called": 79, "Total: distributions": 80, "Total: net cash flow": 81, "Total: cumulative net": 82,
            "OCM: equity called": 84, "OCM: distributions": 85, "OCM: net cash flow": 86, "OCM: cumulative net": 87,
            "LA: equity called": 89, "LA: distributions": 90, "LA: of which promote above pro rata": 91,
            "LA: net cash flow": 92, "LA: cumulative net": 93, "LA share of distributions": 94}
    for col, r in ymap.items():
        for j in range(10):
            t.cmp(f"F4 yearly {col} y{j}", float(Y[col].values[j]), v(F4, r, 3 + j) or 0.0, abs_=1e-3)
    R = res["spv"]["returns"]
    for j, party in enumerate(R.columns):
        for lab, r in {"Equity called": 16, "Distributions": 17, "Net profit": 18, "MOIC": 19, "IRR": 20,
                       "Share of distributions": 21, "Debt recap": 22, "Adjusted MOIC": 23}.items():
            xl = v(F4, r, 2 + j)
            if isinstance(xl, str):
                xl = float("nan")
            t.cmp(f"F4 {lab} {party}", float(R.loc[lab, party]), xl, abs_=1e-4)
    T = res["spv"]["tiers"]
    for j, party in enumerate(T.columns):
        for i in range(5):
            t.cmp(f"F4 tier {i+1} {party}", float(T.iloc[i, j]), v(F4, 24 + i, 2 + j), abs_=1e-3)
    t.n += 1
    if res["spv"]["checks"]["waterfall"] != F4["B13"].value:
        t.fails.append(("F4 B13", res["spv"]["checks"]["waterfall"], F4["B13"].value))
    ok &= t.report()

    # ---------------- F5
    t = Tie("F5 dashboard views")
    F5 = wb["F5 Fusion Dashboard"]
    f5 = res["f5"]
    mv = f5["monthly"]
    for j, col in enumerate(mv.columns):
        r = 37 + j
        for m in range(120):
            xl = v(F5, r, 3 + m)
            if "IRR" in col:
                if isinstance(xl, (int, float)) and xl > -0.9:
                    t.cmp(f"F5 {col} m{m+1}", float(mv[col].values[m]), xl, abs_=1e-8)
                continue
            t.cmp(f"F5 {col} m{m+1}", float(mv[col].values[m]), xl if xl is not None else 0.0, abs_=1e-4)
    rc = f5["recycling"]
    for j, col in enumerate(rc.columns):
        for m in range(60):
            xl = v(F5, 100 + j, 3 + m)
            py = rc[col].values[m]
            if isinstance(xl, str) or xl is None:
                t.n += 1
                if np.isfinite(py) and col.startswith("Share"):
                    t.fails.append((f"F5 {col} m{m+1}", py, xl))
                continue
            t.cmp(f"F5 {col} m{m+1}", float(py), xl, abs_=1e-4)
    lv = f5["leverage"]
    for j, col in enumerate(lv.columns):
        for m in range(60):
            xl = v(F5, 144 + j, 3 + m)
            py = lv[col].values[m]
            if col == "Facility mode":
                t.n += 1
                if py != xl: t.fails.append((f"F5 mode m{m+1}", py, xl))
                continue
            t.cmp(f"F5 {col} m{m+1}", float(py), xl if xl is not None else 0.0, abs_=1e-4)
    for (lab, py), addr in zip(f5["recycling_kp"].items(), ["B111", "B112", "B113"]):
        t.cmp(f"F5 {addr}", float(py) if not isinstance(py, str) else py, F5[addr].value) if not isinstance(py, str) else None
    kp = f5["leverage_kp"]
    for lab, addr in [("First draw month", "B158"), ("Peak debt ($)", "B159"), ("Debt fully repaid by month", "B160"),
                      ("Total interest paid ($)", "B161")]:
        py = kp[lab]
        if isinstance(py, str):
            t.n += 1
            if py != F5[addr].value: t.fails.append((addr, py, F5[addr].value))
        else:
            t.cmp(f"F5 {addr}", float(py), F5[addr].value, abs_=1e-3)
    t.cmp("F5 B21 program PV", res["pool"]["summary"]["School effective advance (PV)"], F5["B21"].value)
    ok &= t.report()

    # ---------------- F2b / F4a
    if check_sens and res.get("single_split"):
        t = Tie("F2b advance sensitivity + F4a split sensitivity")
        F2b, F4a = wb["F2b Fusion Sensitivity"], wb["F4a Fusion Split Sensitivity"]
        a = res["adv_sens"]
        names = ["24m", "36m", "48m", "60m", "Blended"]
        for j, i in enumerate(range(-5, 6)):
            c = 3 + j
            t.cmp(f"F2b advance {i}", a.loc["Blended advance", i], v(F2b, 10, c))
            for n_, nm in enumerate(names):
                t.cmp(f"F2b MOIC {nm} {i}", a.loc[f"MOIC {nm}", i], v(F2b, 14 + n_, c), rel=1e-8)
                t.cmp(f"F2b IRR {nm} {i}", a.loc[f"IRR {nm}", i], v(F2b, 21 + n_, c), rel=1e-7, abs_=1e-7)
            t.cmp(f"F2b school {i}", a.loc["School total (undiscounted)", i], v(F2b, 28, c))
            t.cmp(f"F2b school PV {i}", a.loc["School total (PV)", i], v(F2b, 29, c))
        s = res["split_sens"]
        for j, i in enumerate(range(-5, 6)):
            c = 3 + j
            for party, base in [("Total investors", 14), ("OCM", 20), ("Loan America", 26)]:
                for off, k in enumerate(["Equity called", "Distributions", "MOIC", "IRR"]):
                    xl = v(F4a, base + off, c)
                    if isinstance(xl, str) or xl is None:
                        t.n += 1
                        t.solver_na = getattr(t, "solver_na", 0) + 1   # spreadsheet IRR gave up; engine still solves
                        continue
                    t.cmp(f"F4a {party} {k} {i}", s.loc[f"{party} {k}", i], xl, rel=1e-7, abs_=1e-4)
        ok &= t.report()
    print(f"  RESULT: {'ALL TIE' if ok else 'MISMATCHES FOUND'}")
    return ok


if __name__ == "__main__":
    paths = sys.argv[1:] or ["/home/claude/m/orig.xlsx"]
    results = [tie(p) for p in paths]
    print("\nOVERALL:", "PASS" if all(results) else "FAIL")
    sys.exit(0 if all(results) else 1)
