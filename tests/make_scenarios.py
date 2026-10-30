"""Rebuild the five stress-test workbooks used to verify the engine (needs LibreOffice).
Usage: python tests/make_scenarios.py path/to/LoanAmerica_Model_2026_10_21_v4.xlsx
Then:  python tests/tie_out.py s1.xlsx s2.xlsx s3.xlsx s4.xlsx s5.xlsx
Recalculation uses a LibreOffice headless run; point RECALC at any script that recalculates in place."""
import sys, subprocess, openpyxl

SOURCE = sys.argv[1]
RECALC = "/mnt/skills/public/xlsx/scripts/recalc.py"
S = {
 "s1.xlsx": {"B20": "OFF", "B17": "OFF", "B16": "OFF"},
 "s2.xlsx": {"B22": "IRR", "C101": 0.7, "F101": 0.20, "B102": "ON", "C102": 0.5, "F102": 0.30,
             "B24": "100% to school", "B23": "Price only", "B42": 0.3, "B43": 3},
 "s3.xlsx": {"B16": "OFF", "B17": "ON", "B141": 24, "B158": "On top of pro rata", "B25": "Fixed $", "B129": 5000000,
             "B80": "ON", "C80": "Admin", "D80": "X", "E80": "Across life", "F80": "%", "G80": "Beginning balance", "H80": 0.01,
             "B81": "ON", "C81": "Face fee", "E81": "Upfront", "F81": "%", "G81": "Face", "H81": 0.005,
             "B82": "ON", "C82": "Fixed up", "E82": "Upfront", "F82": "Fixed $", "I82": 50,
             "B83": "ON", "C83": "Fixed life", "E83": "Across life", "F83": "Fixed $", "I83": 2,
             "B36": 5000, "C36": 8000, "D36": 12000, "E36": 15000, "B71": 0.25, "C71": 0.30, "D71": 0.35, "E71": 0.40,
             "B30": 0.1, "C30": 0.3, "D30": 0.3, "E30": 0.3},
 "s4.xlsx": {"B18": "OFF", "B19": "OFF", "B21": "OFF", "C101": 0.6, "E101": 1.5, "B102": "ON", "C102": 0.5, "E102": 2.0,
             "B119": 24, "B125": 6, "F126": 0.03, "G126": 0.04, "B17": "OFF"},
 "s5.xlsx": {"B20": "OFF", "E31": "G_short", "B40": 1.0, "B41": 1.5, "B118": 150000000, "B140": 4, "B141": 30, "B142": 0.5},
}
for name, changes in S.items():
    wb = openpyxl.load_workbook(SOURCE)
    ws = wb["F1 Fusion Assumptions"]
    for k, v in changes.items():
        ws[k] = v
    wb.save(name)
    subprocess.run([sys.executable, RECALC, name, "400"], check=False)
    print("built", name)
