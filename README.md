# Fusion Loans deal dashboard

A Python rebuild of `LoanAmerica_Model_2026_10_21_v4.xlsx` (tabs F1 to F6) with an interactive web dashboard.
All of the math runs in Python (`fusion/`); the dashboard (`app.py`) only displays it.

## Run it

```
pip install -r requirements.txt
streamlit run app.py
```

Your browser opens at http://localhost:8501. The base case loads automatically (`scenarios/base_case_v4.json`).
To start from a different version of the workbook, upload the .xlsx in the sidebar: every F1 input is read in.

## Light and dark mode

The app follows your computer's light or dark setting. To switch by hand: menu (top right) > Settings > Choose app theme.
Tables, charts and cards all adapt.

## Pages

| Menu | Page | Matches Excel tab | What is on it |
|---|---|---|---|
| Deal model | Overview | F0 | Party returns, capital and cash, asset level, school, deal settings and status, 2 key charts |
| Deal model | Assumptions | F1 | Every input and the calculated values next to them, all checks |
| Deal model | Loan economics | F2 | Per $1 summary, tier summary, blended block, the full monthly engine for each term |
| Deal model | Gross vs net returns | F2a | Gross vs net investor and school returns, blended gross cash flow |
| Deal model | Pool cash flows | F3 | Purchases, pool cash flows, school, unlevered investor, SPV fees, leverage, recycling, summary, vintage grids |
| Deal model | SPV and waterfall | F4 | Capital split, returns by party, the monthly waterfall, Loan America by tier, running returns, yearly view |
| Deal model | Program detail | F5 | Returns by party incl. school and Goal, school advance rate, monthly view, recycling and leverage views |
| Deal model | Charts | F6 | All 16 charts in 7 sections |
| Sensitivities | Advance rate | F2b | Investor MOIC / IRR and school take vs the advance rate |
| Sensitivities | Investor split | F4a | SPV / OCM / Loan America returns vs the investor split |
| Tools | Checks | B12 / B13 and every tab's checks | Every check on one page |
| Tools | Compare scenarios | (new) | Saved scenarios side by side |
| Tools | Scenario notes | Sheet3 | Your scenario table (editable) |
| Tools | Model notes | Claude Handoff | The model's documentation and change log |

Extra charts beyond the Excel sit at the bottom of a page under "Additional views (not in the workbook)".

## Code layout

```
fusion/inputs.py   all F1 inputs as one object; Excel loader; JSON save / load
fusion/engine.py   derive (F1 formulas + checks), term_block + split_engine (F2), asset_returns (F2a),
                   pool (F3), spv (F4), advance_sensitivity (F2b), split_sensitivity (F4a), run_model
fusion/irr.py      IRR and NPV that behave like Excel's IRR(values, 0.01) and NPV
app.py, ui_kit.py  Streamlit pages and styling
charts.py          the 18 workbook charts (2 on F0, 16 on F6), series for series
notes/             the Claude Handoff text and Sheet3 table from the workbook (re-read when you upload a workbook)
fusion/excel_labels.json  the workbook's row labels, so the pages use the same wording
tests/tie_out.py   cell-by-cell comparison of the engine against any calculated copy of the workbook
```

Comments in `engine.py` cite the Excel row each line reproduces.

## How it was verified

`python tests/tie_out.py model.xlsx` loads the workbook's inputs, runs the engine, and compares about 38,000
cells per workbook: F1 derived values and checks; all four F2 term blocks for ages 0 to 60 including every tier
helper row, the blended block and summary; F2a including the blended gross cash flow; all 120 months of F3
including both vintage grids; the F4 waterfall month by month, running IRRs, yearly view and chart helper rows;
the F5 monthly, recycling and leverage views and key points; and the F2b / F4a grids.

It was run on the v4 base case and five stress scenarios (payment stop / leverage / recycling off; IRR mode with
two tiers and 100% to school after the last tier, recoveries with a lag; recycling off with leverage on, on-top
promote, fixed $ capital, balance-based and fixed $ fees, uneven loan sizes, advance rates and mix; losses and
prepay off with promote off and a 24-month program; payment stop off with a curve-check failure). Every cell
matched. `tests/make_scenarios.py` rebuilds those workbooks (needs LibreOffice).

Two known, harmless differences:
- F2 row 100 (active loans) past a loan's maturity: Excel divides two ~1e-17 rounding residuals; Python returns 0.
  The row only feeds fixed $ per-loan fees, which are capped at that month's (zero) collections.
- In a few deep-loss sensitivity columns the spreadsheet IRR returns n.a.; the Python solver still finds the root.

## Where to host it

- Now: run locally (above). Deal data never leaves your machine.
- For colleagues: a small internal server or VM running `streamlit run app.py`, behind the company login / VPN.
  Avoid public hosting (for example Streamlit Community Cloud) for deal inputs.

## Changing the model

Change the engine first, then rerun `tests/tie_out.py` against a workbook with the same change made in Excel.
