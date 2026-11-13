# Coca-Cola Sales Analysis Using Excel — submission pack

**Presenter:** Nishant
**Source:** `Coke-sales-analysis-stubs (1).xlsx`, sheet `Data` (headers on row 5, records on rows 6–3893)
**Scope:** all six supplied beverage brands, calendar 2021, 3,888 records, USD (assumed — the source documents no currency)
**Status:** educational assignment dataset, not audited Coca-Cola company reporting

## Files

| File | What it is |
|---|---|
| `Coca_Cola_Sales_Analysis.xlsx` | The deliverable workbook: 17 sheets, 18 native PivotTables on one shared cache (including four summary panels on the dashboard), 3 native slicers, 5 pivot charts, a PivotTable heatmap, formula-driven analysis sheets and a validation sheet. |
| `Coca_Cola_Project_Report.docx` / `.pdf` | Concise project report (executive summary → findings → recommendations → limitations). |
| `Coca_Cola_Excel_Implementation_Guide.docx` / `.pdf` | Click-by-click rebuild guide, 11 steps × 12 fields, plus slicer setup and optional Power Query M code. |
| `Nishant_Loom_Presentation_Script.docx` / `.pdf` / `.txt` | 8-block, 907-word narrated script, five evaluator Q&A, recording checklist. |
| `Reproduce_Coca_Cola_Analysis.py` | Standalone reproduction of every figure (requires only `openpyxl`). |

## Headline numbers (verified)

| Metric | Value |
|---|---|
| Total sales | $8,684,028 |
| Units sold | 17,148,250 |
| Operating profit | $3,173,632 |
| Weighted operating margin | 36.55% (total profit ÷ total sales) |
| Realised sales per unit | $0.5064 |
| Operating profit per unit | $0.1851 |

Simple average of the row-level margin column: 36.31% — 0.24 pp away from the ratio of sums, which is why it is not used.

## How to open and use

1. Open `Coca_Cola_Sales_Analysis.xlsx` in desktop Excel (2016+). Press `Ctrl+Alt+F9` once if values do not recalculate on open.
2. Start at `00_ReadMe`, then `04_Data_Quality`, then `12_Dashboard`.
3. On the dashboard, the three slicers at the right edge (Retailer, Region, Beverage Brand) filter the five pivot charts, the four summary PivotTables below them and the KPI cards together, because all PivotTables share one cache.
4. `05`–`10` are formula baselines that deliberately ignore slicers — they are the full-year auditable numbers.
5. `15_Validation` lists every reconciliation with Pass/Fail.

## What was verified

- Raw rows = cleaned rows (3,888); raw sales/profit/units = cleaned totals.
- `Price × Units = Total Sales` in all 3,888 rows; `Profit ÷ Sales = supplied margin` in every row with sales.
- No duplicate retailer-location-brand-month keys, no negatives, no missing dates, no partial months (324 records each).
- Retailer, region, brand and month blocks each sum to the same portfolio total.
- Slicer test: selecting Sodapop changes the KPI cards to $4,403,630 and 37.42% and cuts the dashboard retailer PivotTable down to the Sodapop row, matching the independent calculation; clearing returns $8,684,028.
- 22 of 23 validation rows show Pass; the 23rd is intentionally "Disclosed" (currency is not documented in the source).

## What was not verified / not automated

- The **Invoice Date timeline** is not in the delivered file. Excel automation could not create it through this COM binding; `PivotTable Analyze > Insert Timeline > Invoice Date` adds it in three clicks. The slicers alone fully cover the filtering demonstrated in the script.
- Charts were exported and visually checked individually; the **full dashboard was not viewed as a single composed page** at multiple screen sizes. Slicer positions are fixed and may need a nudge on very small displays.
- The workbook contains 2 intentional `#N/A` cells (`Recalculated Margin` on the two zero-sales Omaha rows). They are flagged in `Validation Status` and documented in `04_Data_Quality`.
- There is **no prior-year data**, so no year-on-year comparison and no forecast is presented as actual.

## Rebuild from scratch

```bash
python Reproduce_Coca_Cola_Analysis.py --input "Coke-sales-analysis-stubs (1).xlsx"
```

The workbook's build scripts (`work/build/`) require Windows Excel for the PivotTable, slicer and chart stage; they are development tooling, not part of the submission.
