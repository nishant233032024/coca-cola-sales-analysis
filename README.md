# Coca-Cola U.S. Sales & Operating Profitability — 2021 Analysis

An Excel-based analysis of one year of Coca-Cola U.S. retail sales data, answering a single business
question:

> **How can Coca-Cola optimize sales and operating profitability across U.S. retailers, regions,
> products, and time?**

Built for a presentation by **Nishant**. The dataset is educational sample data supplied for the
exercise, not audited company reporting.

## Headline numbers (verified against the data)

| Metric | Value |
|---|---|
| Total sales | $8,684,027.50 |
| Units sold | 17,148,250 |
| Operating profit | $3,173,631.88 |
| Weighted operating margin | 36.55% (total profit ÷ total sales) |
| Realised sales per unit | $0.5064 |
| Operating profit per unit | $0.1851 |

Scope: 3,888 records, 2021-01-02 → 2021-12-25, 54 cities in 50 states, 4 retailers, 5 regions,
6 beverage brands. Operating margin is always the ratio of sums, never an average of row margins
(the simple average is 36.31%, 0.24 pp off).

## What is in the repository

| Path | What it is |
|---|---|
| `deliverables/Coca_Cola_Sales_Analysis.xlsx` | Main workbook: 17 sheets, 18 native PivotTables on one shared cache (including four summary panels on the dashboard), 3 native slicers, 5 pivot charts, a PivotTable heatmap, formula-driven analysis sheets and a 23-row validation sheet. |
| `deliverables/Coca_Cola_Project_Report.docx` / `.pdf` | Project report: scope, data quality, findings, recommendations, limitations. |
| `deliverables/Coca_Cola_Excel_Implementation_Guide.docx` / `.pdf` | Click-by-click rebuild guide for every view in the workbook. |
| `deliverables/Nishant_Loom_Presentation_Script.docx` / `.pdf` / `.txt` | 907-word narrated presentation script with evaluator Q&A and a recording checklist. |
| `deliverables/Reproduce_Coca_Cola_Analysis.py` | Standalone reproduction of every figure in the report (needs only `openpyxl`). |
| `work/build/` | The build pipeline that produced the workbook and documents (Python + openpyxl, then Excel COM for PivotTables, slicers and pivot charts). |
| `work/facts.json` | Every number in the workbook, report and script is generated from this file, so prose and cells cannot drift. |

## How to use

1. Open `deliverables/Coca_Cola_Sales_Analysis.xlsx` in desktop Excel (2016+). Press `Ctrl+Alt+F9` once
   if values do not recalculate on open.
2. Start at `00_ReadMe`, then `04_Data_Quality`, then `12_Dashboard`.
3. On the dashboard the three slicers (Retailer, Region, Beverage Brand) filter the KPI cards, the
   five pivot charts, the heatmap and the four summary PivotTables together, because every
   PivotTable is built on the same cache.
4. Sheets `05`–`10` are formula baselines that deliberately ignore the slicers — they are the
   auditable full-year numbers.
5. `15_Validation` lists every reconciliation with Pass / Disclosed status.

To re-verify the numbers without Excel: `python deliverables/Reproduce_Coca_Cola_Analysis.py`.

## Known limitations

- **Currency is not documented in the source.** USD is an assumption, flagged on `04_Data_Quality`.
- **The Invoice Date timeline is not automated.** Add it with `PivotTable Analyze > Insert Timeline >
  Invoice Date` (three clicks). The slicers alone cover the filtering demonstrated in the script.
- **Retailer and region are confounded**: BevCo sells only in the South, FizzySip only in the West,
  so "BevCo has the best margin" is not a like-for-like comparison.
- Two zero-sales rows (Omaha / Fanta) are retained and flagged rather than deleted.
- The build scripts under `work/build/` require Windows with Excel installed; they are development
  tooling, not part of the submission.
