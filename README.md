# Coca-Cola U.S. Sales & Operating Profitability — 2021 Analysis

An Excel business-intelligence project on one year of U.S. Coca-Cola retail sales data, built to
answer a single commercial question:

> **How can Coca-Cola optimize sales and operating profitability across U.S. retailers, regions,
> products and time?**

Presented by **Nishant**. The dataset is an educational assignment dataset — this is **not** audited
Coca-Cola company reporting and is not affiliated with or endorsed by The Coca-Cola Company.

---

## Headline results

Verified against the source workbook; every figure recomputes from `deliverables/Reproduce_Coca_Cola_Analysis.py`.

| Metric | Value |
|---|---|
| Total sales | **$8,684,027.50** |
| Units sold | **17,148,250** |
| Operating profit | **$3,173,631.88** |
| Weighted operating margin | **36.55%** (total profit ÷ total sales) |
| Realised sales per unit | **$0.5064** |
| Operating profit per unit | **$0.1851** |

Scope: **3,888 records**, January–December 2021, **4 retailers**, **5 regions**, **50 states**,
**54 city markets**, **6 beverage brands**. The simple average of the row-level margin column is
36.31% — 0.24 pp away from the ratio of sums, which is exactly why it is not used anywhere.

---

## Five verified findings

1. **Margin is a retailer story, not a regional one.** FizzySip carries 29.8% of sales at a 32.34%
   margin — 4.21 pp below the portfolio rate, about **$108,744** of profit below benchmark. BevCo,
   only 14.9% of sales, runs at **41.97%** (+$70,047). Sodapop is 50.7% of sales at 37.42%.
2. **The West is the largest and weakest region** — $2,835,100 (32.6% of sales) at 32.56% — but it is
   91% FizzySip by sales, so this is a retailer-terms issue appearing geographically.
3. **2021 grew on price as well as volume.** Unit-weighted price rose $0.379 (Jan) → $0.624 (Dec),
   **+64.6%**, while monthly units rose 1,347,750 → 1,677,000 (**+24.4%**). The monthly margin rate
   stayed inside a narrow **35.84%–36.98%** band all year.
4. **December was the peak, March the trough:** $1,045,775 vs $483,530 (**+116.3%** on the March
   base); July was second at $1,041,438. All four retailers rose from November to December, so the
   peak was broad, not concentrated.
5. **The Coca-Cola brand is the profit engine:** $2,015,890 of sales (23.2%), **39.35%** margin,
   25.0% of portfolio profit — 2.80 pp above the portfolio. Price and volume show no negative
   relationship (r = −0.03 across 24 retailer-brand points; +0.14 across 312 city-brand points),
   so **no elasticity claim is made**.

### Three decisions the evidence supports

1. **Open a margin review with FizzySip** on Dasani Water (27.3% on $488,700) and Diet Coke
   (29.8% on $494,912) before agreeing any volume target.
2. **Protect the 2021 price path** and monitor units monthly instead of discounting for volume.
3. **Treat the West as a retailer-mix issue**, not a regional campaign.

No revenue, profit or ROI uplift is forecast anywhere in the workbook, report or script. The gaps
quoted are arithmetic, not projections.

---

## What is in the repository

| Path | What it is |
|---|---|
| `deliverables/Coca_Cola_Sales_Analysis.xlsx` | **Main workbook:** 17 sheets, **18 native PivotTables on one shared cache**, slicers for Retailer, Region and Beverage Brand, six charts on the dashboard, a PivotTable heatmap, formula-driven analysis sheets and a validation sheet. |
| `deliverables/Coca_Cola_Project_Report.docx` / `.pdf` | Project report: scope, data quality, findings, recommendations, limitations. |
| `deliverables/Coca_Cola_Excel_Implementation_Guide.docx` / `.pdf` | Click-by-click rebuild guide for every view, plus slicer setup and Power Query M code. |
| `deliverables/Nishant_Loom_Presentation_Script.docx` / `.pdf` / `.txt` | Narrated presentation script with evaluator Q&A and a recording checklist. |
| `deliverables/Reproduce_Coca_Cola_Analysis.py` | Standalone reproduction of every figure (needs only `openpyxl`). |
| `deliverables/README.md` | Submission-pack cover sheet. |
| `work/build/` | The build pipeline that produced the workbook and documents (openpyxl, then Excel automation for PivotTables, slicers and pivot charts). |
| `work/facts.json` | Single source of truth for every number, so workbook cells, report prose and the script cannot drift apart. |

### Workbook sheets

| Sheet | Contents |
|---|---|
| `00_ReadMe` | Purpose, scope, sheet guide, navigation |
| `01_Raw_Data` | Verbatim copy of the source (header row 5, records 6–3893) |
| `02_Clean_Data` | Analysis table: 3,888 rows with date, margin and validation helper columns |
| `03_Data_Dictionary` | Field definitions and metric formulas in Excel and DAX |
| `04_Data_Quality` | Issue log and treatment decisions |
| `05_KPI_Summary` | Portfolio KPIs with live formulas against the table |
| `06_Retailer_Analysis` | Retailer league table, share, margin gap, profit vs benchmark |
| `07_Region_Analysis` | Region, state and city-market detail with a materiality rule |
| `08_Brand_Analysis` | Brand table and the Coca-Cola-brand scope split |
| `09_Monthly_Analysis` | Monthly sales, units, profit, margin, MoM growth and the price path |
| `10_Price_Volume` | Price–volume relationships and the materiality input control |
| `11_Pivot_Tables` | PivotTable layouts and build specifications |
| `12_Dashboard` | KPI cards, slicers, pivot charts, heatmap and insight panel |
| `13_Insights_Actions` | Prioritised recommendations with evidence and success KPIs |
| `14_Implementation` | Step-by-step Excel rebuild guide |
| `15_Validation` | Every reconciliation check, including disclosed assumptions |
| `16_Loom_Script` | Presentation script and recording checklist |

---

## How to use it

1. Open `deliverables/Coca_Cola_Sales_Analysis.xlsx` in desktop Excel (2016 or later). If values look
   stale, press `Ctrl+Alt+F9` once to force a full recalculation.
2. Read `00_ReadMe` → `04_Data_Quality` → `12_Dashboard`.
3. On the dashboard, the slicers (Retailer, Region, Beverage Brand) filter the KPI cards, the pivot
   charts and the heatmap together, because every PivotTable shares one cache.
4. Sheets `05`–`10` are formula baselines that deliberately **ignore** the slicers — they are the
   auditable full-year numbers.
5. `15_Validation` records every reconciliation and flags the disclosed assumptions.

To re-verify the numbers without Excel:

```bash
pip install openpyxl
python deliverables/Reproduce_Coca_Cola_Analysis.py
```

---

## Method notes

**Weighted margin, never an average of percentages.** `Operating Margin = SUM(Operating Profit) /
SUM(Total Sales)` for the stated scope. Averaging the 3,888 row-level margins gives 36.31% instead of
36.55%, because it weights a 2,000-unit invoice the same as a 12,750-unit one.

**Realised sales per unit vs listed price.** Total Sales equals Price per Unit × Units Sold in every
row, so the two coincide in this dataset — reported explicitly rather than presented as a discount
finding.

**Materiality rule.** A segment is discussed only when its 2021 sales reach **1% of portfolio sales
($86,840)**, so small markets cannot drive a recommendation.

**Shares** are of the full portfolio unless a sheet states otherwise. **Margin differences** are in
percentage points. **January** month-over-month growth is blank because December 2020 is not in the
data. **No time intelligence** (DATEADD, SAMEPERIODLASTYEAR) is presented, because with a single year
those measures would return blank or identical values.

---

## Data quality decisions

| Issue | Count | Treatment |
|---|---|---|
| Zero-volume lines (Omaha / Fanta, 10 Mar & 11 Apr 2021) | 2 rows | Retained and flagged. A stated 40% margin on zero units makes profit ÷ sales undefined, so the recalculated margin is blank — never 0%. |
| Rows flagged `x` in column A, with an unheaded price/units/margin triplet in columns P–R on two of them | 47 rows | Records kept; the stub columns are excluded from every metric. |
| Sub-cent profit values (e.g. 528.125) | 3,888 rows | Stored unrounded — rounding to cents would shift portfolio profit by **$0.19**. |
| Same-named cities across states (Portland ME/OR, Charleston SC/WV) | 2 names / 4 markets | City analysis groups by **State + City**: 54 markets, 52 names. Grouping on city name alone inflates the top-five concentration from 18.3% to 19.2%. |
| Invoice dates vary by city within a month (Sodapop 16–20 dates, FizzySip 10, BevCo 5, DreamCo 2) | 270 dates | Only the **month** is used for trend analysis. |

No rows were deleted, no values imputed, and no negative or zero values silently replaced.

**Validation:** `Total Sales = Price × Units` and `Operating Margin = Profit ÷ Sales` hold in all
3,888 rows with zero exceptions; retailer, region, brand and monthly blocks each reconcile to the
portfolio totals; there are zero duplicate transaction keys and zero missing values in the 12
required fields.

---

## Known limitations

- **Currency is not documented in the source.** USD is an assumption, flagged on `04_Data_Quality`.
- **The Invoice Date timeline is not automated.** Add it with `PivotTable Analyze → Insert Timeline →
  Invoice Date` (three clicks). The slicers cover the filtering demonstrated in the script.
- **Retailer and region are confounded.** BevCo sells only in the South, FizzySip only in the West, so
  "BevCo has the best margin" is not a like-for-like comparison.
- **One year only.** Within-year pattern is described; seasonality is not proven and no prior-year
  comparison is made.
- **No cost, promotion, competitive or shelf-space data**, so margin gaps are not attributed to cost,
  trade spend or pricing behaviour, and higher recorded sales are not treated as greater efficiency.
- **City rows are invoiced volume**, not consumer offtake, and coverage differs by retailer.
- The build scripts under `work/build/` need Windows with Excel installed; they are development
  tooling, not part of the submission.

---

## Submission checklist

- [ ] `deliverables/Coca_Cola_Sales_Analysis.xlsx` opens and recalculates (`Ctrl+Alt+F9` if needed)
- [ ] Slicers filter the dashboard and `15_Validation` shows no failures
- [ ] Report, implementation guide and Loom script read consistently with the workbook figures
- [ ] Loom recording follows the script timings and shows the dashboard at default filters
- [ ] No unsupported claim: portfolio vs Coca-Cola-brand scope is stated wherever a number appears

## Licence

Code: MIT. Workbook, report, guide and script: © 2026 Nishant, shared for educational and portfolio
use. The dataset originates from the assignment and is included only for reproducibility.
