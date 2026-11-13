# Remaining steps — what is not yet built, and exactly how to finish it

Companion to `Coke_Sales_Blue_Dashboard_v2.xlsx`. Everything in this document is a **desktop click-path**, because
these objects cannot be authored by a script without Excel's permission and cannot be verified from XML alone.
The workbook already contains the code and layouts: `Power_Query_Code` (M), `DAX_Measures` (18 definitions),
`Pivot_Tables` (9 layouts), `Implementation_Guide` (22 steps, rows 5–26).

Work on a **copy** (`Coke_Sales_Blue_Dashboard_v3.xlsx`). Keep v2 untouched as the reconciliation guardrail —
the delivered dashboard's SUMIFS figures are the numbers every pivot must reproduce.

---

## 0. Status snapshot

| Component | State | Where the spec lives |
| --- | --- | --- |
| Cleaned table `SalesData` (3,888 rows, 23 cols), `Calendar` 2020–2022 | Built (static values from Python) | `Cleaned_Data`, `Calendar` |
| Location key `State \| City` | Built, repaired; 54 distinct, 0 `#REF!` | `Cleaned_Data` col H |
| Dashboard: 5 cards, 4 dropdowns, 6 charts, heatmap | Built, formula-driven, recalculated in Excel | `Dashboard` |
| Static analysis sheets (7) + insights + forecast | Built | `Retailer_…` to `Insights_Recommendations` |
| Power Query connection | **Not built** | `Power_Query_Code` |
| Power Pivot model + relationship + date table | **Not built** | `Implementation_Guide` rows 13–14 |
| 18 DAX measures | **Not built** | `DAX_Measures` |
| 9 native PivotTables on one shared source | **Not built** | `Pivot_Tables` |
| Slicers (Retailer, Region, Brand) + Calendar timeline | **Not built** | `Pivot_Tables` row 16 |
| Prior-year actuals / YoY | **Cannot be built** — no 2020 data supplied | `Forecast` col F, `KPI_Summary` |

Estimated effort for the five build items: 60–90 minutes, most of it in the pivots and slicer connections.

---

## 1. Enable the prerequisite

Power Pivot is a COM add-in and is **off by default** even in Excel 2016/2019/365:

`File → Options → Add-ins → Manage: COM Add-ins → Go… → tick "Microsoft Power Pivot for Excel" → OK`

A **Power Pivot** tab appearing on the ribbon is the success signal. Power Query is already built in
(`Data → Get & Transform Data`); no install needed.

---

## 2. Load the source through Power Query (not a copy-paste)

1. **Parameter for the path.** `Data → Get Data → From Other Sources → Blank Query → Manage Parameters → New
   Parameter`: name `pSourceFile`, type **Text**, current value = full path of the *original* workbook
   (`…\Coke_Sales_Blue_Dashboard_and_Implementation.xlsx`). Copy the `pSourceFile` block from
   `Power_Query_Code` into its Advanced Editor if you prefer paste over typing.
2. **Raw connection.** `Data → Get Data → From File → From Excel Workbook` → pick `pSourceFile` → select the
   **Data** sheet → **Transform Data**. In Power Query: remove the top 4 rows, *Use First Row as Headers*.
3. **Trim to the 12 business columns.** Keep Retailer, Retailer ID, Invoice Date, Beverage Brand, Units Sold,
   Price per Unit, Total Sales, Operating Profit, Operating Margin, Region, State, City. **Delete** column A
   (the source markers) and P:R (the Honolulu/Manchester adjustment formulas) — they are not actuals.
4. **Types.** Retailer ID → **Text** (never number: 1128299, 1185732, 1189833, 1197831 are identifiers);
   Invoice Date → Date; Units Sold → Whole Number; price/sales/profit/margin → **Decimal Number**.
   Do not use Fixed Decimal Currency — it rounds the 0.506408963-per-unit precision.
5. **Clean and derive.** Trim + Clean the five dimension columns; add Year, Month Number, Month Start,
   Year-Month, Quarter, and `Location = [State] & " | " & [City]` (state first — city names repeat), then
   Sales Difference, Profit Difference, Sales Check, Profit Check (0.01 tolerance) and Activity Flag, exactly as
   in `Power_Query_Code` rows 7–40.
6. **Calendar query.** New Blank Query → Advanced Editor → paste the Calendar block (2020-01-01 to 2022-12-31,
   with Year, Month Number, Month Name, Month Start, Year-Month, Quarter).
7. **Load.** Close & Load To…: `SalesData` → **Table** on `Cleaned_Data` **and** tick *Add this data to the Data
   Model*; `Calendar` → **Only Create Connection** + *Add to Data Model*.

**Check before continuing:** the table has 3,888 data rows and 23 columns; the Totals row of `SalesData` reads
8,684,027.50 / 17,148,250 / 3,173,631.875. If it does not, stop — the query scope is wrong, and every pivot
downstream will be wrong too.

---

## 3. Build the model

1. `Power Pivot → Manage → Diagram View`.
2. Drag `Calendar[Date]` onto `SalesData[Invoice Date]`. Confirm the cardinality reads **one to many**, solid
   line, single arrow (Calendar → SalesData).
3. `Design → Table Properties` (or `Calendar` right-click) → **Mark as Date Table** → Date column `Date`.
4. In the `Calendar` table, sort **Month Name** by **Month Number** (otherwise month labels sort alphabetically).
5. Hide `SalesData[Invoice Date]` from client tools so the slicer list shows only the Calendar dates.

---

## 4. Create the 18 measures

`Power Pivot → Measures → New Measure`, one at a time, using the name / expression / format columns of
`DAX_Measures` rows 5–22. Group them as follows so mistakes are easy to isolate:

| Group | Measures | Rule that must hold |
| --- | --- | --- |
| Additive | Sales, Units, Operating Profit | Plain `SUM`, never wrapped in `CALCULATE` filters |
| Ratios | Weighted Margin, Revenue per Unit, Profit per Unit | `DIVIDE` of two sums — a percentage column is never averaged |
| Shares | Sales Share, Profit Share | numerator ÷ `CALCULATE(…, ALLSELECTED(...))` |
| Time | Prior Month Sales, MoM Sales Growth, Prior Month Margin, Margin Change pp, Prior Year Sales, YoY Sales Growth | `DATEADD` on `Calendar[Date]`; January and any missing comparator return **BLANK**, not 0 |
| Guards | Sales at Zero Denominator, Active Locations | `ISBLANK` / zero-denominator returns `BLANK()` |

**Non-negotiable checks on this sheet:** Margin = 36.5456%; Revenue/Unit = 0.506408963; **YoY Sales Growth is
blank for all 12 months** (`DATEADD` finds no 2020 rows). If YoY shows 0% or 100%, the date table or the
relationship is wrong — fix that before building pivots.

---

## 5. Build the nine pivots on one source

1. `Insert → PivotTable → Use this workbook's Data Model` → New Worksheet → name it `PT_KPI` → OK. This creates
   the one shared cache that all nine will reuse.
2. For each row of `Pivot_Tables` (5–13): select the existing pivot → `PivotTable Analyze → PivotTable →
   Existing Worksheet` → choose the target cell. **Never** use *New Worksheet* again, and never *From Table/Range*:
   a second cache breaks Report Connections and the slicers will only filter some pivots.
3. Field placement per row: Rows / Columns / Values / Filter as listed. Set values to the DAX measures (not
   automatic Sum of column), number formats from `DAX_Measures`.
4. Sort Sales **descending** on every bar sheet; `PT_Month` and `PT_Heatmap` sort by date **ascending**;
   `PT_Location_Margin` filters Sales ≥ 100,000 and sorts margin descending.
5. Hide grand totals on `PT_Month` (the change columns must not sum) and on `PT_Heatmap` (totals double the
   12-month axis).

**Check:** `PT_KPI` must read 8,684,027.50 / 17,148,250 / 3,173,631.875 / 36.5456% / 0.506409 — identical to the
Dashboard cards. Any difference is a filter or cache problem, not a rounding problem.

---

## 6. Slicers and timeline

1. Click `PT_KPI` → `PivotTable Analyze → Insert Slicer` → tick **Retailer**, **Region**, **Beverage Brand**.
2. `Insert Timeline` → **Calendar[Date]** → then `Timeline → Filter by: Months`, and drag to cover Jan–Dec 2021.
3. On each slicer/timeline: `Slicer → Report Connections` (or right-click → Report Connections) → **tick all
   nine pivots**. This is the step most often half-done; a slicer that filters five of nine pivots looks fine
   until someone checks a subtotal.
4. Arrange the four controls in one row, above the pivots, and set `Slicer Settings → Hide items with no data`.
5. There is no macro reset button: to return to baseline, click the **Clear Filter** funnel on each control.

**Check:** select FizzySip → every pivot drops to 2,584,450 sales / 4,109,000 units / 32.34% margin. Then
BevCo **and** West together → 0 / 0 / 0 with blank ratios (not 0.00%).

---

## 7. Re-verify after the model is live

Run these before you present or publish; all figures are already recorded as passing in `Validation`:

| Test | Expected |
| --- | --- |
| `PT_KPI` vs Dashboard cards, All filters | identical to the cent |
| FizzySip | 2,584,450 / 4,109,000 / 32.3379% |
| Coca-Cola | 2,015,890 / 793,197.25 / 39.3472% |
| 2021-12 | 1,045,775 / 1,677,000 |
| BevCo + West | 0 / 0 / 0, ratios blank |
| Location count | 54, no blanks, no `#REF!` |
| Sales and Profit Check columns | 0 exceptions at 0.01 |
| Duplicate monthly keys | 0 (do not dedupe on retailer + date alone) |
| YoY | blank / "Not supplied" — never 0% |

Then `Data → Refresh All` once more and re-read `PT_KPI`: a refresh that changes a total means the query scope
drifted (usually extra rows or the A/P:R helper columns coming back in).

---

## 8. When real 2020 data arrives (the only true gap)

1. Add the 2020 file/rows to `pSourceFile` (same 12 columns) or point the query at a folder.
2. Confirm `Calendar` already covers 2020 — it does (2020-01-01 to 2022-12-31), so no date-table change.
3. `Refresh All`. YoY and Prior Year Sales unblank for 2021 months with a 2020 comparator; January 2021 stays
   blank only if 2020-01 is genuinely absent.
4. Recolour: prior-year series becomes `#9DC3E6` (light blue) now that it is real, per `Color_Theme`.
5. Re-run the validation table and record the new baseline. Do **not** overwrite the 2021 figures in `README`
   until the numbers reconcile.

Until then every prior-year cell stays blank. That is a missing-input state, not an error to be filled.

---

## 9. Presentation hardening (optional, 10 minutes)

- `View → Freeze Panes` on `Cleaned_Data` is already set at `C2`; do the same on any new pivot sheet.
- `Page Layout → Print Titles` for the wide analysis sheets, landscape, fit to 1 page wide.
- `File → Info → Inspect Workbook` before sending outside the team: no external links, no macros, no hidden sheets.
- Keep the four dropdowns at **All** when saving, so the file opens on the baseline rather than someone's last filter.

---

## Definition of done

- [ ] Power Query queries exist and `Refresh All` is reproducible
- [ ] Relationship is one-to-many, Calendar marked as date table, Month Name sorted by Month Number
- [ ] 18 measures created; margin 36.5456%, YoY blank
- [ ] Nine pivots on one shared cache, no second cache anywhere
- [ ] Three slicers + one timeline, Report Connections ticked for all nine
- [ ] Section 7 table re-run and passing
- [ ] v2 retained unchanged as the reconciliation reference
