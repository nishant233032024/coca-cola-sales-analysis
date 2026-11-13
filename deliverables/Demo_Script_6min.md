# 6-minute video demo script

Workbook: `Coke_Sales_Blue_Dashboard_v2.xlsx` (23 sheets). Target length **5:50–6:10**. Narration below measures
**777 words**: 5:33 at 140 words per minute, or 5:21 at 145. The remaining ~35 seconds covers nine tab switches and
six dropdown changes, so that is what lands it at six minutes — do not add ad-libs. Read the **Say** column aloud; the **Do**
column is your click track. Timings are cumulative — if a segment overruns, cut that paragraph rather than
speeding up.

---

## Before you hit record

| Item | Setting |
| --- | --- |
| File | Open `Coke_Sales_Blue_Dashboard_v2.xlsx`, `Enable Editing` if prompted, land on the **Dashboard** |
| Filters | All four dropdowns at **All** (B4, F4, J4, N4) so you start on baseline |
| Window | Excel maximised at 1920×1080 or 1600×900, ribbon collapsed (`Ctrl+F1`), status bar visible |
| Zoom | 100% on the Dashboard; the sheets are laid out for a full-width read |
| Sheet tabs | Right-click a tab → *Activate* only when needed; keep the tab strip visible so viewers see sheets change |
| Cursor | Slow, deliberate; hover a cell before you click it, and pause ~1s after each change so the numbers settle |
| Audio | Say numbers as they appear — do not pre-empt Excel's recalculation |

---

## The script

### 0:00 — 0:30 · What this is

| Do | Say |
| --- | --- |
| Dashboard visible, cursor idle | "This is a rebuild of the Coke sales workbook. Twenty-three sheets, one year of data — 3,888 records, six brands, four retailers, five regions and fifty-four locations. The original file had a broken Location column, so I rebuilt it and validated every number you're about to see. Total sales: 8.68 million. Margin: 36.5 percent." |

### 0:30 — 1:30 · The dashboard and the four filters

| Do | Say |
| --- | --- |
| Hover the five cards, row 7–8 | "Five cards across the top — sales, units, operating profit, weighted margin, revenue per unit. These are formula-driven, not pasted values, so they respond to the four dropdowns above." |
| Click **B4**, choose **FizzySip** | "Retailer: FizzySip. Sales drop to 2.58 million on 4.1 million units, and margin falls to 32.3 percent — about four points below the portfolio. That gap is the story of this file." |
| **B4** back to All, set **J4** to **Coca-Cola** | "Brand filter instead: Coca-Cola, 2.02 million sales at 39.3 percent margin. Higher margin, smaller base." |
| **J4** to All, **N4** to **2021-12** | "Month: December, 1.05 million — the annual peak — on 1.68 million units." |
| **N4** to All, **B4** to **BevCo**, **F4** to **West** | "Now the important one. BevCo and West together return zero — those two never trade with each other. And note the margin cards are **blank**, not zero. A ratio with a zero denominator is undefined, and this workbook refuses to print 0.00 percent." |
| Reset all four to **All** | "Back to the baseline." |

### 1:30 — 2:15 · Scale versus efficiency

| Do | Say |
| --- | --- |
| Scroll to **SCALE VS EFFICIENCY** (row 11), two charts | "Two charts, deliberately separate. Left: retailer sales, sorted. Sodapop is half the business — 50.7 percent of sales and 51.9 percent of profit. Right: the same retailers by weighted margin, on its own percent axis. Sodapop is also the most profitable at 37.4 percent; BevCo is the strongest at 42 percent, but it only trades in the South, on eight locations. So scale and efficiency are different questions, and they get different charts." |

### 2:15 — 2:55 · Where it sells

| Do | Say |
| --- | --- |
| Scroll to **WHERE IT SELLS** (row 29) | "Region and brand. West is 2.84 million, 32.6 percent of sales, but it runs at 32.6 percent margin — four points below the portfolio. One point on West sales is about 28,000 of operating profit, so that's the gap worth investigating. Brands are concentrated: Coca-Cola and Dasani carry the range, the tail contributes little." |

### 2:55 — 3:40 · Time, then region by month

| Do | Say |
| --- | --- |
| Charts row 48: monthly sales line, then margin line | "Two lines, both chronological, months in calendar order rather than alphabetical. Sales peak in December at 1.05 million, and again in July on volume — 1.76 million units. Margin moves far less, 35.8 to 37 percent, so the year is a volume story, not a pricing story." |
| Scroll to heatmap, rows 65–72 | "Region by month, twelve fixed columns, plain-blue data bars for magnitude. Timing splits: West peaks in July, the other four regions peak in December. A single national supply plan would under-serve July and over-serve the fourth quarter. The real columns sum to 8.68 million, same as the cards — the grid isn't decorative." |

### 3:40 — 4:25 · Locations, and the defect that was fixed

| Do | Say |
| --- | --- |
| Tab **Location_Analysis** | "Fifty-four locations, ranked by sales. New York leads at 392,000; San Francisco is next at 332,000 — and San Francisco is the margin problem, 29.1 percent." |
| Scroll to the qualified margin table | "The second table only ranks locations above 100,000 of sales, because a percentage on a tiny market misleads. Thirty-seven of the fifty-four qualify. Birmingham is best at 51.5 percent, Anchorage weakest at 25.6 percent." |
| Tab **Cleaned_Data**, column H | "This sheet is the fix. In the file I was given, this Location column was cached as a broken reference — literally hash-REF — so every location pivot returned nothing. It's rebuilt as State, pipe, City. State first, because city names repeat across states. Fifty-four distinct values, zero errors — which is why these location sheets have numbers at all." |

### 4:25 — 5:00 · Price, and the forecast

| Do | Say |
| --- | --- |
| Tab **Pricing_Analysis** | "Price against units: pooled correlation 0.247 — weak. Careful with the wording: **association, not elasticity**. Sales are price times units, so they're mechanically linked, and brand, retailer and month confound it. The honest conclusion is 'test it', not 'raise price and expect this'." |
| Tab **Forecast** | "Three months of illustrative forecast — 821,000 a month, the October-to-December average held flat. Labelled illustrative, never added to actual totals. Holdout error sits next to it: 172,000 a month, so treat it as a placeholder, not a plan." |

### 5:00 — 5:35 · The ask

| Do | Say |
| --- | --- |
| Tab **Insights_Recommendations** | "Six findings: why it matters, caveat, action, proposed owner. Protect Sodapop distribution first — half the profit. Second, diagnose cost-to-serve in Seattle and San Francisco before touching price. Third, plan capacity by region, not nationally." |
| Point at the caveat column | "Every row carries its caveat. There's no uplift figure here, because cost, promotion and trade-spend data aren't in the file — and I won't invent them." |
| Tab **Forecast**, column F | "One deliberate gap: prior year says 'not supplied'. There's no 2020 data, so year-on-year stays blank for all twelve months. A flat year and a growing year look identical without a comparator." |

### 5:35 — 6:00 · Proof and close

| Do | Say |
| --- | --- |
| Tab **Validation** | "Finally, proof. Twenty-three acceptance tests — baselines, the three filtered cases you just watched, location count, the sales and profit identities — all passing. Numbers were recalculated in Excel and re-read, so these are live values, not screenshots." |
| Tab **README** | "The README carries the seven sections: data quality, measure definitions, chart list, build steps, insights, limitations, validation. The Power Query, DAX and pivot specs have their own sheets. Thanks for watching." |

---

## Numbers you must land exactly

| Say | Never say |
| --- | --- |
| 8,684,027.50 · "8.68 million" | 8.6 million rounded loosely, or "about 9 million" |
| 36.5456% · "36.5 percent" | "36 percent average margin" (it is a ratio of sums) |
| 17,148,250 units | "17 million units" |
| Sodapop 50.71% of sales, 51.93% of profit, 37.42% margin | "Sodapop is the most efficient retailer" (BevCo is, on 41.97%) |
| West 32.56% margin, −3.99 pp vs portfolio | "West is unprofitable" |
| r = 0.247, **association not elasticity** | "elasticity", "price sensitivity", "a 1% price rise costs X" |
| Forecast 821,587.50/month, illustrative, MAE 172,008.33 | presenting the forecast as a target or a commitment |
| Prior year: "not supplied" | 0%, or a fabricated 2020 comparison |
| 0 frozen rows, 2 zero-activity rows retained | "we cleaned out the bad rows" |

## Delivery notes

- **Never** call the numbers "Coca-Cola company results". The dataset is educational; say "this dataset".
- The strongest 20 seconds of the video are the **BevCo + West blank-ratio** moment — keep it, it proves the
  measure rules rather than describing them.
- If asked why there are no native pivot tables on screen: they're specified on `Pivot_Tables` and
  `Implementation_Guide`, and `Remaining_Steps.md` has the click-path; the delivered dashboard is formula-driven
  and reaches identical figures.
- Total screen changes: 9 tab switches, 6 dropdown changes. Rehearse the dropdown order once — it is the only
  place timing can slip.
