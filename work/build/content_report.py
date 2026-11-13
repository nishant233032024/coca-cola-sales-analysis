"""Recommendation table and project report text, driven by verified facts."""

from content_loom import money, m_, pct


def build_insights(f):
    t = f["total"]
    ret = {r["name"]: r for r in f["levels"]["retailer"]}
    reg = {r["name"]: r for r in f["levels"]["region"]}
    brd = {r["name"]: r for r in f["levels"]["brand"]}
    dc = f["decomp"]
    mat = f["materiality"]
    west = reg["West"]
    soda, fizzy, bev = ret["Sodapop"], ret["FizzySip"], ret["BevCo"]
    bw = f["brand_region"]
    hs = f["high_sales_low_margin"]
    names = {x["name"]: x for x in hs}
    gaps = (west["margin"] - t["margin"]) * 100
    return [
        dict(
            pri=1,
            finding="The West carries the largest sales base but the lowest operating margin in the portfolio.",
            evidence=(f"West sales {money(west['sales'])} ({pct(west['sales_share'])} of total) and operating profit "
                      f"{money(west['profit'])} ({pct(west['profit_share'])}), margin {pct(west['margin'])} vs "
                      f"portfolio {pct(t['margin'])} ({gaps:+.1f} pp). Within-brand evidence: Coca-Cola "
                      f"{pct(bw['Coca-Cola']['West']['margin'])} in the West vs "
                      f"{pct(bw['Coca-Cola']['Southeast']['margin'])} in the Southeast; Dasani Water "
                      f"{pct(bw['Dasani Water']['West']['margin'])} vs {pct(bw['Dasani Water']['South']['margin'])}."),
            action=("Run a matched-cell margin review (brand x city x month) on the lowest-margin material cities, "
                    "starting with Seattle, Anchorage, San Francisco and Providence, to separate price effect from "
                    "mix effect before any commercial change."),
            owner="Proposed role: regional commercial analyst (West) with pricing support",
            kpi="Margin gap in percentage points on matched cells; profit per unit by city-brand",
            risk=("Cost, promotion and store-count data are absent, so lower margin cannot be attributed to cost. "
                  "FizzySip operates only in the West, so retailer and region effects are confounded."),
            more="Cost of goods per unit, trade-spend records, and store counts per city.",
        ),
        dict(
            pri=2,
            finding="Profit is concentrated in Sodapop, and its margin sits between the best and worst accounts.",
            evidence=(f"Sodapop: {pct(soda['sales_share'])} of sales, {pct(soda['profit_share'])} of profit, margin "
                      f"{pct(soda['margin'])}. BevCo: {pct(bev['sales_share'])} of sales but {pct(bev['margin'])} "
                      f"margin. FizzySip: {pct(fizzy['sales_share'])} of sales, {pct(fizzy['profit_share'])} of profit, "
                      f"{pct(fizzy['margin'])} margin - a {abs((fizzy['profit_share']-fizzy['sales_share'])*100):.1f} pp "
                      "negative profit-share gap."),
            action=("Protect Sodapop volume with availability and service-level tracking, and use the BevCo and "
                    "Sodapop margin gap as a benchmarking question for account teams - not as a proven efficiency gap."),
            owner="Proposed role: national account leads; category finance for benchmarking",
            kpi="Sodapop units and margin retained year on year; margin gap to best-performing account in matched cells",
            risk=("Accounts differ in geography and product mix, and no store count or resource measure is supplied, "
                  "so a higher recorded margin is not evidence of greater efficiency."),
            more="Store counts, trade investment per account, and service-level data.",
        ),
        dict(
            pri=3,
            finding="Within 2021, revenue growth was price-led more than volume-led, while margin stayed almost flat.",
            evidence=(f"{dc['jan']} to {dc['dec']}: sales {pct(dc['sales_growth'])}, units {pct(dc['units_growth'])}, "
                      f"realised sales per unit {dc['jan_rpu']:.2f} to {dc['dec_rpu']:.2f} ({pct(dc['price_growth'])}). "
                      f"Monthly margin range {pct(min(m['margin'] for m in f['months']))} to "
                      f"{pct(max(m['margin'] for m in f['months']))} ({f['month_extremes']['margin_range_pp']:.1f} pp)."),
            action=("Add a volume and margin guardrail to every price change: test in matched stores, monitor units per "
                    "store per month, and treat realised price per unit as a tracked KPI next to sales."),
            owner="Proposed role: revenue management / pricing analyst with sales operations",
            kpi="Units per store per month, realised sales per unit, weighted margin before and after each price change",
            risk=("One year of data cannot separate price action from seasonality or market growth. Pooled price-volume "
                  "correlation is +"
                  f"{f['correlation']['pooled']:.2f}, an association, not elasticity."),
            more="Two to three years of history, promotion calendars, and competitor price data.",
        ),
        dict(
            pri=4,
            finding="Demand peaks twice for different reasons: December on realised price, July on units.",
            evidence=(f"December sales {money(f['month_extremes']['high_sales'])} vs March "
                      f"{money(f['month_extremes']['low_sales'])} ({pct(f['month_extremes']['rel'])} higher). July units "
                      f"{f['months'][6]['units']:,} vs January {f['months'][0]['units']:,}. December units "
                      f"{f['months'][11]['units']:,} with realised price {dc['dec_rpu']:.2f}."),
            action=("Plan two separate peaks: build stock and service capacity for the July volume peak, and protect "
                    "availability and price execution for the December revenue peak."),
            owner="Proposed role: demand planning and supply chain",
            kpi="Forecast error by month, in-stock rate in June-August and November-December",
            risk=("Only one year is observed, so this is a within-year pattern, not proven recurring seasonality. No "
                  "stock, availability or promotion field exists to confirm the mechanism."),
            more="Two or more years of monthly history plus in-stock and promotion data.",
        ),
        dict(
            pri=5,
            finding="Sales are concentrated: a small share of cities produces most of the revenue.",
            evidence=(f"Top 10 of {mat['total']} cities = {pct(f['concentration']['top10'])} of sales; "
                      f"{mat['count']} cities at or above $100,000 of annual sales = {pct(mat['sales_share'])} of sales. "
                      f"New York, New York alone: {money(names.get('New York | New York', {}).get('sales', 0))} at "
                      f"{pct(names.get('New York | New York', {}).get('margin', 0))} margin."),
            action=("Concentrate account coverage and analytics on the material cities, and hold the sub-$100,000 cities "
                    "on a maintenance review that first verifies distribution coverage."),
            owner="Proposed role: sales operations / territory planning",
            kpi="Coverage cost per material city; revenue and margin per covered city",
            risk=("Low sales in small cities may reflect limited distribution rather than weak demand; the dataset has "
                  "no coverage measure. The $100,000 threshold is an analyst convention, not a company policy."),
            more="Distribution coverage, store counts, and cost-to-serve by city.",
        ),
    ]


def build_report(f):
    t = f["total"]
    ret = {r["name"]: r for r in f["levels"]["retailer"]}
    reg = {r["name"]: r for r in f["levels"]["region"]}
    brd = {r["name"]: r for r in f["levels"]["brand"]}
    dc = f["decomp"]
    mat = f["materiality"]
    me = f["month_extremes"]
    west, south = reg["West"], reg["South"]
    soda, fizzy, bev, dream = ret["Sodapop"], ret["FizzySip"], ret["BevCo"], ret["DreamCo"]
    coke = brd["Coca-Cola"]
    hs = {x["name"]: x for x in f["high_sales_low_margin"]}
    q = f["quality"]

    return [
        ("Executive summary", [
            f"This project analyses {f['records']:,} recorded sales observations covering the period "
            f"{f['date_min']} to {f['date_max']} ({len(f['years'])} year). In the supplied data, total sales are "
            f"{money(t['sales'])}, units sold are {t['units']:,}, and operating profit is {money(t['profit'])}, a "
            f"weighted operating margin of {pct(t['margin'])}. That is {money(t['sales_per_unit'],2)} of sales per unit "
            f"and {money(t['profit_per_unit'],2)} of operating profit per unit.",
            f"Three findings matter most for decisions. First, profitability is uneven by geography and account: the West "
            f"produces {pct(west['sales_share'])} of sales but only {pct(west['profit_share'])} of profit at "
            f"{pct(west['margin'])}, {abs((west['margin']-t['margin'])*100):.1f} percentage points below the portfolio, "
            f"while the South earns {pct(south['margin'])}. Second, growth inside 2021 was price-led: from January to "
            f"December sales rose {pct(dc['sales_growth'])} while units rose {pct(dc['units_growth'])} and realised sales "
            f"per unit rose {pct(dc['price_growth'])}. Third, demand has two distinct peaks - December on price and July "
            f"on volume - so a single 'peak season' plan would misallocate effort.",
            "The recommendations are therefore investigative rather than prescriptive. The data can show where margin is "
            "weak, but it contains no cost, promotion, distribution or store-count field, so it cannot show why. Each "
            "recommendation names the additional evidence required before a commercial change is made.",
        ]),
        ("Business question and objectives", [
            "The question is: how can Coca-Cola optimise sales and operating profitability across U.S. retailers, "
            "regions, products and time?",
            "The objectives were to (1) establish a clean, auditable record-level dataset; (2) define metrics once and "
            "use them consistently; (3) compare performance across retailer, region, state/city, brand and month; "
            "(4) test whether apparent differences survive controlling for mix; and (5) turn the verified findings into "
            "proportionate actions with named evidence gaps.",
        ]),
        ("Dataset structure and scope", [
            f"The source file is {f['filename']}. It contains one raw data sheet, Data, whose headers sit on row "
            f"{f['header_row']} and whose records run from row {f['data_rows'][0]} to row {f['data_rows'][1]}. Above "
            "the header there are two decorative title rows and one blank row; those are not data. The sheet also holds "
            "47 'x' marker cells in column A and four stray unlabelled numeric values in columns P to R on the last "
            "four rows. Neither is part of the business data, and neither is used anywhere in the analysis.",
            f"Records: {f['records']:,}. Business fields: {len(f['headers'])}. Coverage: {f['counts']['retailers']} "
            f"retailers, {f['counts']['regions']} regions, {f['counts']['states']} states, {f['counts']['cities']} "
            f"cities, {f['counts']['brands']} brands, {f['counts']['months']} months. Each month contains exactly 324 "
            "records, which is 54 cities x 6 brands, so the record grain is one retailer-city-brand-month observation. "
            "There is no order identifier, so a record is not an order and record counts are not transaction counts.",
            "Scope: the workbook covers the full six-brand portfolio supplied. Coca-Cola is one of those six brands and "
            "accounts for "
            f"{pct(coke['sales_share'])} of portfolio sales; brand-level figures are reported separately and are never "
            "described as Coca-Cola-brand-only totals. The dataset is an educational assignment dataset; nothing here is "
            "a statement about Coca-Cola's audited results.",
            "Currency is not documented anywhere in the workbook. All monetary figures are presented as US dollars "
            "because the dataset is U.S. retail; this is a stated assumption, not a verified fact. 'Operating profit' is "
            "used throughout because that is the supplied field; the workbook contains no net profit.",
        ]),
        ("Cleaning and quality decisions", [
            f"Raw rows and cleaned rows both equal {f['records']:,}; no record was dropped, added or edited. The original "
            "sheet is preserved as 01_Raw_Data and the analysis table (tblSales, on 02_Clean_Data) keeps Source Sheet "
            "and Source Row for every record, so any figure can be traced back.",
            f"Identity tests: price x units equals total sales in all {f['records']:,} rows (0 exceptions), and operating "
            f"profit / total sales equals the supplied margin in every row with sales (0 exceptions). Duplicate tests "
            f"found 0 exact duplicates on the twelve business fields and 0 duplicates on the retailer-location-brand-month "
            f"key. Negative values: {q['negatives']}. Invalid or missing dates: 0. Months with incomplete coverage: 0.",
            "Two records (Sodapop, Nebraska, Omaha, Fanta, March and April) have zero units, zero sales and zero profit "
            "but a supplied 40% margin. A margin cannot exist on a zero base, so these are flagged 'Zero sales - check' "
            "in the Validation Status column, retained rather than deleted, and their impact on every total is zero. "
            "They are the reason the workbook distinguishes a supplied margin from a recalculated one.",
            "Floating-point residue (values such as 5637.500000000001) is a storage artefact, not a data error; it is "
            "handled by formatting and by a 0.01 tolerance in the identity tests.",
        ]),
        ("Analytical methods", [
            "All KPIs are ratios of sums. Weighted operating margin = total operating profit / total sales; realised "
            f"sales per unit = total sales / total units; profit per unit = total operating profit / total units. The "
            f"simple average of row-level margins is {pct(f['row_margin_simple_avg'])} and is shown once, deliberately, "
            "to demonstrate why it is not used.",
            "Shares are always stated against a named scope (all data, or a named segment). Margin changes are reported "
            "in percentage points, because a change from 35.8% to 36.0% is +0.2 pp, not +0.6% growth.",
            "The work is built on Excel Tables and formulas (SUM, SUMIFS, SUMPRODUCT, CORREL) so every figure is "
            "recalculable, plus native PivotTables and slicers for interactive views. Python was used to build and "
            "independently verify the numbers; the verification results are on 15_Validation.",
        ]),
        ("Findings", [
            f"1. Portfolio baseline. {money(t['sales'])} sales, {t['units']:,} units, {money(t['profit'])} operating "
            f"profit, {pct(t['margin'])} weighted margin, {money(t['sales_per_unit'],2)} sales per unit, "
            f"{money(t['profit_per_unit'],3)} operating profit per unit. Scope: all six brands, all 3,888 records, "
            f"{f['date_min']} to {f['date_max']}. Chart: dashboard KPI cards; source 05_KPI_Summary.",
            f"2. Retailer concentration with a margin spread. Sodapop {pct(soda['sales_share'])} of sales / "
            f"{pct(soda['profit_share'])} of profit at {pct(soda['margin'])}; FizzySip {pct(fizzy['sales_share'])} / "
            f"{pct(fizzy['profit_share'])} at {pct(fizzy['margin'])}; BevCo {pct(bev['sales_share'])} / "
            f"{pct(bev['profit_share'])} at {pct(bev['margin'])}; DreamCo {pct(dream['sales_share'])} of sales at "
            f"{pct(dream['margin'])}. The spread between the best and worst account is "
            f"{(bev['margin']-fizzy['margin'])*100:.1f} percentage points. Chart: retailer bar on the dashboard; "
            "source 06_Retailer_Analysis.",
            f"3. Regional margin gap. West {pct(west['margin'])} on {pct(west['sales_share'])} of sales; Northeast "
            f"{pct(reg['Northeast']['margin'])}; Southeast {pct(reg['Southeast']['margin'])}; South {pct(south['margin'])}; "
            f"Midwest {pct(reg['Midwest']['margin'])}. The gap holds within brand for Coca-Cola "
            f"({pct(f['brand_region']['Coca-Cola']['West']['margin'])} West vs "
            f"{pct(f['brand_region']['Coca-Cola']['Southeast']['margin'])} Southeast), so it is not purely brand mix. "
            "Chart: regional bar and region x month heatmap.",
            f"4. Price-led growth in 2021. January to December: sales {pct(dc['sales_growth'])}, units "
            f"{pct(dc['units_growth'])}, realised sales per unit {dc['jan_rpu']:.2f} to {dc['dec_rpu']:.2f} "
            f"({pct(dc['price_growth'])}). Weighted margin moved only between "
            f"{pct(min(m['margin'] for m in f['months']))} and {pct(max(m['margin'] for m in f['months']))}. So the "
            "margin rate held while the revenue base grew.",
            f"5. Two different peaks. December is the highest sales month at {money(me['high_sales'])}; March the lowest "
            f"at {money(me['low_sales'])}; the difference is {money(me['diff'])}, or {pct(me['rel'])} of the March base. "
            f"July is the highest units month ({f['months'][6]['units']:,}). Chart: monthly sales and monthly margin "
            "trends.",
            f"6. Brand. Coca-Cola is the largest brand at {pct(coke['sales_share'])} of sales and "
            f"{pct(coke['profit_share'])} of profit, with a {pct(coke['margin'])} margin, "
            f"{(coke['margin']-t['margin'])*100:+.1f} pp vs portfolio. Dasani Water is second at "
            f"{pct(brd['Dasani Water']['sales_share'])} of sales and {pct(brd['Dasani Water']['profit_share'])} of profit, "
            f"with the highest profit per unit ({money(brd['Dasani Water']['profit_per_unit'],3)}).",
            f"7. Concentration. Top 5 cities: {pct(f['concentration']['top5'])} of sales; top 10: "
            f"{pct(f['concentration']['top10'])}; top 20: {pct(f['concentration']['top20'])}. "
            f"{mat['count']} of {mat['total']} cities clear $100,000 and produce {pct(mat['sales_share'])} of sales.",
            f"8. Weak margin in high-volume cells. Seattle {pct(hs['Washington | Seattle']['margin'])} on "
            f"{money(hs['Washington | Seattle']['sales'])}; Anchorage "
            f"{pct(hs['Alaska | Anchorage']['margin'])} on {money(hs['Alaska | Anchorage']['sales'])}; San Francisco "
            f"{pct(hs['California | San Francisco']['margin'])} on {money(hs['California | San Francisco']['sales'])}.",
        ]),
        ("Recommendations and decision implications", [
            "1. Investigate before acting on the Western margin gap. A matched-cell review of price and margin, starting "
            "with the four lowest-margin material cities, is cheap and reversible. Cutting price or volume on the "
            "strength of a cross-sectional margin comparison would not be justified by this data.",
            "2. Protect Sodapop scale and benchmark, not assume, the margin gap to BevCo. Sodapop generates more than "
            "half of profit; the priority is availability and service, with the margin gap treated as a question for the "
            "account teams.",
            "3. Attach a volume guardrail to pricing. Price-led growth is visible in the data, and the risk of trading "
            "volume for price is exactly what a controlled test would measure.",
            "4. Plan for two peaks, not one season. December needs price and availability execution; July needs volume "
            "capacity.",
            "5. Focus coverage on material cities and verify distribution before reading low sales as low demand.",
        ]),
        ("Limitations", [
            "One year only (2021). No prior-year actuals, so no year-on-year comparison and no proven seasonality; the "
            "patterns described are within-year.",
            "No cost, promotion, distribution, availability or store-count fields. Margin differences can be located but "
            "not explained. Recorded sales cannot be converted into productivity or efficiency.",
            "Retailer and region are almost perfectly confounded: BevCo operates only in the South, FizzySip only in the "
            "West, and Sodapop operates in the Northeast, Southeast and Midwest. Region and account effects cannot be "
            "separated.",
            "Grain is assumed to be retailer-city-brand-month; without an order identifier, record counts are not order "
            "counts and duplicate-free monthly keys do not prove order-level that completeness.",
            "Price-volume analysis is cross-sectional. A positive pooled correlation (+"
            f"{f['correlation']['pooled']:.2f}) is association, not evidence of elasticity.",
            "Currency is assumed to be USD and margin is 'operating margin' as supplied; neither definition is "
            "documented in the source file.",
        ]),
        ("Conclusion", [
            f"The dataset supports a clear, defensible answer: Coca-Cola's recorded 2021 sales of {money(t['sales'])} are "
            f"concentrated in Sodapop and in the West, but profit is concentrated in Sodapop and in higher-margin "
            f"southern and southeastern cells. The single largest opportunity the data identifies is not a price change "
            f"or a product cut; it is closing the information gap on why Western and high-volume cells earn "
            f"{abs((west['margin']-t['margin'])*100):.1f} to {(bev['margin']-fizzy['margin'])*100:.1f} percentage points "
            "less margin than the best comparable cells. The workbook makes that question answerable at city, brand and "
            "month level, and names the three data additions - cost, coverage and history - that would convert the "
            "investigation into a decision.",
        ]),
    ]
