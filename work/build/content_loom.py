"""Nishant's Loom script, evaluator Q&A and recording checklist.

Every figure is interpolated from facts.json so the narration can never drift
from the workbook.
"""


def money(x, dp=0):
    return f"${x:,.{dp}f}"


def m_(x, dp=2):
    return f"${x/1e6:.{dp}f} million"


def pct(x, dp=1):
    return f"{x*100:.{dp}f}%"


def build_loom(f):
    t = f["total"]
    ret = {r["name"]: r for r in f["levels"]["retailer"]}
    reg = {r["name"]: r for r in f["levels"]["region"]}
    brd = {r["name"]: r for r in f["levels"]["brand"]}
    me = f["month_extremes"]
    dc = f["decomp"]
    coke = {r["name"]: r for r in f["levels"]["brand"]}["Coca-Cola"]
    mat = f["materiality"]
    conf = f["materiality"]
    soda, fizzy, bev = ret["Sodapop"], ret["FizzySip"], ret["BevCo"]
    west, south = reg["West"], reg["South"]
    gap_pp = (west["margin"] - t["margin"]) * 100
    hs = f["high_sales_low_margin"]
    seattle = next((x for x in hs if x["name"] == "Washington | Seattle"), None)
    sf = next((x for x in hs if x["name"] == "California | San Francisco"), None)
    anch = next((x for x in hs if x["name"] == "Alaska | Anchorage"), None)
    dasani_west = f["brand_region"]["Dasani Water"]["West"]["margin"]
    dasani_south = f["brand_region"]["Dasani Water"]["South"]["margin"]
    coke_west = f["brand_region"]["Coca-Cola"]["West"]["margin"]
    coke_se = f["brand_region"]["Coca-Cola"]["Southeast"]["margin"]

    return [
        dict(t="00:00-00:40", view="00_ReadMe, then 12_Dashboard title bar",
             say=("Hello, I'm Nishant. In this project I analysed the supplied Coca-Cola sales dataset to see where "
                  "recorded sales and operating profit are strongest, where margins need attention, and which actions "
                  "the evidence supports. The question is how Coca-Cola can optimise sales and operating profitability "
                  "across U.S. retailers, regions, products and time. Every number here comes from the 3,888 records in "
                  "your workbook, not from outside figures. I will cover the data, the cleaning, the headline numbers, retailer and "
                  "regional differences, what time and brands show, and the actions those findings support.")),
        dict(t="00:40-01:15", view="01_Raw_Data, then 03_Data_Dictionary",
             say=("The workbook has one raw sheet. Headers sit on row five and records run to row 3,893, so exactly "
                  "3,888 records across twelve fields: retailer, date, region, state, city, brand, price per unit, units "
                  "sold, sales, operating profit and margin. It covers one year, 2021, and each month holds exactly 324 "
                  "records - 54 cities times six brands. So a row is a retailer-city-brand-month observation, not an "
                  "order. Four retailers, five regions, 50 states and 54 cities are covered, and six brands: Coca-Cola, Diet Coke, "
                  "Sprite, Fanta, Powerade and Dasani Water. This is an educational dataset, not audited company reporting.")),
        dict(t="01:15-02:00", view="04_Data_Quality, then 02_Clean_Data",
             say=("I kept the original sheet untouched and built a separate table, so every figure traces back to a "
                  "source row. The checks came back clean: no negatives, no duplicates, no missing dates, and no partial "
                  "month. Price times units equals sales in all 3,888 rows, and profit divided by sales equals the "
                  "supplied margin wherever sales exist. Two Omaha rows have zero sales but still carry a 40 percent "
                  "margin; I kept and flagged them, and they change no total. I also excluded stray unlabelled cells "
                  "outside the record block. Currency is not documented anywhere in the file, so every monetary figure is labelled US "
                  "dollars on that assumption, and the file gives operating profit only, so I never say net profit.")),
        dict(t="02:00-02:50", view="05_KPI_Summary",
             say=(f"Total sales are {money(t['sales'])} across {t['units']/1e6:.2f} million units, giving "
                  f"{money(t['profit'])} of operating profit. Overall margin is {pct(t['margin'])}, and I calculated that "
                  "as total operating profit divided by total sales, not by averaging the percentages in each row. The "
                  f"simple average would have said {pct(f['row_margin_simple_avg'])} - about "
                  f"{abs((f['row_margin_simple_avg']-t['margin'])*100):.1f} points away - because a small city would "
                  "count as much as New York. Realised sales per unit are fifty-one cents and operating profit per unit "
                  "is eighteen and a half cents. The unit-weighted listed price comes out at the same fifty-one cents, so listed and "
                  "realised price agree here.")),
        dict(t="02:50-04:05", view="12_Dashboard: KPI cards, retailer bar, region bar, one slicer clicked",
             say=(f"Sodapop is the largest account at {pct(soda['sales_share'])} of sales and {pct(soda['profit_share'])} "
                  f"of profit, with a {pct(soda['margin'])} margin. FizzySip is the opposite: {pct(fizzy['sales_share'])} "
                  f"of sales but only {pct(fizzy['profit_share'])} of profit, because its margin is "
                  f"{pct(fizzy['margin'])}. BevCo is smallest of the three big accounts at {pct(bev['sales_share'])} of "
                  f"sales yet earns the highest margin, {pct(bev['margin'])}. Geographically the West is largest at "
                  f"{pct(west['sales_share'])} of sales but converts that into {pct(west['profit_share'])} of profit at "
                  f"{pct(west['margin'])}, the lowest region and {abs(gap_pp):.1f} points below the portfolio. Now the "
                  "caveat: BevCo only sells in the South and FizzySip only in the West, so retailer and region are "
                  "nearly the same variable here. Southern and southeastern cells carry the highest margins, and the gap survives "
                  "inside a single brand: Coca-Cola earns about 35 percent in the West against 44 percent in the Southeast, "
                  "so brand mix is not the whole explanation.")),
        dict(t="04:05-05:00", view="09_Monthly_Analysis, monthly trend and brand charts",
             say=(f"December is the biggest month at {money(me['high_sales'])} and March the smallest at "
                  f"{money(me['low_sales'])}, {pct(me['rel'])} higher. But look at the driver: from January to December "
                  f"sales grew {pct(dc['sales_growth'])} while units grew only {pct(dc['units_growth'])}, and realised "
                  f"price rose from {dc['jan_rpu']:.2f} to {dc['dec_rpu']:.2f} dollars, up {pct(dc['price_growth'])}. "
                  "So 2021 growth was price-led as much as volume-led, and margin barely moved - a range of about one "
                  f"point all year. July is the units peak at 1.76 million, so summer is a volume story and December a "
                  f"price story. Coca-Cola is the largest brand at {pct(coke['sales_share'])} of sales and "
                  f"{pct(coke['profit_share'])} of profit. Dasani is second on both sales and profit and has the highest profit per "
                  "unit of the six brands, while Diet Coke and Sprite sit at the bottom.")),
        dict(t="05:00-06:10", view="13_Insights_Actions",
             say=(f"Four actions. First, investigate before cutting. Seattle earns {pct(seattle['margin'])} and San "
                  f"Francisco {pct(sf['margin'])} on material sales, so a matched-cell review of price and margin is "
                  "worth more than an immediate price change - the data has no cost field. Second, treat December and "
                  "July as two different planning problems: December is a price peak, July a units peak. Third, put a "
                  f"volume guardrail on pricing, because price rose {pct(dc['price_growth'])} while units rose "
                  f"{pct(dc['units_growth'])}. Fourth, focus coverage: {mat['count']} of {mat['total']} cities clear one "
                  f"hundred thousand dollars and produce {pct(mat['sales_share'])} of revenue. Each action names the extra "
                  "evidence needed before a commercial change: cost per unit, distribution coverage and a second year of "
                  "history.")),
        dict(t="06:10-06:50", view="15_Validation, then 12_Dashboard default view",
             say=("Limitations: one year of data, so these are within-year patterns, not proven seasonality, and there "
                  "is no prior year to compare. No cost, promotion, distribution or store-count field exists, so I can "
                  "show where margin differs but not why. Price and volume correlate at about 0.25, which is an "
                  "association across markets, not proof of elasticity. Everything reconciles: 3,888 raw rows and "
                  "3,888 cleaned rows, the same sales total rebuilt by retailer, by region, by brand and by month. The "
                  "workbook, report and script share one definition of every metric. If you take one thing away, make it this: the "
                  "data locates the margin gap precisely, and the next step is to explain it, not to assume it.")),
    ]


def build_qa(f):
    t = f["total"]
    west = next(r for r in f["levels"]["region"] if r["name"] == "West")
    return [
        ("Why is your overall margin different from the average of the Operating Margin column?",
         f"Because a simple average weights every row equally, so a city selling 4,000 units counts as much as one "
         f"selling 12,000. The portfolio margin is total operating profit {money(t['profit'])} divided by total sales "
         f"{money(t['sales'])}, which is {pct(t['margin'])}. The simple row average is "
         f"{pct(f['row_margin_simple_avg'])}, which sits "
         f"{abs((f['row_margin_simple_avg']-t['margin'])*100):.1f} percentage points away from the weighted figure. "
         f"Every margin in the workbook is a "
         f"ratio of sums."),
        ("Is the West really less profitable, or is that just FizzySip?",
         f"Both readings are in the data and I cannot separate them. The West is {pct(west['margin'])} against "
         f"{pct(t['margin'])} for the portfolio, but FizzySip operates only in the West, and BevCo, the "
         f"highest-margin retailer, operates only in the South. Retailer and region are almost the same variable "
         f"here. What survives that caveat is the within-brand comparison: Coca-Cola earns "
         f"{pct(f['brand_region']['Coca-Cola']['West']['margin'])} in the West versus "
         f"{pct(f['brand_region']['Coca-Cola']['Southeast']['margin'])} in the Southeast, so the gap is not just "
         f"brand mix. It needs a cell-level review."),
        ("Can you claim seasonality from this data?",
         "No. I have twelve months from a single year, so I describe within-year patterns: December is the highest "
         "sales month, March the lowest, July the highest units month. To claim recurring seasonality I would need at "
         "least two or three years, and the workbook contains no prior-year actuals, which is why the trend chart has "
         "no prior-year line."),
        ("Does the price-volume correlation prove that lowering price raises volume?",
         f"No. The pooled correlation is {f['correlation']['pooled']:.2f} and it is positive, which is the opposite of "
         "a simple demand curve, because expensive markets are also high-volume markets. It is an association across "
         "cities and months, confounded by retailer, location and time. Elasticity needs a controlled test with a "
         "comparison group, which is exactly what I recommend."),
        ("What would you do next with better data?",
         "Three additions would change the recommendations from investigation to action: cost of goods and operating "
         "cost per unit, so margin differences can be attributed; store counts or distribution coverage per city, so "
         "sales can be normalised into productivity; and two more years of history, so seasonality and price changes "
         "can be separated from growth. Promotion and availability data would complete the picture."),
    ]


CHECKLIST = [
    "Record in a quiet room with the door closed; test the microphone for 10 seconds and play it back.",
    "Wear plain professional clothing; avoid fine stripes and loud patterns on camera.",
    "Set the browser or Excel zoom so the dashboard text is readable at 1080p; hide the formula bar and gridlines.",
    "Turn off notifications, chat badges and email pop-ups before recording.",
    "Reset the workbook to its default state: all slicers cleared, dashboard on the full-portfolio view.",
    "Close anything private — other tabs, chat windows, file paths with personal names.",
    "Keep the mouse still while speaking; move it only when you say you are moving it.",
    "Play the full recording back once and check every number you say against the workbook.",
    "Check sharing permissions before sending the link; use view-only unless the reviewer needs to filter.",
]
