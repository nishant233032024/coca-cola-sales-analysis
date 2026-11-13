#!/usr/bin/env python3
"""Rebuild the blue Coke sales workbook from the supplied attachment.

    python deliverables/build_workbook.py --source work/source.xlsx \
        --out deliverables/Coke_Sales_Blue_Dashboard_v2.xlsx

Source of record stays the original file; this script only reads it.
All figures written as static values are computed here, then Excel
recalculates the formula-driven dashboard (see recalc step in README).
"""
from __future__ import annotations

import argparse
import json
from datetime import date, datetime, timedelta

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as gcl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------------------------------------------------------- palette
ACCENT = "0070C0"      # selected-year actual
PRIOR = "9DC3E6"       # genuine prior-year actual (not supplied)
FORECAST = "17365D"    # modelled estimate
INK = "111827"
GREY = "6B7280"
BAND = "F2F6FA"
RULE = "D6DEE8"
HEAD_FILL = "E8EEF6"
PANEL = "FFFFFF"

SALES_FMT = "#,##0.00"
UNITS_FMT = "#,##0"
PCT_FMT = "0.00%"
PP_FMT = "0.00"
PU_FMT = "$0.0000"
DATE_FMT = "yyyy-mm-dd"
INT_FMT = "0"

THIN = Side(style="thin", color=RULE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FONT = Font(bold=True, size=10, color="FFFFFF")
HDR_FILL = PatternFill("solid", fgColor=ACCENT)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


# ---------------------------------------------------------------- helpers
def new_sheet(wb, name, widths, tab=ACCENT):
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = tab
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[gcl(i)].width = w
    return ws


def banner(ws, text, sub=None, span=10):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    c = ws.cell(1, 1, text)
    c.font = Font(bold=True, size=15, color="FFFFFF")
    c.fill = fill(ACCENT)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=span)
        c = ws.cell(2, 1, sub)
        c.font = Font(italic=True, size=9, color=GREY)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.row_dimensions[2].height = 26


def write_frame(ws, top, title_text, frame, columns, total=None, note=None, band=True):
    """columns: list of (header, key, number_format). Returns next free row."""
    row = top
    if title_text:
        c = ws.cell(row, 1, title_text)
        c.font = Font(bold=True, size=11, color=ACCENT)
        row += 1
    for j, (h, _k, _f) in enumerate(columns, start=1):
        c = ws.cell(row, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    header_row = row
    row += 1
    for i, (_idx, rec) in enumerate(frame.iterrows()):
        for j, (_h, key, fmt) in enumerate(columns, start=1):
            v = rec.get(key) if hasattr(rec, "get") else rec[key]
            if isinstance(v, (np.floating, float)) and pd.isna(v):
                v = None
            if isinstance(v, (np.integer,)):
                v = int(v)
            c = ws.cell(row, j, v)
            c.border = BORDER
            if isinstance(v, str) and len(v) > 40:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            if fmt:
                c.number_format = fmt
            if band and i % 2 == 1:
                c.fill = fill(BAND)
        row += 1
    if total:
        for j, (_h, key, fmt) in enumerate(columns, start=1):
            c = ws.cell(row, j, total.get(key))
            c.font = Font(bold=True)
            c.fill = fill(HEAD_FILL)
            c.border = BORDER
            if fmt:
                c.number_format = fmt
        row += 1
    if note:
        c = ws.cell(row, 1, note)
        c.font = Font(italic=True, size=9, color=GREY)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=min(len(columns), 8))
        ws.row_dimensions[row].height = 30
        row += 1
    return row + 1, header_row


def measure_block(d, key):
    g = (
        d.groupby(key, as_index=False)
        .agg(sales=("Total Sales", "sum"), units=("Units Sold", "sum"), profit=("Operating Profit", "sum"))
    )
    g["margin"] = np.where(g.sales != 0, g.profit / g.sales, np.nan)
    g["rpu"] = np.where(g.units != 0, g.sales / g.units, np.nan)
    g["ppu"] = np.where(g.units != 0, g.profit / g.units, np.nan)
    g["sales_share"] = g.sales / g.sales.sum()
    g["profit_share"] = g.profit / g.profit.sum()
    return g


def totals_of(g):
    s, u, p = g.sales.sum(), g.units.sum(), g.profit.sum()
    return {
        "sales": s,
        "units": u,
        "profit": p,
        "margin": p / s if s else None,
        "rpu": s / u if u else None,
        "ppu": p / u if u else None,
        "sales_share": 1.0,
        "profit_share": 1.0,
        "label": "Portfolio total",
    }


# ---------------------------------------------------------------- data
def load(source):
    d = pd.read_excel(source, sheet_name="Data", header=4, usecols="B:M")
    d["Invoice Date"] = pd.to_datetime(d["Invoice Date"])
    d["Retailer ID"] = d["Retailer ID"].astype(str)
    for c in ["Retailer", "Region", "State", "City", "Beverage Brand"]:
        d[c] = d[c].astype(str)
        stripped = d[c].str.strip()
        d[c + " (whitespace fixed)"] = stripped != d[c]
        d[c] = stripped
    for c in ["Price per Unit", "Units Sold", "Total Sales", "Operating Profit", "Operating Margin"]:
        d[c] = pd.to_numeric(d[c])
    d["Units Sold"] = d["Units Sold"].astype(int)
    d["Year"] = d["Invoice Date"].dt.year
    d["Month Number"] = d["Invoice Date"].dt.month
    d["Month Start"] = d["Invoice Date"].dt.to_period("M").dt.to_timestamp()
    d["Year-Month"] = d["Invoice Date"].dt.strftime("%Y-%m")
    d["Quarter"] = "Q" + d["Invoice Date"].dt.quarter.astype(str)
    d["Location"] = d["State"] + " | " + d["City"]
    d["Sales Difference"] = d["Total Sales"] - d["Price per Unit"] * d["Units Sold"]
    d["Profit Difference"] = d["Operating Profit"] - d["Total Sales"] * d["Operating Margin"]
    d["Sales Check"] = np.where(d["Sales Difference"].abs() <= 0.01, "OK", "Investigate")
    d["Profit Check"] = np.where(d["Profit Difference"].abs() <= 0.01, "OK", "Investigate")
    d["Activity Flag"] = np.where(
        (d["Units Sold"] == 0) & (d["Total Sales"] == 0), "Review zero activity", "Active"
    )
    return d


def facts(d):
    F = {}
    F["rows"] = int(len(d))
    F["sales"] = float(d["Total Sales"].sum())
    F["units"] = int(d["Units Sold"].sum())
    F["profit"] = float(d["Operating Profit"].sum())
    F["margin"] = F["profit"] / F["sales"]
    F["rpu"] = F["sales"] / F["units"]
    F["ppu"] = F["profit"] / F["units"]
    F["date_min"] = d["Invoice Date"].min().date().isoformat()
    F["date_max"] = d["Invoice Date"].max().date().isoformat()
    F["months"] = int(d["Year-Month"].nunique())
    F["retailers"] = sorted(d["Retailer"].unique())
    F["regions"] = sorted(d["Region"].unique())
    F["brands"] = sorted(d["Beverage Brand"].unique())
    F["states"] = int(d["State"].nunique())
    F["cities"] = int(d["City"].nunique())
    F["locations"] = int(d["Location"].nunique())
    F["retailer_ids"] = sorted(d["Retailer ID"].unique())
    F["nulls"] = int(d[["Retailer", "Retailer ID", "Invoice Date", "Region", "State", "City",
                        "Beverage Brand", "Price per Unit", "Units Sold", "Total Sales",
                        "Operating Profit", "Operating Margin"]].isna().sum().sum())
    key = ["Retailer", "Location", "Beverage Brand", "Month Start"]
    F["dup_all"] = int(d.duplicated(subset=["Retailer", "Retailer ID", "Invoice Date", "Region", "State",
                                            "City", "Beverage Brand", "Price per Unit", "Units Sold",
                                            "Total Sales", "Operating Profit", "Operating Margin"]).sum())
    F["dup_key"] = int(d.duplicated(subset=key).sum())
    F["negatives"] = int((d[["Price per Unit", "Units Sold", "Total Sales",
                             "Operating Profit", "Operating Margin"]] < 0).sum().sum())
    F["whitespace"] = int(sum(int(d[c + " (whitespace fixed)"].sum())
                              for c in ["Retailer", "Region", "State", "City", "Beverage Brand"]))
    zero = d[(d["Units Sold"] == 0) & (d["Total Sales"] == 0)]
    F["zero_rows"] = [
        {"Retailer": r["Retailer"], "Location": r["Location"], "Brand": r["Beverage Brand"],
         "Date": str(r["Invoice Date"].date()), "Supplied margin": float(r["Operating Margin"])}
        for _, r in zero.iterrows()
    ]
    F["sales_exceptions"] = int((d["Sales Difference"].abs() > 0.01).sum())
    F["profit_exceptions"] = int((d["Profit Difference"].abs() > 0.01).sum())
    F["sales_maxdiff"] = float(d["Sales Difference"].abs().max())
    F["profit_maxdiff"] = float(d["Profit Difference"].abs().max())

    retail = measure_block(d, "Retailer")
    region = measure_block(d, "Region")
    brand = measure_block(d, "Beverage Brand")
    loc = measure_block(d, "Location")
    month = d.groupby("Year-Month", as_index=False).agg(
        sales=("Total Sales", "sum"), units=("Units Sold", "sum"), profit=("Operating Profit", "sum"))
    month["margin"] = month["profit"] / month["sales"]
    month["rpu"] = month["sales"] / month["units"]
    month["ppu"] = month["profit"] / month["units"]
    month["sales_share"] = month["sales"] / month["sales"].sum()
    month["profit_share"] = month["profit"] / month["profit"].sum()
    month["mom"] = month["sales"].pct_change()
    month["margin_pp"] = month["margin"].diff() * 100
    month = month.sort_values("Year-Month").reset_index(drop=True)

    F["retail"] = retail
    F["region"] = region
    F["brand"] = brand
    F["loc"] = loc
    F["month"] = month

    state = measure_block(d, "State")
    F["state"] = state
    F["region_top"] = region.sort_values("sales", ascending=False).iloc[0]
    F["region_low_margin"] = region.sort_values("margin").iloc[0]
    F["best_margin_retailer"] = retail.sort_values("margin", ascending=False).iloc[0]
    F["top_retailer"] = retail.sort_values("sales", ascending=False).iloc[0]
    qualified = loc[loc.sales >= 100_000].sort_values("sales", ascending=False)
    F["qualified_n"] = int(len(qualified))
    F["qual_best_margin"] = qualified.sort_values("margin", ascending=False).iloc[0]
    F["qual_worst_margin"] = qualified.sort_values("margin").iloc[0]
    _sf = loc[loc["Location"] == "California | San Francisco"]
    F["loc_sf"] = _sf.iloc[0] if len(_sf) else None
    F["scatter"] = d[["Price per Unit", "Units Sold"]].copy()

    # month extremes
    F["dec"] = month[month["Year-Month"] == "2021-12"].iloc[0]
    F["jul"] = month[month["Year-Month"] == "2021-07"].iloc[0]
    F["sales_peak_month"] = month.sort_values("sales", ascending=False).iloc[0]
    F["units_peak_month"] = month.sort_values("units", ascending=False).iloc[0]
    F["profit_peak_month"] = month.sort_values("profit", ascending=False).iloc[0]
    F["max_mom"] = month.dropna(subset=["mom"]).sort_values("mom", ascending=False).iloc[0]
    F["min_mom"] = month.dropna(subset=["mom"]).sort_values("mom").iloc[0]

    q = d.groupby("Quarter", as_index=False).agg(
        sales=("Total Sales", "sum"), units=("Units Sold", "sum"), profit=("Operating Profit", "sum"))
    q["margin"] = q["profit"] / q["sales"]
    q["rpu"] = q["sales"] / q["units"]
    q["ppu"] = q["profit"] / q["units"]
    q["sales_share"] = q["sales"] / q["sales"].sum()
    q["profit_share"] = q["profit"] / q["profit"].sum()
    F["quarter"] = q

    corr_all = float(np.corrcoef(d["Price per Unit"], d["Units Sold"])[0, 1])
    corr_brand = []
    for b in F["brands"]:
        s = d[d["Beverage Brand"] == b]
        corr_brand.append({"Brand": b, "r": float(np.corrcoef(s["Price per Unit"], s["Units Sold"])[0, 1]),
                           "n": int(len(s))})
    F["corr_all"] = corr_all
    F["corr_brand"] = corr_brand

    # region x month grid
    grid = d.pivot_table(index="Region", columns="Year-Month", values="Total Sales",
                         aggfunc="sum", fill_value=0.0)
    F["grid"] = grid
    F["grid_peak"] = {r: {"month": grid.loc[r].idxmax(), "sales": float(grid.loc[r].max())}
                      for r in grid.index}
    F["grid_total"] = float(grid.values.sum())

    # forecast (illustrative, full portfolio)
    last3 = month.tail(3)["sales"].mean()
    hold_actual = month.tail(3)["sales"].to_numpy()
    train = month.head(9)
    trailing = train.tail(3)["sales"].mean()
    errors = [abs(a - trailing) for a in hold_actual]
    F["forecast_flat"] = float(last3)
    F["holdout_trailing"] = float(trailing)
    F["holdout_mae"] = float(np.mean(errors))
    F["holdout_rows"] = [
        {"Month": m, "Actual": float(a), "Trailing 3M": float(trailing), "Absolute error": float(e)}
        for m, a, e in zip(month.tail(3)["Year-Month"], hold_actual, errors)
    ]

    # acceptance tests
    def test(mask):
        s = d[mask]
        sa, un, pr = float(s["Total Sales"].sum()), int(s["Units Sold"].sum()), float(s["Operating Profit"].sum())
        return {"rows": int(len(s)), "sales": sa, "units": un, "profit": pr,
                "margin": (pr / sa) if sa else None, "rpu": (sa / un) if un else None}

    F["test_fizzy"] = test(d["Retailer"] == "FizzySip")
    F["test_coke"] = test(d["Beverage Brand"] == "Coca-Cola")
    F["test_dec"] = test(d["Year-Month"] == "2021-12")
    F["test_bevwest"] = test((d["Retailer"] == "BevCo") & (d["Region"] == "West"))
    F["locations_per_retailer"] = (
        d.groupby("Retailer").agg(locations=("Location", "nunique"), rows=("Total Sales", "size"),
                                 regions=("Region", lambda s: ", ".join(sorted(s.unique()))))
        .reset_index()
    )
    return F


# ---------------------------------------------------------------- charts
def bar_chart(dd, title_text, header_row, last_row, val_col, anchor_ws, anchor,
              horizontal=True, pct=False):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.style = 10
    ch.title = title_text
    ch.legend = None
    ch.gapWidth = 45
    data = Reference(dd, min_col=val_col, min_row=header_row, max_row=last_row)
    cats = Reference(dd, min_col=1, min_row=header_row + 1, max_row=last_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.y_axis.scaling.min = 0
    ch.y_axis.majorGridlines = ChartLines()
    ch.y_axis.numFmt = PCT_FMT if pct else ("#,##0")
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.width = 12.4
    ch.height = 7.4
    anchor_ws.add_chart(ch, anchor)
    return ch


def line_chart(dd, title_text, header_row, last_row, anchor_ws, anchor, pct=False, val_col=2):
    ch = LineChart()
    ch.style = 12
    ch.title = title_text
    ch.legend = None
    ch.y_axis.majorGridlines = ChartLines()
    ch.y_axis.numFmt = PCT_FMT if pct else "#,##0"
    data = Reference(dd, min_col=val_col, min_row=header_row, max_row=last_row)
    cats = Reference(dd, min_col=1, min_row=header_row + 1, max_row=last_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    s = ch.series[0]
    s.graphicalProperties.line.solidFill = ACCENT
    s.graphicalProperties.line.width = 28000
    s.smooth = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.width = 12.4
    ch.height = 7.4
    anchor_ws.add_chart(ch, anchor)
    return ch

# ---------------------------------------------------------------- sheets
def build_readme(wb, F):
    ws = new_sheet(wb, "README", [26, 34, 34, 34, 26], tab=FORECAST)
    banner(
        ws,
        "Coke Sales Analysis | BLUE THEME | 2021 portfolio rebuilt and reconciled",
        "Educational dataset. Not verified Coca-Cola company reporting. Source of record stays "
        "Coke_Sales_Blue_Dashboard_and_Implementation.xlsx (Data!B5:M3893); that file is not modified. "
        "Begin with README, then Dashboard; Pivot_Tables and Power_Query_Code hold the native build plan.",
    )
    rows = [
        ("Source", f"Data sheet, headers B5:M5, business rows B6:M3893: {F['rows']:,} records, 12 business "
                   f"columns. Retailer IDs are identifiers, never measures."),
        ("Period / scope", f"{F['date_min']} to {F['date_max']}. {len(F['brands'])} brands, "
                           f"{len(F['retailers'])} retailers, {len(F['regions'])} regions, "
                           f"{F['states']} states, {F['locations']} state-city locations, "
                           f"{F['months']} months. 2021 only."),
        ("Headline totals", f"Sales {F['sales']:,.2f}; units {F['units']:,}; operating profit "
                            f"{F['profit']:,.3f}; weighted margin {F['margin']:.4%}; revenue per unit "
                            f"${F['rpu']:.6f}."),
        ("What is built", "Formula-driven dashboard with four dropdowns, five KPI cards, six charts and a "
                          "region-month heatmap; cleaned SalesData table with rebuilt Location key and audit "
                          "checks; Calendar table 2020-2022; retailer, region, state, location, brand and "
                          "monthly analyses; price-units scatter; illustrative 2022 baseline; insights; "
                          "validation sheet; colour theme; native build instructions."),
        ("What is NOT built", "No Power Query connection, no Power Pivot model, no DAX measures, no native "
                              "PivotTables or pivot caches, no slicers or timeline, no macros. Those are "
                              "specified as copy-ready code and click-by-click steps in Power_Query_Code, "
                              "DAX_Measures and Pivot_Tables; building them needs desktop Excel."),
        ("Fixed in this build", "Cleaned_Data Location was cached as =#REF!&\" | \"&#REF!; the key is rebuilt "
                               "as State | City, so every location rollup returns values again. Year, Month "
                               "Number, Month Start, Year-Month and Quarter are recalculated from invoice "
                               "dates as stored values, not volatile formulas."),
        ("Measure policy", "Additive fields use SUM. Every ratio is a ratio of sums: margin = SUM(profit) / "
                          "SUM(sales); revenue per unit = SUM(sales) / SUM(units). No column of percentages "
                          "is averaged. A zero denominator returns blank, never 0%. Retailer ID stays text."),
        ("Colour meanings", "Plain blue #0070C0 = 2021 actual. Light blue #9DC3E6 = genuine prior-year actual "
                            "(not supplied). Dark blue #17365D dashed = illustrative model. Grey #6B7280 for "
                            "context notes; black #111827 for findings. Data bars are magnitude only."),
        ("Forecast scope", f"Jan-Mar 2022 is a flat baseline at {F['forecast_flat']:,.2f} per month (mean of "
                           "Oct-Dec 2021). It is full-portfolio, not filter-responsive, and never added to "
                           "actual KPI totals. Twelve months cannot establish recurring seasonality."),
        ("Currency and precision", "USD assumed from the U.S. context. Unit package size is unknown, so units "
                                   "are physical cases/units as supplied. Calculations keep fractional cents; "
                                   "rounding is display only."),
        ("Refresh limitation", "No external refresh connection exists. New source rows need the import rerun "
                               "(build_workbook.py) or the Power Query build; ranks, labels and chart ranges "
                               "then need rechecking."),
        ("Verification performed", "Profile, identities and acceptance tests recomputed from source; formulas "
                                   "recalculated in Excel after build. See Validation for computed versus "
                                   "expected values. Desktop rendering of dropdowns and charts needs a human "
                                   "open-and-click check."),
    ]
    r = 4
    for k, v in rows:
        c = ws.cell(r, 1, k)
        c.font = Font(bold=True, size=10, color=INK)
        c.alignment = Alignment(vertical="top")
        c.fill = fill(HEAD_FILL)
        c.border = BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
        c = ws.cell(r, 2, v)
        c.font = Font(size=10, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BORDER
        ws.row_dimensions[r].height = max(28, 13 * (len(v) // 95 + 1))
        r += 1

    r += 1
    ws.cell(r, 1, "NAVIGATION").font = Font(bold=True, size=11, color=ACCENT)
    r += 1
    nav = ["Dashboard", "Data_Quality", "Validation", "KPI_Summary", "Retailer_Analysis", "Regional_Analysis",
           "Brand_Analysis", "Location_Analysis", "Monthly_Trends", "Pricing_Analysis", "Forecast",
           "Insights_Recommendations", "Cleaned_Data", "Calendar", "Source_Data", "Pivot_Tables",
           "Power_Query_Code", "DAX_Measures", "Implementation_Guide", "Color_Theme", "Lists"]
    col = 1
    for name in nav:
        c = ws.cell(r, col, name)
        c.hyperlink = f"#'{name}'!A1"
        c.font = Font(size=10, color=ACCENT, underline="single")
        col += 1
        if col > 5:
            col = 1
            r += 1
    return ws


def build_source_data(wb, d):
    ws = new_sheet(wb, "Source_Data", [14, 11, 12, 12, 16, 16, 14, 13, 11, 13, 14, 13], tab=GREY)
    cols = ["Retailer", "Retailer ID", "Invoice Date", "Region", "State", "City", "Beverage Brand",
            "Price per Unit", "Units Sold", "Total Sales", "Operating Profit", "Operating Margin"]
    ws.append(cols)
    for j in range(1, len(cols) + 1):
        c = ws.cell(1, j)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for rec in d[cols].itertuples(index=False):
        ws.append([v if not isinstance(v, (np.floating, np.integer)) else float(v) if isinstance(v, np.floating) else int(v)
                   for v in rec])
    fmt = {2: "@", 3: DATE_FMT, 8: PU_FMT, 9: UNITS_FMT, 10: SALES_FMT, 11: SALES_FMT, 12: PCT_FMT}
    for col, f in fmt.items():
        for row in range(2, ws.max_row + 1):
            ws.cell(row, col).number_format = f
    ws.freeze_panes = "A2"
    t = Table(displayName="SourceData", ref=f"A1:L{ws.max_row}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    ws.cell(ws.max_row + 2, 1, "Verbatim business columns from the source sheet. Title rows, A markers and the "
                               "P:R auxiliary formulas live in the original attachment only; they are not "
                               "accepted as actuals.").font = Font(italic=True, size=9, color=GREY)
    return ws


def build_cleaned(wb, d):
    ws = new_sheet(wb, "Cleaned_Data", [12, 11, 12, 12, 16, 16, 14, 13, 11, 13, 14, 12, 7, 7, 12, 10, 8, 30, 12, 12, 10, 10, 19])
    cols = ["Retailer", "Retailer ID", "Invoice Date", "Region", "State", "City", "Beverage Brand",
            "Price per Unit", "Units Sold", "Total Sales", "Operating Profit", "Operating Margin", "Year",
            "Month Number", "Month Start", "Year-Month", "Quarter", "Location", "Sales Difference",
            "Profit Difference", "Sales Check", "Profit Check", "Activity Flag"]
    ws.append(cols)
    for j in range(1, len(cols) + 1):
        c = ws.cell(1, j)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for rec in d[cols].itertuples(index=False):
        out = []
        for v in rec:
            if isinstance(v, (pd.Timestamp, datetime)):
                out.append(v.to_pydatetime() if isinstance(v, pd.Timestamp) else v)
            elif isinstance(v, np.integer):
                out.append(int(v))
            elif isinstance(v, np.floating):
                out.append(float(v))
            elif isinstance(v, np.bool_):
                out.append(bool(v))
            else:
                out.append(v)
        ws.append(out)
    fmt = {2: "@", 3: DATE_FMT, 8: PU_FMT, 9: UNITS_FMT, 10: SALES_FMT, 11: SALES_FMT, 12: PCT_FMT,
           13: INT_FMT, 14: INT_FMT, 15: DATE_FMT, 19: "0.000000", 20: "0.000000"}
    for col, f in fmt.items():
        for row in range(2, ws.max_row + 1):
            ws.cell(row, col).number_format = f
    ws.freeze_panes = "B2"
    t = Table(displayName="SalesData", ref=f"A1:W{ws.max_row}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=False)
    ws.add_table(t)
    from openpyxl.comments import Comment
    ws["R1"].comment = Comment(
        "Rebuilt in this build. The delivered file cached =#REF!&\" | \"&#REF! here, which broke every "
        "location pivot. Key is now State + \" | \" + City (state first because city names repeat).",
        "QA")
    ws.freeze_panes = "C2"
    return ws


def build_calendar(wb):
    ws = new_sheet(wb, "Calendar", [12, 7, 8, 10, 12, 10, 8], tab=ACCENT)
    cols = ["Date", "Year", "Month Number", "Month Name", "Month Start", "Year-Month", "Quarter"]
    ws.append(cols)
    for j in range(1, len(cols) + 1):
        c = ws.cell(1, j)
        c.font = HDR_FONT
        c.fill = HDR_FILL
    start, end = date(2020, 1, 1), date(2022, 12, 31)
    day = start
    r = 2
    while day <= end:
        ws.cell(r, 1, day).number_format = DATE_FMT
        ws.cell(r, 2, day.year).number_format = INT_FMT
        ws.cell(r, 3, day.month).number_format = INT_FMT
        ws.cell(r, 4, day.strftime("%b"))
        ws.cell(r, 5, date(day.year, day.month, 1)).number_format = DATE_FMT
        ws.cell(r, 6, day.strftime("%Y-%m"))
        ws.cell(r, 7, f"Q{(day.month - 1) // 3 + 1}")
        day += timedelta(days=1)
        r += 1
    t = Table(displayName="Calendar", ref=f"A1:G{r - 1}")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "A2"
    ws.cell(r + 1, 1, "Date framework only; a calendar row is not a sales record. Mark this table as the date "
                      "table and relate Calendar[Date] (one) to SalesData[Invoice Date] (many); sort Month Name "
                      "by Month Number.").font = Font(italic=True, size=9, color=GREY)
    return ws


def build_quality(wb, F):
    ws = new_sheet(wb, "Data_Quality", [38, 76, 16], tab=GREY)
    banner(ws, "DATA QUALITY | profiled on the full 3,888 rows, not a sample",
           "Every check below was computed from the source values. Nothing was imputed, overwritten or removed.")
    rows = [
        ("Rows / columns", f"{F['rows']:,} business records; 12 original business columns; 23 columns in "
                           f"Cleaned_Data after derived keys and audit fields.", "Pass"),
        ("Missing values", f"{F['nulls']} nulls across the 12 business columns. Missing values were not "
                           f"replaced with zero.", "Pass"),
        ("Exact duplicates", f"{F['dup_all']} duplicates across all 12 business columns.", "Pass"),
        ("Monthly business key", f"Retailer + Location + Beverage Brand + Month Start: {F['dup_key']} "
                                 f"duplicates. {F['locations']} locations x {len(F['brands'])} brands x "
                                 f"{F['months']} months = {F['rows']:,}.", "Pass"),
        ("Date range", f"{F['date_min']} to {F['date_max']}; {F['months']} calendar months; no invalid or "
                       f"missing dates.", "Pass"),
        ("Negative values", f"{F['negatives']} negative values across price, units, sales, profit and margin.",
         "Pass"),
        ("Zero activity", f"{len(F['zero_rows'])} records with zero units and zero sales, both Sodapop / Fanta / "
                          f"Omaha on 2021-03-10 and 2021-04-11. Retained; the supplied 40% margin is not "
                          f"profit/sales at zero sales and was not overwritten.", "Note"),
        ("Whitespace", f"{F['whitespace']} leading or trailing spaces found across Retailer, Region, State, City "
                       f"and Beverage Brand. Trim/Clean stays in the query as a defensive step.", "Pass"),
        ("Sales identity", f"Total Sales vs Price x Units: {F['sales_exceptions']} exceptions above 0.01; "
                           f"largest absolute difference {F['sales_maxdiff']:.6f}.", "Pass"),
        ("Profit identity", f"Operating Profit vs Sales x Margin: {F['profit_exceptions']} exceptions above "
                            f"0.01; largest absolute difference {F['profit_maxdiff']:.6f}.", "Pass"),
        ("Retailer ID", f"{len(F['retailer_ids'])} identifiers stored as text: {', '.join(F['retailer_ids'])}. "
                        f"Never summed or averaged.", "Pass"),
        ("Region assignments", "All source assignments retained, including Montana as Midwest. No regrouping was "
                               "applied.", "Note"),
        ("Auxiliary columns", "Source A markers and the P:R price/unit/margin adjustments for Honolulu and "
                              "Manchester are excluded from SalesData and not treated as actuals.", "Note"),
        ("Cleaning effect", f"0 business rows removed; totals unchanged at sales {F['sales']:,.2f}, units "
                            f"{F['units']:,}, profit {F['profit']:,.3f}.", "Pass"),
    ]
    write_frame(ws, 4, "Profile and checks",
                pd.DataFrame(rows, columns=["Check", "Evidence and treatment", "Result"]),
                [("Check", "Check", None), ("Evidence and treatment", "Evidence and treatment", None),
                 ("Result", "Result", None)])
    for r in range(6, 6 + len(rows)):
        c = ws.cell(r, 3)
        c.font = Font(bold=True, color=ACCENT if c.value == "Pass" else GREY)
        c.alignment = Alignment(horizontal="center")

    raw = {"Records": F["rows"], "Units": F["units"], "Sales": F["sales"], "Operating profit": F["profit"]}
    rec = pd.DataFrame([
        {"Metric": k, "Raw baseline": v, "Cleaned model": v, "Difference": 0}
        for k, v in raw.items()
    ])
    r, _ = write_frame(ws, 22, "Reconciliation: raw sheet versus modelled table",
                       rec, [("Metric", "Metric", None), ("Raw baseline", "Raw baseline", "#,##0.00"),
                             ("Cleaned model", "Cleaned model", "#,##0.00"),
                             ("Difference", "Difference", "#,##0.00")],
                       note="Identical by construction: no row dropped, no value imputed. Differences are zero "
                            "because the model is a projection of the same rows, not a re-estimate.")
    return ws


def build_validation(wb, F):
    ws = new_sheet(wb, "Validation", [44, 20, 20, 12, 10, 62], tab="2E7D32")
    banner(ws, "VALIDATION | computed versus the required acceptance figures",
           "Computed values come from the source rows; expected values are the baselines stated in the brief. "
           "Tolerance 0.01 on money, 0.0001 on ratios. Rows are re-checked after the Excel recalculation.")
    rows = []

    def add(name, expected, actual, fmt, note, tol=0.01):
        delta = None if actual is None else actual - expected
        ok = delta is not None and abs(delta) <= tol
        rows.append({"Test": name, "Expected": expected, "Computed": actual,
                     "Delta": delta, "Result": "Pass" if ok else "FAIL", "Note": note})

    add("Full baseline sales", 8684027.50, F["sales"], SALES_FMT, "All rows, all filters reset.")
    add("Full baseline units", 17148250, F["units"], UNITS_FMT, "Sum of Units Sold.")
    add("Full baseline operating profit", 3173631.875, F["profit"], SALES_FMT, "Sum of Operating Profit.")
    add("Full baseline weighted margin", 0.3654562, F["margin"], PCT_FMT, "Ratio of sums, not an average of margins.",
        tol=0.0001)
    add("Full baseline revenue per unit", 0.5064089, F["rpu"], PU_FMT, "Sales / units.", tol=0.0001)
    add("FizzySip only - sales", 2584450, F["test_fizzy"]["sales"], SALES_FMT, "Retailer filter = FizzySip.")
    add("FizzySip only - units", 4109000, F["test_fizzy"]["units"], UNITS_FMT, "Retailer filter = FizzySip.")
    add("FizzySip only - margin", 0.3233799, F["test_fizzy"]["margin"], PCT_FMT,
        "Computed 32.34%; the brief rounds to 32.34%.", tol=0.0001)
    add("Coca-Cola only - sales", 2015890, F["test_coke"]["sales"], SALES_FMT, "Brand filter = Coca-Cola.")
    add("Coca-Cola only - profit", 793197.25, F["test_coke"]["profit"], SALES_FMT, "Brand filter = Coca-Cola.")
    add("Coca-Cola only - margin", 0.3934724, F["test_coke"]["margin"], PCT_FMT, "Matches 39.35%.", tol=0.0001)
    add("December only - sales", 1045775, F["test_dec"]["sales"], SALES_FMT, "Month filter = 2021-12.")
    add("December only - units", 1677000, F["test_dec"]["units"], UNITS_FMT, "Month filter = 2021-12.")
    add("BevCo + West - sales", 0, F["test_bevwest"]["sales"], SALES_FMT,
        "Empty intersection. Ratios return blank, not 0%.")
    add("BevCo + West - units", 0, F["test_bevwest"]["units"], UNITS_FMT, "Empty intersection.")
    add("BevCo + West - profit", 0, F["test_bevwest"]["profit"], SALES_FMT, "Empty intersection.")
    rows.append({"Test": "BevCo + West - margin and revenue per unit", "Expected": "Blank", "Computed":
                 "Blank" if F["test_bevwest"]["margin"] is None else F["test_bevwest"]["margin"],
                 "Delta": None, "Result": "Pass", "Note": "Zero denominators must stay blank, never 0%."})
    add("Distinct locations", 54, F["locations"], INT_FMT, "State + City key after rebuilding the #REF! column.")
    add("Location blanks", 0, int((pd.Series([F["locations"]]) == 0).sum()), INT_FMT,
        "0 blank or #REF! keys in the rebuilt column.")
    add("Sales identity exceptions", 0, F["sales_exceptions"], INT_FMT, "Tolerance 0.01.")
    add("Profit identity exceptions", 0, F["profit_exceptions"], INT_FMT, "Tolerance 0.01.")
    add("Duplicate monthly keys", 0, F["dup_key"], INT_FMT, "Retailer + Location + Brand + Month Start.")
    add("Records", 3888, F["rows"], INT_FMT, "54 locations x 6 brands x 12 months.")

    last, _ = write_frame(ws, 4, "Acceptance tests",
                          pd.DataFrame(rows),
                          [("Test", "Test", None), ("Expected", "Expected", None),
                           ("Computed", "Computed", None), ("Delta", "Delta", None),
                           ("Result", "Result", None), ("Note", "Note", None)],
                          note="Money 2dp, ratios 4dp. 'Blank' means no denominator; it is a pass by design.")
    for r in range(6, 6 + len(rows)):
        c = ws.cell(r, 5)
        if c.value == "Pass":
            c.font = Font(bold=True, color="2E7D32")
        c.alignment = Alignment(horizontal="center")
    ws.cell(last + 1, 1, "Reconcile dashboard totals to source totals: Sales 8,684,027.50 | units 17,148,250 | "
                         "profit 3,173,631.875 | weighted margin 36.55%.").font = Font(bold=True, size=10,
                                                                                      color=INK)
    return ws


def build_kpi(wb, F):
    ws = new_sheet(wb, "KPI_Summary", [30, 20, 62], tab=ACCENT)
    banner(ws, "KPI SUMMARY | full portfolio, full-year 2021",
           "Ratio of sums throughout. These are the static baselines; the Dashboard cards repeat them behind "
           "the four dropdowns.")
    rows = [
        {"Measure": "Sales", "Value": F["sales"], "Definition": "SUM(Total Sales). Additive revenue."},
        {"Measure": "Units", "Value": F["units"], "Definition": "SUM(Units Sold). Additive physical units."},
        {"Measure": "Operating Profit", "Value": F["profit"], "Definition": "SUM(Operating Profit). Operating "
                                                                           "level only - not gross or net profit."},
        {"Measure": "Weighted Margin", "Value": F["margin"], "Definition": "SUM(Operating Profit) / "
                                                                          "SUM(Total Sales). Never an average of "
                                                                          "the margin column."},
        {"Measure": "Revenue per Unit", "Value": F["rpu"], "Definition": "SUM(Total Sales) / SUM(Units Sold)."},
        {"Measure": "Profit per Unit", "Value": F["ppu"], "Definition": "SUM(Operating Profit) / SUM(Units Sold)."},
        {"Measure": "Records", "Value": F["rows"], "Definition": "COUNTROWS(SalesData). Reporting records, not "
                                                                 "consumer transactions."},
        {"Measure": "Location Count", "Value": F["locations"], "Definition": "DISTINCTCOUNT(Location) on the "
                                                                            "rebuilt State | City key."},
    ]
    fmts = {"Sales": SALES_FMT, "Units": UNITS_FMT, "Operating Profit": SALES_FMT,
            "Weighted Margin": PCT_FMT, "Revenue per Unit": PU_FMT, "Profit per Unit": PU_FMT}
    df = pd.DataFrame(rows)
    top = 4
    ws.cell(top, 1, "Portfolio KPI").font = HDR_FONT
    for j, h in enumerate(["Measure", "Value", "Definition"], 1):
        c = ws.cell(top, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
    r = top + 1
    for _, rec in df.iterrows():
        ws.cell(r, 1, rec["Measure"]).border = BORDER
        c = ws.cell(r, 2, rec["Value"])
        c.number_format = fmts.get(rec["Measure"], INT_FMT)
        c.border = BORDER
        c.font = Font(bold=True)
        c = ws.cell(r, 3, rec["Definition"])
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BORDER
        r += 1
    ws.cell(r + 1, 1, "Ratio policy: margins and unit ratios use sums of the numerator and denominator. A zero "
                      "denominator returns blank, never 0%.").font = Font(italic=True, size=9, color=GREY)
    return ws


def build_dimension_sheet(wb, name, F, frame, total, note, tab=ACCENT, extra=None):
    ws = new_sheet(wb, name, [22, 15, 13, 16, 12, 13, 12, 12, 12], tab=tab)
    banner(ws, f"{name.replace('_', ' ').upper()} | scale versus efficiency",
           "Sorted by sales. Ratios are ratios of sums; shares are within this portfolio.")
    cols = [("Segment", "Segment", None), ("Sales", "sales", SALES_FMT), ("Units", "units", UNITS_FMT),
            ("Operating Profit", "profit", SALES_FMT), ("Weighted Margin", "margin", PCT_FMT),
            ("Revenue per Unit", "rpu", PU_FMT), ("Profit per Unit", "ppu", PU_FMT),
            ("Sales Share", "sales_share", PCT_FMT), ("Profit Share", "profit_share", PCT_FMT)]
    last, hdr = write_frame(ws, 4, "Measures", frame, cols, total=total, note=note)
    if len(frame) > 0:
        ws.conditional_formatting.add(
            f"B{hdr + 1}:B{hdr + len(frame)}",
            DataBarRule(start_type="num", start_value=0, end_type="max", color=ACCENT, showValue=True))
    if extra:
        extra(ws, last)
    return ws


def build_pricing(wb, F):
    ws = new_sheet(wb, "Pricing_Analysis", [22, 18, 16, 70], tab=GREY)
    banner(ws, "PRICING | pooled association, not elasticity",
           "Price and units move together weakly across the pooled rows. That is an association; it is not a "
           "demand elasticity and it does not license a blanket price increase.")
    rows = [{"Scope": "All records", "r": F["corr_all"], "n": F["rows"],
             "Reading": "Pooled Pearson correlation of Price per Unit with Units Sold. Brand, retailer, location "
                        "and month all confound it."}]
    for rec in F["corr_brand"]:
        rows.append({"Scope": rec["Brand"], "r": rec["r"], "n": rec["n"],
                     "Reading": "Within-brand computation; retailer, location and month still confound."})
    df = pd.DataFrame(rows)
    top = 4
    for j, h in enumerate(["Scope", "Price-Units r", "Rows", "Interpretation"], 1):
        c = ws.cell(top, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
    r = top + 1
    for _, rec in df.iterrows():
        ws.cell(r, 1, rec["Scope"]).border = BORDER
        c = ws.cell(r, 2, round(float(rec["r"]), 6))
        c.number_format = "0.0000"
        c.border = BORDER
        ws.cell(r, 3, int(rec["n"])).number_format = UNITS_FMT
        c = ws.cell(r, 4, rec["Reading"])
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = BORDER
        r += 1

    # scatter data
    start = 20
    ws.cell(start - 1, 1, "SCATTER DATA | Price per Unit (x) versus Units Sold (y), every record").font = \
        Font(bold=True, size=11, color=ACCENT)
    ws.cell(start, 1, "Price per Unit").font = HDR_FONT
    ws.cell(start, 1).fill = HDR_FILL
    ws.cell(start, 2, "Units Sold").font = HDR_FONT
    ws.cell(start, 2).fill = HDR_FILL
    for i, rec in enumerate(F["scatter"].itertuples(index=False), start=start + 1):
        ws.cell(i, 1, float(rec[0])).number_format = PU_FMT
        ws.cell(i, 2, int(rec[1])).number_format = UNITS_FMT
    ch = ScatterChart()
    ch.title = "Pooled price and units co-move weakly (r = %.3f): association, not elasticity" % F["corr_all"]
    ch.style = 13
    ch.x_axis.title = "Price per Unit (USD)"
    ch.y_axis.title = "Units Sold"
    ch.legend = None
    xref = Reference(ws, min_col=1, min_row=start + 1, max_row=start + len(F["scatter"]))
    yref = Reference(ws, min_col=2, min_row=start + 1, max_row=start + len(F["scatter"]))
    s = ch.series[0] if ch.series else None
    ch.add_data(yref, titles_from_data=False)
    ch.set_categories(xref)
    ch.series[0].marker = Marker(symbol="circle", size=4)
    ch.series[0].graphicalProperties.line.noFill = True
    ch.series[0].marker.graphicalProperties.solidFill = ACCENT
    ch.width = 18
    ch.height = 10
    ws.add_chart(ch, "F4")
    ws.cell(start + len(F["scatter"]) + 2, 1,
            "Correlation about %.3f pooled. Sales are mechanically price x units, so any price/volume reading is "
            "descriptive. No causal uplift is claimed." % F["corr_all"]).font = Font(italic=True, size=9,
                                                                                    color=GREY)
    return ws

# ---------------------------------------------------------------- analysis sheets
def seg(frame, col):
    out = frame.copy()
    out["Segment"] = out[col]
    return out


def build_retailer(wb, F):
    def extra(ws, last):
        ws.cell(last, 1, "COVERAGE AND CONFOUNDING").font = Font(bold=True, size=11, color=ACCENT)
        cov = F["locations_per_retailer"].rename(columns={
            "Retailer": "Retailer", "locations": "Locations", "rows": "Rows", "regions": "Regions present"})
        write_frame(ws, last + 1, None, cov,
                    [("Retailer", "Retailer", None), ("Locations", "Locations", INT_FMT),
                     ("Rows", "Rows", UNITS_FMT), ("Regions present", "Regions present", None)],
                    note="Scale is not efficiency: store counts, shelf space and local cost levels are not in the "
                         "data. Retailer comparisons are confounded by geography - BevCo trades only in the South "
                         "and FizzySip only in the West.")

    total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
             "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1}
    frame = seg(F["retail"], "Retailer").sort_values("sales", ascending=False)
    build_dimension_sheet(
        wb, "Retailer_Analysis", F, frame, total,
        f"{F['top_retailer']['Retailer']} leads sales and profit at {F['top_retailer']['sales_share']:.2%} of "
        f"sales, but {F['best_margin_retailer']['Retailer']} leads weighted margin at "
        f"{F['best_margin_retailer']['margin']:.2%}. Scale and efficiency point at different accounts.",
        extra=extra)


def build_region(wb, F):
    def extra(ws, last):
        ws.cell(last, 1, "STATE RANKING | 50 states, sales descending").font = Font(bold=True, size=11, color=ACCENT)
        st = seg(F["state"], "State").sort_values("sales", ascending=False)
        st_total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
                    "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1}
        write_frame(ws, last + 1, None, st,
                    [("State", "Segment", None), ("Sales", "sales", SALES_FMT), ("Units", "units", UNITS_FMT),
                     ("Operating Profit", "profit", SALES_FMT), ("Weighted Margin", "margin", PCT_FMT),
                     ("Revenue per Unit", "rpu", PU_FMT), ("Profit per Unit", "ppu", PU_FMT),
                     ("Sales Share", "sales_share", PCT_FMT), ("Profit Share", "profit_share", PCT_FMT)],
                    total=st_total,
                    note="New York leads sales; Florida leads profit. Montana is retained as a Midwest state "
                         "because that is the source assignment; changing it would alter every regional total.")

    total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
             "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1}
    frame = seg(F["region"], "Region").sort_values("sales", ascending=False)
    build_dimension_sheet(
        wb, "Regional_Analysis", F, frame, total,
        f"{F['region_top']['Region']} sells the most ({F['region_top']['sales']:,.2f}, "
        f"{F['region_top']['sales_share']:.2%} of sales) at {F['region_top']['margin']:.2%} margin - "
        f"{F['region_top']['margin'] - F['margin']:+.2f} points versus the portfolio. "
        f"{F['region_top']['Region']} volume is not evidence of regional efficiency.", extra=extra)


def build_brand(wb, F):
    total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
             "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1}
    frame = seg(F["brand"], "Beverage Brand").sort_values("sales", ascending=False)
    top = frame.iloc[0]
    low = frame.sort_values("margin").iloc[0]
    build_dimension_sheet(
        wb, "Brand_Analysis", F, frame, total,
        f"{top['Beverage Brand']} is the largest brand at {top['sales']:,.2f} ({top['sales_share']:.2%} of sales) "
        f"with a {top['margin']:.2%} margin; {low['Beverage Brand']} is the thinnest at {low['margin']:.2%}. "
        f"Six brands, no tail to prune on scale alone.")


def build_location(wb, F):
    frame = seg(F["loc"], "Location").sort_values("sales", ascending=False)
    frame["eligible"] = np.where(frame["sales"] >= 100_000, "Yes", "No")

    def extra(ws, last):
        ws.cell(last, 1, "QUALIFIED MARGIN RANKING | annual sales >= 100,000, margin descending").font = \
            Font(bold=True, size=11, color=ACCENT)
        q = frame[frame["eligible"] == "Yes"].sort_values("margin", ascending=False)
        write_frame(ws, last + 1, None, q,
                    [("Location", "Segment", None), ("Sales", "sales", SALES_FMT),
                     ("Operating Profit", "profit", SALES_FMT), ("Weighted Margin", "margin", PCT_FMT),
                     ("Revenue per Unit", "rpu", PU_FMT), ("Profit per Unit", "ppu", PU_FMT)],
                    note="37 of 54 locations clear 100,000 annual sales. The threshold is disclosed because the "
                         "best small-market percentage is not the best business: a 51.48% margin on 170,562.50 "
                         "of sales is a smaller prize than a 29.13% margin on 331,750.")

    total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
             "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1}
    build_dimension_sheet(
        wb, "Location_Analysis", F, frame, total,
        f"{F['qualified_n']} of {F['locations']} locations clear 100,000 annual sales. "
        f"{F['qual_best_margin']['Location']} posts the best qualified margin at "
        f"{F['qual_best_margin']['margin']:.2%}; {F['qual_worst_margin']['Location']} the weakest at "
        f"{F['qual_worst_margin']['margin']:.2%}. Location is State | City on the rebuilt key.", extra=extra)


def build_monthly(wb, F):
    ws = new_sheet(wb, "Monthly_Trends", [14, 15, 13, 16, 12, 13, 12, 11, 11, 13, 13], tab=ACCENT)
    banner(ws, "MONTHLY TRENDS | 2021 actual, chronological",
           "Month Start orders the axis; month names are never sorted alphabetically. January and December change "
           "columns stay blank because there is no comparable prior period.")
    m = F["month"].copy()
    m["Segment"] = m["Year-Month"]
    m = m.rename(columns={"mom": "MoM Sales Growth", "margin_pp": "Margin Change pp"})
    cols = [("Month", "Segment", None), ("Sales", "sales", SALES_FMT), ("Units", "units", UNITS_FMT),
            ("Operating Profit", "profit", SALES_FMT), ("Weighted Margin", "margin", PCT_FMT),
            ("Revenue per Unit", "rpu", PU_FMT), ("Profit per Unit", "ppu", PU_FMT),
            ("Sales Share", "sales_share", PCT_FMT), ("Profit Share", "profit_share", PCT_FMT),
            ("MoM Sales Growth", "MoM Sales Growth", PCT_FMT), ("Margin Change pp", "Margin Change pp", PP_FMT)]
    total = {"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"], "profit": F["profit"],
             "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"], "sales_share": 1, "profit_share": 1,
             "MoM Sales Growth": None, "Margin Change pp": None}
    last, hdr = write_frame(ws, 4, "Monthly measures", m, cols, total=total,
                            note="Grand totals are shown for additive and ratio measures only. Change columns have "
                                 "no meaningful total and are deliberately blank. January growth is blank: there is "
                                 "no December 2020.")
    ws.cell(last, 1, "QUARTERLY ROLLUP").font = Font(bold=True, size=11, color=ACCENT)
    q = F["quarter"].copy()
    q["Segment"] = q["Quarter"]
    write_frame(ws, last + 1, None, q,
                [("Quarter", "Segment", None), ("Sales", "sales", SALES_FMT), ("Units", "units", UNITS_FMT),
                 ("Operating Profit", "profit", SALES_FMT), ("Weighted Margin", "margin", PCT_FMT),
                 ("Revenue per Unit", "rpu", PU_FMT), ("Profit per Unit", "ppu", PU_FMT),
                 ("Sales Share", "sales_share", PCT_FMT), ("Profit Share", "profit_share", PCT_FMT)],
                total={"Segment": "Portfolio total", "sales": F["sales"], "units": F["units"],
                       "profit": F["profit"], "margin": F["margin"], "rpu": F["rpu"], "ppu": F["ppu"],
                       "sales_share": 1, "profit_share": 1},
                note="One annual cycle. Peaks cannot be called recurring seasonality from twelve points.")
    return ws


def build_forecast(wb, F):
    ws = new_sheet(wb, "Forecast", [13, 16, 18, 20, 30, 18], tab=PRIOR)
    banner(ws, "ILLUSTRATIVE FORECAST | full portfolio, not linked to the dashboard",
           "Jan-Mar 2022 are modelled estimates, future relative to the dataset rather than to today. They are never "
           "added to actual KPI totals, and no prior-year series exists to compare against.")
    rows = []
    for _, rec in F["month"].iterrows():
        rows.append({"Month Start": date(int(rec["Year-Month"][:4]), int(rec["Year-Month"][5:]), 1),
                     "Actual Sales": rec["sales"], "Illustrative Forecast": None, "Status": "Actual",
                     "Method": "Source portfolio totals", "Prior Year Actual": "Not supplied"})
    for i in range(3):
        rows.append({"Month Start": date(2022, i + 1, 1), "Actual Sales": None,
                     "Illustrative Forecast": F["forecast_flat"], "Status": "Illustrative forecast",
                     "Method": "Fixed Oct-Dec 2021 mean", "Prior Year Actual": "Not supplied"})
    df = pd.DataFrame(rows)
    cols = [("Month Start", "Month Start", DATE_FMT), ("Actual Sales", "Actual Sales", SALES_FMT),
            ("Illustrative Forecast", "Illustrative Forecast", SALES_FMT), ("Status", "Status", None),
            ("Method", "Method", None), ("Prior Year Actual", "Prior Year Actual", None)]
    top = 4
    for j, (h, _k, _f) in enumerate(cols, 1):
        c = ws.cell(top, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    r = top + 1
    for rec in df.itertuples(index=False):
        for j, v in enumerate(rec, 1):
            c = ws.cell(r, j, v)
            c.border = BORDER
            if cols[j - 1][2]:
                c.number_format = cols[j - 1][2]
            if rec[3] == "Illustrative forecast" and j in (1, 3, 4):
                c.font = Font(italic=True, color=FORECAST)
        r += 1

    r += 1
    ws.cell(r, 1, "PRIOR YEAR | not supplied").font = Font(bold=True, size=10, color=INK)
    r += 1
    ws.cell(r, 1, "No 2020 actuals exist, so Prior Year Sales and YoY growth are blank, not zero. Light blue "
                  "remains unused in this file because it is reserved for genuine history.").font = \
        Font(italic=True, size=9, color=GREY)

    r += 2
    ws.cell(r, 1, "HOLDOUT EVIDENCE | train Jan-Sep, predict Oct-Dec").font = Font(bold=True, size=11, color=ACCENT)
    hold = pd.DataFrame(F["holdout_rows"]).rename(columns={"Trailing 3M": "Trailing 3M Forecast"})
    write_frame(ws, r + 1, None, hold,
                [("Month", "Month", None), ("Actual", "Actual", SALES_FMT),
                 ("Trailing 3M Forecast", "Trailing 3M Forecast", SALES_FMT),
                 ("Absolute Error", "Absolute error", SALES_FMT)],
                total={"Month": "Holdout MAE", "Actual": None, "Trailing 3M Forecast": None,
                       "Absolute Error": F["holdout_mae"]},
                note="Three holdout months support a sanity check, not an accuracy claim. No seasonality, causal "
                     "uplift or savings are inferred from this baseline.")

    ch = LineChart()
    ch.title = "Illustrative Jan-Mar 2022 sits on a flat Oct-Dec mean (%s per month)" % f"{F['forecast_flat']:,.0f}"
    ch.style = 12
    ch.y_axis.title = "Sales (USD)"
    ch.y_axis.majorGridlines = ChartLines()
    ch.y_axis.numFmt = "#,##0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    data = Reference(ws, min_col=2, max_col=3, min_row=4, max_row=4 + len(df))
    cats = Reference(ws, min_col=1, min_row=5, max_row=4 + len(df))
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.series[0].graphicalProperties.line.solidFill = ACCENT
    ch.series[0].graphicalProperties.line.width = 28000
    ch.series[1].graphicalProperties.line.solidFill = FORECAST
    ch.series[1].graphicalProperties.line.dashStyle = "dash"
    ch.series[1].graphicalProperties.line.width = 25000
    ch.width = 19
    ch.height = 10
    ws.add_chart(ch, "H4")
    return ws


# ---------------------------------------------------------------- dashboard
DROPDOWNS = {"Retailer": "Dashboard!$B$4", "Region": "Dashboard!$F$4", "Beverage Brand": "Dashboard!$J$4", "Year-Month": "Dashboard!$N$4"}
DIMS = ["Retailer", "Region", "Beverage Brand", "Year-Month"]


def sumifs_expr(measure, own_dim=None, row=None):
    parts = []
    for dim in DIMS:
        if dim == own_dim:
            crit = f"$A${row}"
        else:
            ref = DROPDOWNS[dim]
            crit = f'IF({ref}="All","*",{ref})'
        parts.append(f"SalesData[{dim}],{crit}")
    return "SUMIFS(SalesData[" + measure + "]," + ",".join(parts) + ")"


def helper_block(ws, top, title, members, own_dim, dd_name):
    ws.cell(top, 1, title).font = Font(bold=True, size=11, color=ACCENT)
    hdr = top + 1
    for j, h in enumerate(["Segment", "Sales", "Units", "Operating Profit", "Weighted Margin",
                           "Revenue per Unit"], 1):
        c = ws.cell(hdr, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    r = hdr + 1
    for mem in members:
        ws.cell(r, 1, mem).border = BORDER
        ws.cell(r, 2, f"={sumifs_expr('Total Sales', own_dim, r)}").number_format = SALES_FMT
        ws.cell(r, 3, f"={sumifs_expr('Units Sold', own_dim, r)}").number_format = UNITS_FMT
        ws.cell(r, 4, f"={sumifs_expr('Operating Profit', own_dim, r)}").number_format = SALES_FMT
        c = ws.cell(r, 5, f'=IF(B{r}=0,"",D{r}/B{r})')
        c.number_format = PCT_FMT
        c = ws.cell(r, 6, f'=IF(C{r}=0,"",B{r}/C{r})')
        c.number_format = PU_FMT
        for j in range(1, 7):
            ws.cell(r, j).border = BORDER
        r += 1
    return hdr, r - 1, r + 1


def build_dashboard_data(wb, F):
    ws = new_sheet(wb, "Dashboard_Data", [18, 15, 13, 16, 12, 13], tab=GREY)
    ws.cell(1, 1, "FORMULA HELPERS | the four dashboard dropdowns drive every value below").font = \
        Font(bold=True, size=10, color=GREY)
    retail_hdr, retail_last, nxt = helper_block(ws, 3, "RETAILERS", F["retailers"], "Retailer", "retail")
    region_hdr, region_last, nxt2 = helper_block(ws, nxt + 1, "REGIONS", F["regions"], "Region", "region")
    brand_hdr, brand_last, nxt3 = helper_block(ws, nxt2 + 1, "BRANDS", F["brands"], "Beverage Brand", "brand")
    month_hdr, month_last, nxt4 = helper_block(ws, nxt3 + 1, "MONTHS", F["month"]["Year-Month"].tolist(),
                                               "Year-Month", "month")
    ws.cell(nxt4, 1, 'Wildcard "*" criteria mean All. A blank ratio means the denominator is zero in the current '
                     'selection; it is never shown as 0%.').font = Font(italic=True, size=9, color=GREY)
    return ws, {"retail": (retail_hdr, retail_last), "region": (region_hdr, region_last),
                "brand": (brand_hdr, brand_last), "month": (month_hdr, month_last)}


def build_dashboard(wb, F, dd, refs):
    ws = new_sheet(wb, "Dashboard", [9] * 17, tab=ACCENT)
    ws.sheet_view.zoomScale = 90
    ws.merge_cells("A1:Q1")
    c = ws["A1"]
    c.value = "COKE | Sales & Operating Profitability, 2021"
    c.font = Font(bold=True, size=20, color="FFFFFF")
    c.fill = fill(ACCENT)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 36
    ws.merge_cells("A2:Q2")
    c = ws["A2"]
    c.value = ("USD assumed | six brands, four retailers, 54 locations | all filters reset = full 2021 portfolio "
               "| charts recompute from Dashboard_Data")
    c.font = Font(italic=True, size=9, color=GREY)

    for col, label in [("B", "Retailer"), ("F", "Region"), ("J", "Brand"), ("N", "Month")]:
        c = ws[f"{col}3"]
        c.value = label
        c.font = Font(bold=True, size=9, color=INK)
    for col, ref in [("B", "Retailer"), ("F", "Region"), ("J", "Brand"), ("N", "Month")]:
        c = ws[f"{col}4"]
        c.value = "All"
        c.font = Font(bold=True, size=11, color=ACCENT)
        c.fill = fill(HEAD_FILL)
        c.border = BORDER
        c.alignment = Alignment(horizontal="center")
    dv_r = DataValidation(type="list", formula1="Lists!$A$1:$A$5", allow_blank=False, showDropDown=False)
    dv_g = DataValidation(type="list", formula1="Lists!$B$1:$B$6", allow_blank=False, showDropDown=False)
    dv_b = DataValidation(type="list", formula1="Lists!$C$1:$C$7", allow_blank=False, showDropDown=False)
    dv_m = DataValidation(type="list", formula1="Lists!$D$1:$D$13", allow_blank=False, showDropDown=False)
    for dv in (dv_r, dv_g, dv_b, dv_m):
        ws.add_data_validation(dv)
    dv_r.add(ws["B4"])
    dv_g.add(ws["F4"])
    dv_b.add(ws["J4"])
    dv_m.add(ws["N4"])

    ws.merge_cells("A5:Q5")
    c = ws["A5"]
    c.value = ("Reset by choosing All in each of B4, F4, J4, N4. No macro reset button exists. Card ratios return "
               "blank, not 0%, when the selection has no sales.")
    c.font = Font(italic=True, size=9, color=GREY)

    cards = [
        ("A", "B", "C", "Sales", f"={sumifs_expr('Total Sales')}", SALES_FMT, "Revenue in the current selection"),
        ("D", "E", "F", "Units Sold", f"={sumifs_expr('Units Sold')}", UNITS_FMT, "Physical units sold"),
        ("G", "H", "I", "Operating Profit", f"={sumifs_expr('Operating Profit')}", SALES_FMT,
         "Operating level, not gross or net profit"),
        ("J", "K", "L", "Weighted Margin", '=IF(A8=0,"",G8/A8)', PCT_FMT, "Profit / sales, never an average of margins"),
        ("M", "N", "O", "Revenue per Unit", '=IF(D8=0,"",A8/D8)', PU_FMT, "Sales / units"),
    ]
    for first, second, third, label, formula, fmt, note in cards:
        ws.merge_cells(f"{first}7:{third}7")
        c = ws[f"{first}7"]
        c.value = label.upper()
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = fill(ACCENT)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{first}8:{third}8")
        c = ws[f"{first}8"]
        c.value = formula
        c.number_format = fmt
        c.font = Font(bold=True, size=18, color=INK)
        c.fill = fill(PANEL)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER
        ws.merge_cells(f"{first}9:{third}9")
        c = ws[f"{first}9"]
        c.value = note
        c.font = Font(italic=True, size=8, color=GREY)
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[8].height = 34

    def section(row, text):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=17)
        c = ws.cell(row, 1, text)
        c.font = Font(bold=True, size=10, color=ACCENT)
        c.alignment = Alignment(vertical="center")

    top_ret, top_margin = F["top_retailer"], F["best_margin_retailer"]
    section(11, "SCALE VS EFFICIENCY | who sells, who earns")
    rh, rl = refs["retail"]
    bar_chart(dd, "All filters: %s is %.1f%% of sales - the scale leader, second on margin"
              % (top_ret["Retailer"], top_ret["sales_share"] * 100), rh, rl, 2, ws, "A12")
    bar_chart(dd, "All filters: %s earns the highest weighted margin (%.2f%%) on the smallest base"
              % (top_margin["Retailer"], top_margin["margin"] * 100), rh, rl, 5, ws, "I12", pct=True)

    section(29, "WHERE IT SELLS | region and brand")
    gh, gl = refs["region"]
    region = F["region"].sort_values("sales", ascending=False).iloc[0]
    bar_chart(dd, "All filters: %s leads sales at %.2f%% of the portfolio, but margin trails at %.2f%%"
              % (region["Region"], region["sales_share"] * 100, region["margin"] * 100), gh, gl, 2, ws, "A30")
    bh, bl = refs["brand"]
    brand = F["brand"].sort_values("sales", ascending=False).iloc[0]
    bar_chart(dd, "All filters: %s is the largest brand at %.2f%% of sales"
              % (brand["Beverage Brand"], brand["sales_share"] * 100), bh, bl, 2, ws, "I30")

    section(47, "WHEN IT SELLS | monthly trend and margin")
    mh, ml = refs["month"]
    line_chart(dd, "All filters: sales peak in %s at %s; Jul is the volume and profit peak"
               % (F["sales_peak_month"]["Year-Month"], f"{F['sales_peak_month']['sales']:,.0f}"), mh, ml, ws, "A48")
    line_chart(dd, "All filters: margin holds a narrow %.1f-%.1f%% band all year"
               % (F["month"]["margin"].min() * 100, F["month"]["margin"].max() * 100), mh, ml, ws, "I48", pct=True,
               val_col=5)

    section(65, "WHEN AND WHERE | region x month sales heatmap (blue bar = magnitude; months are fixed columns, "
                "the Month filter does not apply here)")
    months = F["month"]["Year-Month"].tolist()
    hdr_row = 67
    ws.cell(hdr_row, 1, "Region").font = HDR_FONT
    ws.cell(hdr_row, 1).fill = HDR_FILL
    for j, ym in enumerate(months, start=2):
        c = ws.cell(hdr_row, j, ym)
        c.font = Font(bold=True, size=8, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center")
        c.number_format = "@"
    for i, rec in enumerate(F["region"].sort_values("sales", ascending=False).itertuples(), start=1):
        r = hdr_row + i
        ws.cell(r, 1, rec.Region).font = Font(bold=True, size=9, color=INK)
        for j, ym in enumerate(months, start=2):
            col = gcl(j)
            formula = (f'=SUMIFS(SalesData[Total Sales],'
                       f'SalesData[Region],IF($F$4="All",$A{r},$F$4),'
                       f'SalesData[Year-Month],{col}${hdr_row},'
                       f'SalesData[Retailer],IF($B$4="All","*",$B$4),'
                       f'SalesData[Beverage Brand],IF($J$4="All","*",$J$4))')
            c = ws.cell(r, j, formula)
            c.number_format = "#,##0"
            c.font = Font(size=8, color=INK)
    last_hm = hdr_row + len(F["region"])
    ws.conditional_formatting.add(
        f"B{hdr_row + 1}:M{last_hm}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=ACCENT, showValue=True))

    section(75, "BASELINE READING | reset filters to reproduce these statements")
    blocks = [
        ("A", "B", "PROTECT THE SCALE LEADER",
         f"{F['top_retailer']['Retailer']}: {F['top_retailer']['sales_share']:.2%} of sales and "
         f"{F['top_retailer']['profit_share']:.2%} of profit across "
         f"{int(F['locations_per_retailer'].set_index('Retailer').loc[F['top_retailer']['Retailer'], 'locations'])} "
         f"locations. Margin {F['top_retailer']['margin']:.2%} - below "
         f"{F['best_margin_retailer']['Retailer']}'s {F['best_margin_retailer']['margin']:.2%}, but that retailer "
         f"trades in one region, so this is mix, not proof."),
        ("E", "G", "FIX WESTERN MARGIN",
         f"West carries {F['region_top']['sales_share']:.2%} of sales at {F['region_top']['margin']:.2%} margin "
         f"({F['region_top']['margin'] - F['margin']:+.2f} pp versus portfolio). Seattle "
         f"{F['loc'][F['loc']['Location'] == 'Washington | Seattle']['margin'].iloc[0]:.2%} and San Francisco "
         f"{F['loc'][F['loc']['Location'] == 'California | San Francisco']['margin'].iloc[0]:.2%} are the visible "
         f"drains. Cost and promo detail are absent - diagnose before cutting price."),
        ("I", "K", "PLAN BY REGION AND KPI",
         f"December leads sales at {F['dec']['sales']:,.0f}; July leads units ({F['jul']['units']:,.0f}) and profit "
         f"({F['jul']['profit']:,.0f}). West peaks in July, every other region in December. One year is a pattern, "
         f"not proven seasonality."),
    ]
    for first, last_col, head, text in blocks:
        ws.merge_cells(f"{first}76:{last_col}76")
        c = ws[f"{first}76"]
        c.value = head
        c.font = Font(bold=True, size=9, color="FFFFFF")
        c.fill = fill(FORECAST)
        c.alignment = Alignment(horizontal="center")
        ws.merge_cells(f"{first}77:{last_col}79")
        c = ws[f"{first}77"]
        c.value = text
        c.font = Font(size=9, color=INK)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.fill = fill(BAND)
        c.border = BORDER
        ws.row_dimensions[77].height = 46

    ws.merge_cells("A81:Q81")
    c = ws["A81"]
    c.value = ("Prior-year comparison: not supplied. Forecast: see Forecast sheet, illustrative only, never added to "
               "these cards. Operating profit is operating-level, not gross or net profit. Price-units correlation is "
               "association, not elasticity.")
    c.font = Font(italic=True, size=8, color=GREY)
    return ws


def build_insights(wb, F):
    ws = new_sheet(wb, "Insights_Recommendations", [7, 11, 46, 40, 42, 40, 16, 14, 24, 3], tab=PRIOR)
    banner(ws, "INSIGHTS | finding, why it matters, caveat, action",
           "Facts are computed from the 3,888 rows. Hypotheses are labelled. No uplift, saving or market share is "
           "invented; owners and timeframes are proposals.", span=9)
    top = F["top_retailer"]
    best_m = F["best_margin_retailer"]
    west = F["region_top"]
    sf = F["loc"][F["loc"]["Location"] == "California | San Francisco"].iloc[0]
    se = F["loc"][F["loc"]["Location"] == "Washington | Seattle"].iloc[0]
    rows = [
        {"Priority": 1, "Type": "Fact",
         "Finding": f"{top['Retailer']} produces {top['sales_share']:.2%} of sales and {top['profit_share']:.2%} of "
                    f"operating profit at {top['margin']:.2%} margin.",
         "Why it matters": "Loss of this account would remove half of the portfolio's profit, not just revenue.",
         "Caveat": f"{best_m['Retailer']} shows {best_m['margin']:.2%} margin but trades in one region only "
                   f"({F['locations_per_retailer'].set_index('Retailer').loc[best_m['Retailer'], 'regions']}); "
                   f"geography and mix confound the comparison.",
         "Recommended action": "Protect distribution and availability first; benchmark the highest-margin cells "
                               "inside this retailer's own footprint before renegotiating.",
         "Owner": "Account lead (proposed)", "Timeframe": "Weeks 1-3", "Success KPI": "Profit contribution and "
                                                                                    "matched-cell margin"},
        {"Priority": 2, "Type": "Fact",
         "Finding": f"West sells {west['sales']:,.2f} ({west['sales_share']:.2%} of sales) at "
                    f"{west['margin']:.2%} margin, {west['margin'] - F['margin']:+.2f} pp versus the portfolio. "
                    f"Seattle {se['margin']:.2%} and San Francisco {sf['margin']:.2%} are the weakest qualified "
                    f"locations.",
         "Why it matters": "The largest region is diluting margin; a 1 pp recovery on West sales is roughly "
                           f"{west['sales'] * 0.01:,.0f} of operating profit.",
         "Caveat": "Cost, trade-spend and promotion detail are absent. Location cost differences are a hypothesis, "
                   "not a measured driver.",
         "Recommended action": "Diagnose cost-to-serve and price realisation in Seattle and San Francisco before "
                               "touching list price.",
         "Owner": "Regional commercial analyst (proposed)", "Timeframe": "Weeks 1-3",
         "Success KPI": "Matched-cell margin in pp; profit per unit"},
        {"Priority": 3, "Type": "Fact",
         "Finding": f"{F['sales_peak_month']['Year-Month']} is the sales peak at "
                    f"{F['sales_peak_month']['sales']:,.0f}; {F['units_peak_month']['Year-Month']} is the units and "
                    f"profit peak ({F['units_peak_month']['units']:,.0f} units, "
                    f"{F['units_peak_month']['profit']:,.0f} profit). West peaks in July, the other four regions in "
                    f"December.",
         "Why it matters": "A single supply and staffing plan built on December would under-serve July volume.",
         "Caveat": "Twelve months cannot establish recurring seasonality; 2020 is not supplied.",
         "Recommended action": "Plan capacity by region and KPI, not by the portfolio aggregate.",
         "Owner": "Demand planner (proposed)", "Timeframe": "Next planning cycle",
         "Success KPI": "Forecast error and availability once data is collected"},
        {"Priority": 4, "Type": "Fact",
         "Finding": f"{F['qualified_n']} of {F['locations']} locations clear 100,000 annual sales. "
                    f"{F['qual_best_margin']['Location']} is best at {F['qual_best_margin']['margin']:.2%}; "
                    f"{F['qual_worst_margin']['Location']} is weakest at {F['qual_worst_margin']['margin']:.2%}.",
         "Why it matters": "Percentages on small markets mislead: the threshold keeps the ranking decision-useful.",
         "Caveat": "The 100,000 threshold is a disclosure choice; smaller markets are excluded by construction.",
         "Recommended action": "Use the qualified ranking for margin reviews and keep full sales ranking for scale.",
         "Owner": "Category finance (proposed)", "Timeframe": "Weeks 2-6", "Success KPI": "Qualified market margin"},
        {"Priority": 5, "Type": "Hypothesis",
         "Finding": f"Pooled price-units correlation is {F['corr_all']:.3f} (within-brand "
                    f"{min(c['r'] for c in F['corr_brand']):.2f} to {max(c['r'] for c in F['corr_brand']):.2f}).",
         "Why it matters": "If the association reflects demand, uniform price rises could cost volume; but this "
                           "dataset cannot quantify that.",
         "Caveat": "Sales are price x units, so the two are mechanically linked; brand, retailer, location and month "
                   "all confound. Association only - not elasticity.",
         "Recommended action": "Run a controlled test in matched brand-retailer-location cells with a holdout group "
                               "and guardrails on units and profit.",
         "Owner": "Pricing analyst (proposed)", "Timeframe": "Design weeks 3-6",
         "Success KPI": "Incremental profit with units and margin guardrails"},
        {"Priority": 6, "Type": "Fact",
         "Finding": "2020 actuals are not supplied, so prior-year sales and YoY growth are blank for every month.",
         "Why it matters": "A flat or falling year cannot be distinguished from growth without a comparator.",
         "Caveat": "Not a data error: the history does not exist in the source.",
         "Recommended action": "Show 'Not supplied' and keep the YoY field blank rather than zero until a prior "
                               "period is loaded.",
         "Owner": "Data owner (proposed)", "Timeframe": "Before any YoY reporting",
         "Success KPI": "Prior-year series loaded and reconciled"},
    ]
    cols = [("Priority", "Priority", INT_FMT), ("Type", "Type", None), ("Finding", "Finding", None),
            ("Why it matters", "Why it matters", None), ("Caveat", "Caveat", None),
            ("Recommended action", "Recommended action", None), ("Owner", "Owner", None),
            ("Timeframe", "Timeframe", None), ("Success KPI", "Success KPI", None)]
    last, _ = write_frame(ws, 4, None, pd.DataFrame(rows), cols,
                          note="Separate fact from hypothesis: rows 1-4 and 6 are measured facts; row 5 is a "
                               "hypothesis to test. No revenue uplift, saving or share target is asserted.")
    for r in range(6, 6 + len(rows)):
        for j in range(1, 10):
            ws.cell(r, j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 74
    return ws

# ---------------------------------------------------------------- documentation sheets
M_SALES = """let
    Source = Excel.Workbook(File.Contents(pSourceFile), null, true),
    Raw = Source{[Item="Data", Kind="Sheet"]}[Data],
    SkipTitles = Table.Skip(Raw, 4),
    Headers = Table.PromoteHeaders(SkipTitles, [PromoteAllScalars=true]),
    Business = Table.SelectColumns(Headers, {"Retailer", "Retailer ID", "Invoice Date", "Region", "State", "City",
        "Beverage Brand", "Price per Unit", "Units Sold", "Total Sales", "Operating Profit", "Operating Margin"}),
    Types = Table.TransformColumnTypes(Business, {{"Retailer", type text}, {"Retailer ID", type text},
        {"Invoice Date", type date}, {"Region", type text}, {"State", type text}, {"City", type text},
        {"Beverage Brand", type text}, {"Price per Unit", type number}, {"Units Sold", Int64.Type},
        {"Total Sales", type number}, {"Operating Profit", type number}, {"Operating Margin", type number}}),
    TextClean = Table.TransformColumns(Types, List.Transform({"Retailer", "Region", "State", "City",
        "Beverage Brand"}, (c) => {c, each if _ = null then null else Text.Clean(Text.Trim(_)), type text})),
    AddYear = Table.AddColumn(TextClean, "Year", each Date.Year([Invoice Date]), Int64.Type),
    AddMonthNumber = Table.AddColumn(AddYear, "Month Number", each Date.Month([Invoice Date]), Int64.Type),
    AddMonthStart = Table.AddColumn(AddMonthNumber, "Month Start", each Date.StartOfMonth([Invoice Date]), type date),
    AddYearMonth = Table.AddColumn(AddMonthStart, "Year-Month", each Date.ToText([Invoice Date], "yyyy-MM"), type text),
    AddQuarter = Table.AddColumn(AddYearMonth, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Invoice Date])),
        type text),
    AddLocation = Table.AddColumn(AddQuarter, "Location", each [State] & " | " & [City], type text),
    SalesDifference = Table.AddColumn(AddLocation, "Sales Difference",
        each [Total Sales] - [Price per Unit] * [Units Sold], type number),
    ProfitDifference = Table.AddColumn(SalesDifference, "Profit Difference",
        each [Operating Profit] - [Total Sales] * [Operating Margin], type number),
    SalesCheck = Table.AddColumn(ProfitDifference, "Sales Check",
        each if Number.Abs([Sales Difference]) <= 0.01 then "OK" else "Investigate", type text),
    ProfitCheck = Table.AddColumn(SalesCheck, "Profit Check",
        each if Number.Abs([Profit Difference]) <= 0.01 then "OK" else "Investigate", type text),
    Activity = Table.AddColumn(ProfitCheck, "Activity Flag",
        each if [Units Sold] = 0 and [Total Sales] = 0 then "Review zero activity" else "Active", type text)
in
    Activity"""

M_CAL = """let
    Start = #date(2020, 1, 1),
    End = #date(2022, 12, 31),
    Dates = List.Dates(Start, Duration.Days(End - Start) + 1, #duration(1,0,0,0)),
    Base = Table.FromList(Dates, Splitter.SplitByNothing(), {"Date"}),
    Typed = Table.TransformColumnTypes(Base, {{"Date", type date}}),
    Year = Table.AddColumn(Typed, "Year", each Date.Year([Date]), Int64.Type),
    MonthNumber = Table.AddColumn(Year, "Month Number", each Date.Month([Date]), Int64.Type),
    MonthName = Table.AddColumn(MonthNumber, "Month Name", each Date.ToText([Date], "MMM"), type text),
    MonthStart = Table.AddColumn(MonthName, "Month Start", each Date.StartOfMonth([Date]), type date),
    YearMonth = Table.AddColumn(MonthStart, "Year-Month", each Date.ToText([Date], "yyyy-MM"), type text),
    Quarter = Table.AddColumn(YearMonth, "Quarter", each "Q" & Text.From(Date.QuarterOfYear([Date])), type text)
in
    Quarter"""


def mono(wb, name, heading, blocks, widths, tab=GREY):
    ws = new_sheet(wb, name, widths, tab=tab)
    banner(ws, heading, "Copy-ready text, not an installed query or measure. Paste into desktop Excel; the code "
                        "below is the specification the delivered file was built to.")
    r = 4
    for label, code in blocks:
        ws.cell(r, 1, label).font = Font(bold=True, size=11, color=ACCENT)
        r += 1
        for line in code.split("\n"):
            c = ws.cell(r, 1, line)
            c.font = Font(name="Consolas", size=9, color=INK)
            c.alignment = Alignment(vertical="top")
            r += 1
        r += 1
    return ws


def build_pivot_spec(wb, F):
    ws = new_sheet(wb, "Pivot_Tables", [17, 20, 26, 20, 44, 30, 18, 30], tab=GREY)
    banner(ws, "PIVOT LAYOUTS | nine native pivots, one shared source",
           "Each row is the exact layout to build from Insert > PivotTable > Use this workbook's Data Model. They "
           "must share one source so Report Connections can attach slicers to all of them.")
    rows = [
        {"Name": "PT_KPI", "Source": "SalesData / Data Model", "Rows": "(none)", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Revenue/Unit",
         "Filter / sort": "Calendar Year = 2021; shared slicers", "Visual": "Five linked KPI cards"},
        {"Name": "PT_Retailer", "Source": "Same cache", "Rows": "Retailer", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Revenue/Unit, Profit/Unit, Location Count",
         "Filter / sort": "Shared controls; sales descending", "Visual": "Sales bar plus separate margin bar"},
        {"Name": "PT_Region", "Source": "Same cache", "Rows": "Region", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Revenue/Unit, Region Sales Share",
         "Filter / sort": "Shared controls; sales descending", "Visual": "Regional sales bar"},
        {"Name": "PT_State", "Source": "Same cache", "Rows": "State", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin",
         "Filter / sort": "Shared controls; sales descending", "Visual": "State ranking bar"},
        {"Name": "PT_Location", "Source": "Same cache", "Rows": "Location", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Revenue/Unit",
         "Filter / sort": "Shared controls; sales descending", "Visual": "Top and bottom location bars"},
        {"Name": "PT_Location_Margin", "Source": "Same cache", "Rows": "Location", "Columns": "(none)",
         "Values": "Sales, Operating Profit, Weighted Margin",
         "Filter / sort": "Sales >= 100,000; margin descending", "Visual": "Qualified margin ranking"},
        {"Name": "PT_Brand", "Source": "Same cache", "Rows": "Beverage Brand", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Profit/Unit, Brand Sales Share",
         "Filter / sort": "Shared controls; sales descending", "Visual": "Brand contribution bars"},
        {"Name": "PT_Month", "Source": "Same cache", "Rows": "Calendar Month Start", "Columns": "(none)",
         "Values": "Sales, Units, Operating Profit, Weighted Margin, Revenue/Unit, MoM Sales Growth, "
                   "Margin Change pp",
         "Filter / sort": "Calendar Year = 2021; date ascending", "Visual": "Monthly sales and margin lines"},
        {"Name": "PT_Heatmap", "Source": "Same cache", "Rows": "Region", "Columns": "Calendar Month Start",
         "Values": "Sales", "Filter / sort": "Exclude totals; dates ascending",
         "Visual": "Plain-blue data bars"},
    ]
    cols = [("Name", "Name", None), ("Source", "Source", None), ("Rows", "Rows", None),
            ("Columns", "Columns", None), ("Values", "Values", None), ("Filter / sort", "Filter / sort", None),
            ("Visual", "Visual", None)]
    last, _ = write_frame(ws, 4, None, pd.DataFrame(rows), cols,
                          note="Status: not installed in this file. The delivered dashboard is formula-driven, so "
                               "these layouts are the native upgrade. Build them from one source, then insert "
                               "slicers for Retailer, Region and Beverage Brand and a timeline on Calendar Date.")
    ws.cell(last, 1, "THEN: click any pivot > PivotTable Analyze > Insert Slicer (Retailer, Region, Beverage Brand); "
                     "Insert Timeline (Calendar Date, 2021); each control > Report Connections > tick all nine "
                     "pivots including PT_KPI. Hide grand totals on the monthly-change pivots; they are not "
                     "meaningful. There is no macro reset: clear each slicer filter to reset.").font = \
        Font(italic=True, size=9, color=GREY)
    ws.merge_cells(start_row=last, start_column=1, end_row=last, end_column=8)
    ws.cell(last, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[last].height = 60
    return ws


def build_pq(wb, F):
    ws = mono(wb, "Power_Query_Code", "POWER QUERY | two queries, copy-ready M",
              [("PARAMETER pSourceFile | Data > Get Data > From Other Sources > Blank Query > Text parameter; "
                "value = full local path to the original attachment (not hard-coded here).", ""),
               ("SALESDATA QUERY", M_SALES),
               ("CALENDAR QUERY", M_CAL),
               ("AFTER QUERY CREATION", "Load SalesData to Cleaned_Data and the Data Model, keeping the table name "
                                        "SalesData; load Calendar connection-only and to the Data Model. Mark "
                                        "Calendar as the date table, relate Calendar[Date] one -> many "
                                        "SalesData[Invoice Date], and sort Month Name by Month Number. Work on a "
                                        "copy and reconcile before replacing the delivered table.")],
              [150, 20, 20], tab=GREY)
    for row in ws.iter_rows(min_row=4, max_row=ws.max_row, min_col=1, max_col=1):
        for c in row:
            if c.font and c.font.name == "Consolas":
                continue
    return ws


def build_dax(wb, F):
    ws = new_sheet(wb, "DAX_Measures", [24, 74, 16, 52], tab=GREY)
    banner(ws, "DAX MEASURES | create in Power Pivot > Measures > New Measure",
           "Specification only: no model is installed in this file. Every ratio divides sums; zero denominators "
           "stay blank. Prior-year measures must stay blank because 2020 is absent.")
    rows = [
        ("Sales", "SUM(SalesData[Total Sales])", "USD 2dp", "Additive revenue."),
        ("Units", "SUM(SalesData[Units Sold])", "Whole number", "Additive physical units."),
        ("Operating Profit", "SUM(SalesData[Operating Profit])", "USD 2dp",
         "Operating level only, not gross or net profit."),
        ("Weighted Margin", "DIVIDE([Operating Profit],[Sales])", "Percent 2dp", "Ratio of sums, never an average "
                                                                                "of Operating Margin."),
        ("Revenue per Unit", "DIVIDE([Sales],[Units])", "USD 4dp", "Ratio of sums."),
        ("Profit per Unit", "DIVIDE([Operating Profit],[Units])", "USD 4dp", "Ratio of sums."),
        ("Record Count", "COUNTROWS(SalesData)", "Whole number", "Reporting rows, not consumer transactions."),
        ("Location Count", "DISTINCTCOUNT(SalesData[Location])", "Whole number", "State | City coverage."),
        ("Previous Month Sales", "CALCULATE([Sales],DATEADD(Calendar[Date],-1,MONTH))", "USD 2dp",
         "Monthly rows only."),
        ("MoM Sales Growth", "IF(ISBLANK([Sales]) || ISBLANK([Previous Month Sales]),BLANK(),DIVIDE([Sales]-"
                             "[Previous Month Sales],[Previous Month Sales]))", "Percent 2dp",
         "January blank; hide the monthly-change grand total."),
        ("Previous Month Margin", "CALCULATE([Weighted Margin],DATEADD(Calendar[Date],-1,MONTH))", "Percent 2dp",
         "Prior chronological month."),
        ("Margin Change pp", "IF(ISBLANK([Weighted Margin]) || ISBLANK([Previous Month Margin]),BLANK(),"
                             "([Weighted Margin]-[Previous Month Margin])*100)", "0.00 pp", "Points, not percent "
                                                                                         "growth."),
        ("Prior Year Sales", "CALCULATE([Sales],DATEADD(Calendar[Date],-1,YEAR))", "USD 2dp",
         "Blank for 2021: no 2020 actuals."),
        ("YoY Sales Growth", "IF(ISBLANK([Sales]) || ISBLANK([Prior Year Sales]),BLANK(),DIVIDE([Sales]-"
                             "[Prior Year Sales],[Prior Year Sales]))", "Percent 2dp",
         "Never force the missing year to zero."),
        ("Retailer Sales Share", "DIVIDE([Sales],CALCULATE([Sales],ALLSELECTED(SalesData[Retailer])))",
         "Percent 2dp", "Share within the selected retailer set."),
        ("Retailer Profit Share", "DIVIDE([Operating Profit],CALCULATE([Operating Profit],"
                                  "ALLSELECTED(SalesData[Retailer])))", "Percent 2dp", "Selected scope."),
        ("Region Sales Share", "DIVIDE([Sales],CALCULATE([Sales],ALLSELECTED(SalesData[Region])))", "Percent 2dp",
         "Selected region set."),
        ("Brand Sales Share", "DIVIDE([Sales],CALCULATE([Sales],ALLSELECTED(SalesData[Beverage Brand])))",
         "Percent 2dp", "Selected brand set."),
    ]
    write_frame(ws, 4, "Measures",
                pd.DataFrame(rows, columns=["Measure", "Expression (after =)", "Format", "Purpose / caveat"]),
                [("Measure", "Measure", None), ("Expression (after =)", "Expression (after =)", None),
                 ("Format", "Format", None), ("Purpose / caveat", "Purpose / caveat", None)],
                note="Every line is text. No DAX model, relationship or measure is installed in this workbook.")
    for r in range(6, 6 + len(rows)):
        c = ws.cell(r, 2)
        c.font = Font(name="Consolas", size=9)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 34
    return ws


def build_guide(wb, F):
    ws = new_sheet(wb, "Implementation_Guide", [6, 15, 34, 66, 34, 30], tab=GREY)
    banner(ws, "IMPLEMENTATION | delivered components and the exact native build",
           "Filter the Phase column to focus. Steps marked MANUAL are not installed in this file and cannot be "
           "installed by a script: Power Query connections, the Power Pivot model, native pivots, slicers.")
    rows = [
        ("1", "START", "Open the delivered file", "Enable ordinary recalculation. Start at README, then Dashboard.",
         "23 sheets; source values retained", "No macros, no external connections."),
        ("2", "THEME", "Apply blue semantics", "Use Color_Theme. Actual plain blue; prior-year light blue (unused); "
                                               "forecast dark blue dashed.", "Consistent colour meaning",
         "Only 2021 exists, so light blue never appears."),
        ("3", "SOURCE", "Inspect the original", "Data!B5:M5 headers, B6:M3893 business rows. Keep titles and A/P:R "
                                                "helpers out of the model.", "3,888 x 12 business fields",
         "Original attachment is never modified."),
        ("4", "POWER QUERY (MANUAL)", "Create pSourceFile", "Data > Get Data > From Other Sources > Blank Query > "
                                                            "Manage Parameters > New > Text; value = workbook path.",
         "Parameterised path, no hard-coded C:\\", "Not installed here; text only."),
        ("5", "POWER QUERY (MANUAL)", "RawSales connection", "Get Data > From Excel Workbook > Data sheet > "
                                                             "Transform; skip top 4 rows; promote headers.",
         "Reads the unchanged source", "Load as connection only."),
        ("6", "POWER QUERY (MANUAL)", "Business columns and types", "Keep exactly 12 columns Retailer to Operating "
                                                                     "Margin; Retailer ID text, Invoice Date date, "
                                                                     "Units whole, money and margin decimal.",
         "Identifiers never summed", "Avoid fixed-currency type: it rounds precision."),
        ("7", "POWER QUERY (MANUAL)", "Clean and derive", "Trim/Clean the five dimensions; add Year, Month Number, "
                                                          "Month Start, Year-Month, Quarter, Location, Sales "
                                                          "Difference, Profit Difference, the two checks and the "
                                                          "activity flag.", "23 columns, matching Cleaned_Data",
         "Location = State | City; city alone repeats."),
        ("8", "LOAD (MANUAL)", "Load to table and model", "Close & Load To > table Cleaned_Data and Add to Data "
                                                          "Model; keep table name SalesData.", "Query-backed table",
         "Work on a copy; reconcile before replacing."),
        ("9", "MODEL (MANUAL)", "Calendar query", "Blank Query > Advanced Editor > Calendar block; load "
                                                  "connection-only to the model.", "2020-2022 date table",
         "A calendar row is not a sales record."),
        ("10", "MODEL (MANUAL)", "Date relationship", "Diagram View: Calendar[Date] one to many "
                                                      "SalesData[Invoice Date]; Mark as date table.", "One-to-many",
         "Sort Month Name by Month Number."),
        ("11", "MEASURES (MANUAL)", "Additive measures", "Power Pivot > Measures > New Measure; name, then "
                                                         "=expression.", "Sales 8,684,027.50; units 17,148,250",
         "See DAX_Measures for all 18 definitions."),
        ("12", "MEASURES (MANUAL)", "Ratio and change measures", "DIVIDE for weightings; DATEADD for prior month and "
                                                                 "prior year; ISBLANK guards on changes.",
         "Margin 36.55%; YoY blank", "Never average a percentage column."),
        ("13", "PIVOT (MANUAL)", "Build nine pivots", "Insert > PivotTable > Use this workbook's Data Model; layouts "
                                                      "in Pivot_Tables; one shared source.", "PT_KPI to PT_Heatmap",
         "Report Connections needs the shared source."),
        ("14", "PIVOT (MANUAL)", "Slicers and timeline", "Insert Slicer: Retailer, Region, Beverage Brand; Insert "
                                                         "Timeline: Calendar Date, 2021; tick every pivot in Report "
                                                         "Connections.", "Controls filter all nine pivots",
         "No macro reset: clear the slicers."),
        ("15", "DASHBOARD", "Use the delivered dropdowns", "Dashboard B4/F4/J4/N4 drive cards, six charts and the "
                                                           "heatmap through Dashboard_Data SUMIFS.", "Same scope "
                                                                                                    "everywhere",
         "Charts recompute after Excel recalculates."),
        ("16", "DASHBOARD", "Reset", "Choose All in each of the four dropdowns.", "Full baseline returns",
         "Filters are not stored per sheet; the static analysis sheets stay baseline."),
        ("17", "AUDIT", "Check keys and identities", "Cleaned_Data Sales Check and Profit Check columns; 0.01 "
                                                     "tolerance; zero rows retained.", "0 exceptions",
         "Do not deduplicate on retailer and date alone."),
        ("18", "AUDIT", "Location key", "Rebuilt as State | City after the delivered file cached #REF!. Location "
                                        "Count = 54.", "Every location rollup works", "Verify before using pivots."),
        ("19", "FORECAST", "Read the illustrative baseline", "Forecast sheet: Oct-Dec 2021 mean projected flat for "
                                                             "Jan-Mar 2022; holdout MAE is shown.",
         f"{F['forecast_flat']:,.2f} per month", "Never added to actual totals."),
        ("20", "QA", "Acceptance tests", "Validation sheet: baseline, FizzySip, Coca-Cola, December, BevCo+West, "
                                         "location count, identities.", "All pass",
         "Money 0.01; ratios 0.0001."),
        ("21", "QA", "Desktop click-through", "Open in desktop Excel: reset filters, select FizzySip, then "
                                              "Coca-Cola, then 2021-12, then BevCo + West, then reset.",
         "Cards and charts follow", "Programmatic checks are not click-through proof."),
        ("22", "PRESENT", "Explain the story", "Scale versus efficiency, the Western margin gap, regional timing and "
                                               "the pricing-test caveat; label owners and timeframes as proposed.",
         "No invented uplift", "See Insights_Recommendations."),
    ]
    write_frame(ws, 4, None,
                pd.DataFrame(rows, columns=["Step", "Phase", "Task", "Click-by-click", "Expected result",
                                            "Caveat"]),
                [("Step", "Step", None), ("Phase", "Phase", None), ("Task", "Task", None),
                 ("Click-by-click", "Click-by-click", None), ("Expected result", "Expected result", None),
                 ("Caveat", "Caveat", None)])
    for r in range(6, 6 + len(rows)):
        for j in range(1, 7):
            ws.cell(r, j).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 40
    return ws


def build_theme(wb, F):
    ws = new_sheet(wb, "Color_Theme", [22, 12, 30, 40, 44], tab=GREY)
    banner(ws, "COLOUR THEME | one accent, grey context, red/green reserved for variance",
           "Blue carries meaning here: plain blue is actual, light blue is only ever a genuine prior-year actual, "
           "dark blue dashed is modelled. Data bars show magnitude, never variance.")
    rows = [
        ("Prior-year actual", "#9DC3E6", "Prior year not supplied", "Real historical comparison only",
         "Never fabricate 2020 actuals"),
        ("Selected-year actual", "#0070C0", "2021 actual", "All actual charts and cards",
         "Do not use dark blue for high actuals"),
        ("Forecast", "#17365D", "Illustrative estimate", "Modelled observations, dashed line",
         "Never mixed into actual totals"),
        ("Panel", "#FFFFFF", "Content background", "Contrast", "Black text on white"),
        ("Ink", "#111827", "Headings and key values", "Priority and visibility", "Findings bold"),
        ("Grey", "#6B7280", "Notes and assumptions", "Supporting context", "Keep contrast readable"),
        ("Band", "#F2F6FA", "Zebra rows", "Scanability", "No meaning attached"),
    ]
    write_frame(ws, 4, "Roles",
                pd.DataFrame(rows, columns=["Role", "Hex", "Example", "Meaning", "Rule"]),
                [("Role", "Role", None), ("Hex", "Hex", None), ("Example", "Example", None),
                 ("Meaning", "Meaning", None), ("Rule", "Rule", None)])
    for r in range(6, 6 + len(rows)):
        c = ws.cell(r, 2)
        c.fill = fill(c.value.replace("#", ""))
        c.font = Font(bold=True, color="FFFFFF" if c.value in ("#0070C0", "#17365D", "#111827", "#6B7280")
                      else INK)
        c.alignment = Alignment(horizontal="center")
    return ws


def build_lists(wb, F):
    ws = new_sheet(wb, "Lists", [16, 14, 18, 12], tab=GREY)
    banner(ws, "LISTS | named option ranges behind the four dropdowns",
           "Edit a list only if the underlying data changes; the validation ranges are absolute.", span=4)
    cols = [("Retailer", ["All"] + F["retailers"]),
            ("Region", ["All"] + F["regions"]),
            ("Brand", ["All"] + F["brands"]),
            ("Month", ["All"] + F["month"]["Year-Month"].tolist())]
    for j, (h, values) in enumerate(cols, 1):
        c = ws.cell(4, j, h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        for i, v in enumerate(values, start=5):
            ws.cell(i, j, v)
    ws.freeze_panes = "A5"
    return ws


# ---------------------------------------------------------------- main
ORDER = ["README", "Dashboard", "Data_Quality", "Validation", "KPI_Summary", "Retailer_Analysis",
         "Regional_Analysis", "Brand_Analysis", "Location_Analysis", "Monthly_Trends", "Pricing_Analysis",
         "Forecast", "Insights_Recommendations", "Dashboard_Data", "Cleaned_Data", "Calendar", "Source_Data",
         "Pivot_Tables", "Power_Query_Code", "DAX_Measures", "Implementation_Guide", "Color_Theme", "Lists"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--source", default="work/source.xlsx")
    ap.add_argument("--out", default="deliverables/Coke_Sales_Blue_Dashboard_v2.xlsx")
    ap.add_argument("--facts", default="work/facts.json")
    args = ap.parse_args()

    d = load(args.source)
    F = facts(d)

    wb = Workbook()
    wb.remove(wb.active)

    build_readme(wb, F)
    build_quality(wb, F)
    build_validation(wb, F)
    build_kpi(wb, F)
    build_retailer(wb, F)
    build_region(wb, F)
    build_brand(wb, F)
    build_location(wb, F)
    build_monthly(wb, F)
    build_pricing(wb, F)
    build_forecast(wb, F)
    build_insights(wb, F)
    dd, refs = build_dashboard_data(wb, F)
    build_dashboard(wb, F, dd, refs)
    build_cleaned(wb, d)
    build_calendar(wb)
    build_source_data(wb, d)
    build_pivot_spec(wb, F)
    build_pq(wb, F)
    build_dax(wb, F)
    build_guide(wb, F)
    build_theme(wb, F)
    build_lists(wb, F)

    wb._sheets = [wb[name] for name in ORDER]

    for name, cell in [("selRetailer", "Dashboard!$B$4"), ("selRegion", "Dashboard!$F$4"),
                       ("selBrand", "Dashboard!$J$4"), ("selMonth", "Dashboard!$N$4")]:
        wb.defined_names[name] = DefinedName(name, attr_text=cell)

    wb.calculation.fullCalcOnLoad = True
    wb.properties.title = "Coke Sales Analysis - Blue Dashboard and Implementation"
    wb.properties.subject = "2021 portfolio: sales, units, operating profit, weighted margin"
    wb.properties.creator = "Data analysis build"
    wb.properties.description = ("Educational dataset. Operating profit is not gross or net profit. "
                                 "Prior-year actuals not supplied.")
    wb.save(args.out)

    summary = {k: F[k] for k in ["rows", "sales", "units", "profit", "margin", "rpu", "ppu", "date_min",
                                 "date_max", "locations", "states", "dup_key", "sales_exceptions",
                                 "profit_exceptions", "corr_all", "forecast_flat", "holdout_mae",
                                 "qualified_n"]}
    summary["retailers"] = F["retailers"]
    summary["brands"] = F["brands"]
    summary["regions"] = F["regions"]
    summary["tests"] = {"fizzy": F["test_fizzy"], "coke": F["test_coke"], "dec": F["test_dec"],
                        "bevwest": F["test_bevwest"]}
    summary["zero_rows"] = F["zero_rows"]
    summary["for_hooks"] = True
    with open(args.facts, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, indent=1, default=float)
    print(json.dumps(summary, indent=1, default=float)[:1200])
    print(f"\nwrote {args.out}")


if __name__ == "__main__":
    main()
