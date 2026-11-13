"""Builds the workbook skeleton, data model, formulas and dashboard layout.

Native PivotTables / slicers / pivot charts are added afterwards by build_pivots.py
(Excel COM), because openpyxl cannot create them.
"""
from __future__ import annotations

import json
import os

import openpyxl
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import ScatterChart, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.shapes import GraphicalProperties

import style as S
from content_loom import build_loom, build_qa, CHECKLIST, money, pct
from content_report import build_insights, build_report

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
SOURCE = os.path.join(HERE, "..", "source.xlsx")
OUT = os.path.join(ROOT, "deliverables", "Coca_Cola_Sales_Analysis_base.xlsx")

F = json.load(open(os.path.join(HERE, "..", "facts.json"), encoding="utf-8"))
T = F["total"]
RET = {r["name"]: r for r in F["levels"]["retailer"]}
REG = {r["name"]: r for r in F["levels"]["region"]}
BRD = {r["name"]: r for r in F["levels"]["brand"]}
MONTHS = F["months"]

SHEETS = [
    ("00_ReadMe", "Start here: purpose, scope, sheet map, definitions, capability notes"),
    ("01_Raw_Data", "The supplied sheet, copied unchanged (titles, markers and stray cells included)"),
    ("02_Clean_Data", "Analysis-ready Excel Table tblSales: 3,888 records, source row kept, checks added"),
    ("03_Data_Dictionary", "Every field: definition, type, source column, derivation"),
    ("04_Data_Quality", "Data_Quality_Log, raw-vs-clean reconciliation, flagged records"),
    ("05_KPI_Summary", "Portfolio KPIs and Coca-Cola brand view, all as ratios of sums"),
    ("06_Retailer_Analysis", "Retailer sales, units, profit, weighted margin, shares, retailer x region"),
    ("07_Region_Analysis", "Region KPIs and the 50-state ranking with a materiality flag"),
    ("08_Brand_Analysis", "Six-brand view, Coca-Cola brand vs rest, brand x region margins"),
    ("09_Monthly_Analysis", "12 months, MoM growth, margin change in pp, region x month sales"),
    ("10_Price_Volume", "Price bands, price-volume association, city x brand scatter source"),
    ("11_Pivot_Tables", "Native PivotTables on a shared cache (built with Excel COM)"),
    ("12_Dashboard", "Native slicers + pivot charts + GETPIVOTDATA KPI cards"),
    ("13_Insights_Actions", "Three to five prioritised actions with evidence and evidence gaps"),
    ("14_Implementation", "Step-by-step Excel build guide for every finding"),
    ("15_Validation", "Check / Expected / Actual / Pass-Fail / Explanation"),
    ("16_Loom_Script", "Timestamped 6-8 minute script, evaluator Q&A, recording checklist"),
]

CLEAN_COLS = [
    ("Record ID", 11, None), ("Retailer", 14, None), ("Retailer ID", 12, None),
    ("Invoice Date", 12, S.DATE), ("Region", 14, None), ("State", 16, None),
    ("City", 16, None), ("Location", 26, None), ("Beverage Brand", 15, None),
    ("Price per Unit", 12, S.MONEY2), ("Units Sold", 11, S.INT), ("Total Sales", 13, S.MONEY2),
    ("Operating Profit", 14, S.MONEY2), ("Operating Margin (source)", 12, S.PCT2),
    ("Year", 7, "0"), ("Month Number", 8, "0"), ("Month Name", 11, None),
    ("Month Start", 12, S.DATE), ("Year-Month", 10, None), ("Quarter", 8, '"Q"0'),
    ("Source Sheet", 12, None), ("Source Row", 10, "0"),
    ("Expected Sales (Price x Units)", 14, S.MONEY2), ("Sales Difference", 13, S.MONEY2),
    ("Recalculated Margin", 13, S.PCT2), ("Margin Difference (pp)", 12, S.PP),
    ("Validation Status", 26, None),
]


def load_source_rows():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    ws = wb["Data"]
    out = []
    for r in range(1, 3894):
        out.append([ws.cell(r, c).value for c in range(1, 19)])
    return out


# ---------------------------------------------------------------- 00_ReadMe
def s_readme(wb):
    ws = wb.create_sheet("00_ReadMe")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 30, "B": 104})
    S.title(ws, "COCA-COLA SALES ANALYSIS  |  U.S. RETAILERS, 2021", last_col=2)
    ws.merge_cells("A2:B2")
    c = ws.cell(2, 1, "Excel sales & operating-profitability project  |  Presenter: Nishant  |  "
                      "Analytical scope: all six supplied beverage brands, 2021")
    c.font = Font(size=10, color=S.GREY, italic=True)
    c.alignment = Alignment(indent=1, vertical="center")

    r = 4
    S.section(ws, r, "BUSINESS QUESTION", 2); r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    c = ws.cell(r, 1, "How can Coca-Cola optimise sales and operating profitability across U.S. retailers, "
                      "regions, products and time?")
    c.font = Font(size=11, bold=True, color=S.DARK)
    c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
    ws.row_dimensions[r].height = 32

    r += 2
    S.section(ws, r, "SHEET MAP", 2); r += 1
    S.headers(ws, r, ["Sheet", "What it is for"], height=18); r += 1
    for name, purpose in SHEETS:
        S.put(ws, r, 1, name, bold=True, color=S.DARK)
        S.put(ws, r, 2, purpose, align="left", wrap=True)
        r += 1

    r += 1
    S.section(ws, r, "DATA SCOPE (verified from the source file)", 2); r += 1
    scope = [
        ("Original file", F["filename"]),
        ("Raw sheet used", "Data  (headers on row 5, records on rows 6-3893)"),
        ("Records", f"{F['records']:,} business records; {len(F['headers'])} business fields"),
        ("Period", f"{F['date_min']} to {F['date_max']}  ({len(F['years'])} year, 12 months, 324 records each)"),
        ("Entities", f"{F['counts']['retailers']} retailers, {F['counts']['regions']} regions, "
                     f"{F['counts']['states']} states, {F['counts']['cities']} cities, {F['counts']['brands']} brands"),
        ("Record grain (assumed)", "One retailer x city x brand x month observation. No order ID exists, "
                                   "so a record is not an order."),
        ("Currency", "Not documented in the source. Presented as USD on the assumption that this is U.S. retail; "
                     "flagged on 04_Data_Quality."),
        ("Profit definition", "The source supplies 'Operating Profit' / 'Operating Margin' only. No net profit exists, "
                              "so 'net profit' is never used."),
        ("Status", "Educational assignment dataset. Not audited Coca-Cola company reporting."),
    ]
    for k, v in scope:
        S.put(ws, r, 1, k, bold=True)
        S.put(ws, r, 2, v, align="left", wrap=True)
        ws.row_dimensions[r].height = 26 if len(v) > 90 else 16
        r += 1

    r += 1
    S.section(ws, r, "METRIC DEFINITIONS (identical in every sheet, the report and the narration)", 2); r += 1
    defs = [
        ("Total Sales / Units / Operating Profit", "SUM over tblSales. Additive."),
        ("Weighted Operating Margin", "Total Operating Profit / Total Sales. A ratio of sums - never an average of row margins."),
        ("Realised Sales per Unit", "Total Sales / Total Units. Reported next to the unit-weighted listed price."),
        ("Unit-weighted Listed Price", "SUMPRODUCT(Price per Unit, Units Sold) / SUM(Units Sold). Not a simple average of price."),
        ("Operating Profit per Unit", "Total Operating Profit / Total Units."),
        ("Sales Share / Profit Share", "Segment value / the same measure in the stated comparison scope (all six brands, full 2021, unless labelled otherwise)."),
        ("Month-over-Month Growth", "Current month sales / previous month sales - 1."),
        ("Margin change", "Reported in percentage points (pp), not as percentage growth."),
    ]
    for k, v in defs:
        S.put(ws, r, 1, k, bold=True)
        S.put(ws, r, 2, v, align="left", wrap=True)
        ws.row_dimensions[r].height = 26 if len(v) > 90 else 16
        r += 1

    r += 1
    S.section(ws, r, "HOW TO USE IT", 2); r += 1
    howto = [
        "1. Read this sheet, then 04_Data_Quality, then 12_Dashboard.",
        "2. On 12_Dashboard, use the three native slicers (Retailer, Region, Beverage Brand) at the right edge. They "
        "filter the pivot charts, the summary PivotTables below them and the KPI cards together, because all "
        "PivotTables share one cache. Add a date timeline with Insert > Timeline if you want one.",
        "3. Clear a slicer with the funnel icon in its header; the dashboard returns to the full-portfolio view.",
        "4. 05 to 10 are formula-driven baselines that do not respond to slicers - they are the auditable full-year numbers.",
        "5. 15_Validation lists every reconciliation, and 14_Implementation rebuilds each view by hand.",
        "6. Press Ctrl+Alt+F9 once after opening if your Excel does not recalculate automatically.",
    ]
    for line in howto:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        c = ws.cell(r, 1, line)
        c.font = Font(size=10, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.row_dimensions[r].height = 30 if len(line) > 110 else 18
        r += 1

    r += 1
    S.section(ws, r, "WHAT IS NATIVE, WHAT IS FORMULA-BASED (stated plainly)", 2); r += 1
    caps = [
        ("Native PivotTables", "18 PivotTables on one shared cache: KPI, Retailer, Region, Month, Brand, Region x Month, "
                               "Brand x Region, Retailer x Region, five single-measure chart sources, the dashboard "
                               "heatmap and four dashboard summary panels (Retailer, Region, Brand, Month) that print the "
                               "numbers beside the charts."),
        ("Native slicers", "Retailer, Region and Beverage Brand slicers, connected to every PivotTable through the "
                           "shared cache. The Invoice Date timeline is a one-click manual step (Insert > Timeline), "
                           "not automated in this build."),
        ("Native pivot charts", "Retailer, Region, Monthly sales, Monthly margin and Brand charts on 12_Dashboard, driven "
                                "by the slicers."),
        ("FORMULA-BASED", "05 to 10 are SUM / SUMIFS / SUMPRODUCT / CORREL formulas over tblSales. They are fully "
                          "recalculable but they are NOT slicer-aware - a slicer does not filter a plain SUM."),
        ("STATIC", "The city x brand scatter source on 10_Price_Volume is an aggregation of tblSales produced by "
                   "SUMIFS; the heatmap colouring is conditional formatting on a PivotTable, not a chart."),
        ("NOT PRESENT", "No macros, no Power Query connections, no Power Pivot Data Model, no DAX measures, no external "
                        "links, no forecast of future periods."),
    ]
    for k, v in caps:
        S.put(ws, r, 1, k, bold=True)
        S.put(ws, r, 2, v, align="left", wrap=True)
        ws.row_dimensions[r].height = 30 if len(v) > 90 else 16
        r += 1
    ws.sheet_properties.tabColor = S.DARK


# ---------------------------------------------------------------- 01_Raw_Data
def s_raw(wb):
    ws = wb.create_sheet("01_Raw_Data")
    data = load_source_rows()
    for i, row in enumerate(data, start=1):
        for j, v in enumerate(row, start=1):
            if v is not None:
                ws.cell(i, j, v)
    for c in range(1, 19):
        cell = ws.cell(5, c)
        cell.font = Font(bold=True, size=9, color=S.WHITE)
        cell.fill = S.fill(S.BLUE)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in range(1, 19):
        ws.cell(1, c).font = Font(bold=True, size=12, color=S.DARK)
    ws.cell(2, 1).font = Font(bold=True, size=10, color=S.GREY)
    ws.freeze_panes = "B6"
    ws.column_dimensions["A"].width = 5
    for col in "BCDEFGHIJKLM":
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["P"].width = 10
    ws.column_dimensions["Q"].width = 9
    ws.column_dimensions["R"].width = 9
    for r in range(6, 3894):
        ws.cell(r, 4).number_format = S.DATE
        for c in (9, 10, 11, 12, 13):
            ws.cell(r, c).number_format = S.MONEY2 if c in (9, 11, 12) else S.INT
        ws.cell(r, 14).number_format = S.PCT
    ws.auto_filter.ref = "B5:M3893"
    ws.sheet_properties.tabColor = "FF9CA3AF"


# ---------------------------------------------------------------- 02_Clean_Data
def s_clean(wb):
    ws = wb.create_sheet("02_Clean_Data")
    for i, (name, w, _) in enumerate(CLEAN_COLS, start=1):
        ws.cell(1, i, name)
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    S.headers(ws, 1, [c[0] for c in CLEAN_COLS], height=34)

    src = openpyxl.load_workbook(SOURCE, data_only=True)["Data"]
    month_names = ["January", "February", "March", "April", "May", "June", "July",
                   "August", "September", "October", "November", "December"]
    for idx, r in enumerate(range(6, 3894), start=2):
        (ret, rid, dt, reg, st, city, brand, price, units, sales, profit, margin) = [
            src.cell(r, c).value for c in range(2, 14)]
        ws.cell(idx, 1, f"R{idx-1:04d}")
        ws.cell(idx, 2, ret)
        ws.cell(idx, 3, str(rid))
        ws.cell(idx, 4, dt)
        ws.cell(idx, 5, reg)
        ws.cell(idx, 6, st)
        ws.cell(idx, 7, city)
        ws.cell(idx, 8, f"{st} | {city}")
        ws.cell(idx, 9, brand)
        ws.cell(idx, 10, float(price))
        ws.cell(idx, 11, int(units))
        ws.cell(idx, 12, float(sales))
        ws.cell(idx, 13, float(profit))
        ws.cell(idx, 14, float(margin))
        ws.cell(idx, 15, dt.year)
        ws.cell(idx, 16, dt.month)
        ws.cell(idx, 17, month_names[dt.month - 1])
        ws.cell(idx, 18, dt.replace(day=1))
        ws.cell(idx, 19, f"{dt.year}-{dt.month:02d}")
        ws.cell(idx, 20, (dt.month - 1) // 3 + 1)
        ws.cell(idx, 21, "Data")
        ws.cell(idx, 22, r)
        ws.cell(idx, 23, f"=J{idx}*K{idx}")
        ws.cell(idx, 24, f"=L{idx}-W{idx}")
        ws.cell(idx, 25, f"=IF(L{idx}=0,NA(),M{idx}/L{idx})")
        ws.cell(idx, 26, f'=IF(OR(L{idx}=0,ISNA(Y{idx})),"",N{idx}-Y{idx})')
        ws.cell(idx, 27, (
            f'=IF(L{idx}=0,"Zero sales - check",'
            f'IF(ABS(X{idx})>0.01,"Sales mismatch",'
            f'IF(ISNA(Y{idx}),"Margin undefined",'
            f'IF(ABS(Z{idx})>0.0005,"Margin mismatch","OK"))))'))

    last = 3889
    tab = Table(displayName="tblSales", ref=f"A1:AA{last}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tab)
    ws.freeze_panes = "B2"
    for r in range(2, last + 1):
        ws.cell(r, 4).number_format = S.DATE
        ws.cell(r, 18).number_format = S.DATE
        ws.cell(r, 10).number_format = S.MONEY2
        ws.cell(r, 11).number_format = S.INT
        ws.cell(r, 12).number_format = S.MONEY2
        ws.cell(r, 13).number_format = S.MONEY2
        ws.cell(r, 14).number_format = S.PCT2
        ws.cell(r, 23).number_format = S.MONEY2
        ws.cell(r, 24).number_format = S.MONEY2
        ws.cell(r, 25).number_format = S.PCT2
        ws.cell(r, 26).number_format = S.PP
    # highlight exceptions
    ws.conditional_formatting.add(
        f"AA2:AA{last}",
        CellIsRule(operator="equal", formula=['"OK"'], font=Font(color=S.GREEN)))
    ws.conditional_formatting.add(
        f"AA2:AA{last}",
        CellIsRule(operator="notEqual", formula=['"OK"'], font=Font(color=S.RED, bold=True),
                   fill=S.fill("FFFDE7E7")))
    ws.sheet_properties.tabColor = S.BLUE


# ---------------------------------------------------------------- 03_Data_Dictionary
def s_dictionary(wb):
    ws = wb.create_sheet("03_Data_Dictionary")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 28, "B": 58, "C": 13, "D": 18, "E": 46})
    S.title(ws, "DATA DICTIONARY  |  tblSales", "Every field, its definition, and where it comes from. "
            "Helper fields are derived in Excel and never overwrite a source value.", last_col=5)
    S.headers(ws, 4, ["Field", "Definition", "Type", "Source column", "Derivation / notes"], height=22)
    rows = [
        ("Record ID", "Stable identifier for the cleaned record.", "Text", "new", "R0001... in source-row order."),
        ("Retailer", "Retailer that supplied the sales record.", "Text", "B", "Unchanged. Four values."),
        ("Retailer ID", "Numeric retailer identifier, held as text.", "Text", "C", "Text so no identifier is truncated."),
        ("Invoice Date", "Date the invoice was recorded.", "Date", "D", "All valid, 2021 only."),
        ("Region", "U.S. region of the city.", "Text", "E", "Unchanged. Five values."),
        ("State", "U.S. state of the city.", "Text", "F", "Unchanged. 50 values."),
        ("City", "City of the record.", "Text", "G", "Unchanged."),
        ("Location", "State | City key used for city-level analysis.", "Text", "new", "Concatenation; no two states share a city name here."),
        ("Beverage Brand", "Beverage brand sold.", "Text", "H", "Unchanged. Six brands, portfolio scope."),
        ("Price per Unit", "Listed price per unit at the observation.", "Currency", "I", "Unchanged. $0.10-$1.10."),
        ("Units Sold", "Units sold in the observation.", "Integer", "J", "Unchanged."),
        ("Total Sales", "Recorded sales value.", "Currency", "K", "Unchanged. Never adjusted."),
        ("Operating Profit", "Recorded operating profit.", "Currency", "L", "Unchanged. Operating, not net, profit."),
        ("Operating Margin (source)", "Margin as supplied in the file.", "Percent", "M", "Unchanged. Kept for the identity test."),
        ("Year", "Calendar year.", "Integer", "new", "=YEAR(Invoice Date). 2021 throughout."),
        ("Month Number", "Calendar month 1-12.", "Integer", "new", "=MONTH(Invoice Date)."),
        ("Month Name", "Month name.", "Text", "new", "Derived from the month number."),
        ("Month Start", "First day of the month.", "Date", "new", "=DATE(Year, Month, 1). Used for chronological sorting."),
        ("Year-Month", "Reporting label YYYY-MM.", "Text", "new", "Sortable label; the date column drives the sort."),
        ("Quarter", "Calendar quarter.", "Text", "new", "Q1-Q4."),
        ("Source Sheet", "Sheet the record came from.", "Text", "new", "Always 'Data'."),
        ("Source Row", "Row number in 01_Raw_Data.", "Integer", "new", "Enables row-level audit back to the original."),
        ("Expected Sales (Price x Units)", "Price per unit x units sold.", "Currency", "new", "=Price per Unit * Units Sold."),
        ("Sales Difference", "Source sales minus expected sales.", "Currency", "new", "0 for every record."),
        ("Recalculated Margin", "Operating profit / total sales.", "Percent", "new", "NA() when sales = 0."),
        ("Margin Difference (pp)", "Source margin minus recalculated margin.", "Percent", "new", "Blank when undefined."),
        ("Validation Status", "OK, or the specific identity failure.", "Text", "new", "Conditional formatting flags non-OK rows."),
    ]
    r = 5
    for row in rows:
        for i, v in enumerate(row, start=1):
            S.put(ws, r, i, v, align="left" if i in (1, 2, 4, 5) else "center",
                  bold=(i == 1), wrap=(i in (2, 5)), literal=True)
        r += 1
    S.note(ws, r + 1, "No source column is overwritten. Total Sales and Operating Profit keep their recorded values "
                      "so the identity tests on 04_Data_Quality remain meaningful.", last_col=5)
    ws.freeze_panes = "A5"
    ws.sheet_properties.tabColor = S.LIGHT


# ---------------------------------------------------------------- 04_Data_Quality
def s_quality(wb):
    ws = wb.create_sheet("04_Data_Quality")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 44, "B": 20, "C": 14, "D": 46, "E": 46, "F": 26})
    S.title(ws, "DATA QUALITY LOG  |  what was checked, what was changed, what it cost",
            "Every issue found, its treatment and its effect on totals. Nothing was deleted to make a number match.",
            last_col=6)
    S.headers(ws, 4, ["Issue", "Field", "Affected rows", "Treatment", "Justification", "Impact on totals"], height=30)
    q = F["quality"]
    logs = [
        ("Two decorative title rows and one blank row above the header", "Rows 2-4", "3", 
         "Excluded from the analysis table; kept on 01_Raw_Data", "They are titles, not records", "None"),
        ("47 'x' markers in column A", "Column A", "47", "Not imported into tblSales",
         "Marker cells, not a business field", "None"),
        ("Stray unlabelled numbers in columns P:R (last four rows)", "P:R, rows 3890-3893", "4",
         "Not imported; noted here", "Outside the B:M record block; no field name, no unit", "None"),
        ("Zero units, sales and profit with a supplied 40% margin",
         "Total Sales / Operating Profit / Operating Margin", "2",
         "Retained and flagged 'Zero sales - check'",
         "A margin cannot exist on a zero base; deleting could hide a real data-entry gap",
         "Zero. Sales and profit are 0 in both rows"),
        ("Sales identity: Price x Units vs Total Sales", "Total Sales", f"{q['sales_identity_breaks']} exceptions",
         "No change; difference column added", "Source values are preserved, not overwritten", "None"),
        ("Margin identity: Operating Profit / Total Sales vs supplied margin", "Operating Margin",
         f"{q['margin_identity_breaks']} exceptions", "No change; recalculated margin added",
         "Rounding and the 2 zero-base rows only", "None"),
        ("Exact duplicate records on all 12 business fields", "All", f"{q['exact_dupes']}",
         "No action", "No duplicates found", "None"),
        ("Duplicate retailer-location-brand-month keys", "Key fields", f"{q['key_dupes']}",
         "No action", "Monthly grain confirmed unique", "None"),
        ("Negative sales, units or profit", "Numeric fields", f"{q['negatives']}",
         "No action", "No negatives found; returns would have been retained, not deleted", "None"),
        ("Missing or invalid invoice dates", "Invoice Date", "0",
         "No action", "All 3,888 dates parse and fall in 2021", "None"),
        ("Incomplete or partial months", "Invoice Date", "0",
         "No action", "Each month holds exactly 324 records", "None"),
        ("Floating-point residue (e.g. 5637.500000000001)", "Numeric fields", "many",
         "Formatted display; 0.01 tolerance in identity tests", "Storage artefact, not a data error", "Below $0.01"),
        ("Currency not documented", "Price / Sales / Profit", "all",
         "Labelled USD as an assumption", "No currency field or note in the source", "No numeric change; label only"),
        ("Operating vs net profit not documented", "Operating Profit", "all",
         "'Operating profit' used everywhere", "The only profit field supplied", "None"),
        ("Retailer and region are confounded", "Retailer / Region", "all",
         "Disclosed on every regional comparison", "BevCo sells only in the South; FizzySip only in the West", "None on totals; limits attribution"),
    ]
    r = 5
    for row in logs:
        for i, v in enumerate(row, start=1):
            S.put(ws, r, i, v, align="left" if i in (1, 2, 4, 5) else "center", wrap=True)
        ws.row_dimensions[r].height = 40
        r += 1

    r += 1
    S.section(ws, r, "RECONCILIATION  |  raw sheet vs analysis table (live formulas)", 6); r += 1
    S.headers(ws, r, ["Measure", "Raw sheet (01_Raw_Data)", "Clean table (tblSales)", "Difference", "Status", ""],
              height=22)
    r += 1
    rec_rows = [
        ("Records", "=COUNTA('01_Raw_Data'!B6:B3893)", "=COUNTA(tblSales[Record ID])"),
        ("Total Sales", "=SUM('01_Raw_Data'!K6:K3893)", "=SUM(tblSales[Total Sales])"),
        ("Units Sold", "=SUM('01_Raw_Data'!J6:J3893)", "=SUM(tblSales[Units Sold])"),
        ("Operating Profit", "=SUM('01_Raw_Data'!L6:L3893)", "=SUM(tblSales[Operating Profit])"),
    ]
    first = r
    for label, a, b in rec_rows:
        S.put(ws, r, 1, label, bold=True)
        fmt = S.INT if label in ("Records", "Units Sold") else S.MONEY2
        S.put(ws, r, 2, a, fmt=fmt)
        S.put(ws, r, 3, b, fmt=fmt)
        S.put(ws, r, 4, f"=C{r}-B{r}", fmt=fmt)
        S.put(ws, r, 5, f'=IF(ABS(D{r})<=0.01,"Pass","Fail")', align="center")
        r += 1
    ws.conditional_formatting.add(f"E{first}:E{r-1}",
                                  CellIsRule(operator="equal", formula=['"Pass"'], font=Font(color=S.GREEN, bold=True)))

    r += 1
    S.section(ws, r, "FLAGGED RECORDS  |  retained, not deleted", 6); r += 1
    S.headers(ws, r, ["Retailer", "Location", "Brand", "Invoice date", "Supplied margin", "Why it is flagged"],
              height=22)
    r += 1
    for z in q["zero_sales"]:
        S.put(ws, r, 1, z["retailer"], align="left")
        S.put(ws, r, 2, z["location"].replace(" | ", ", "), align="left")
        S.put(ws, r, 3, z["brand"], align="left")
        S.put(ws, r, 4, z["date"], align="left")
        S.put(ws, r, 5, z["supplied_margin"], fmt=S.PCT, align="center")
        S.put(ws, r, 6, "Zero units, sales and profit; margin has no base", align="left", wrap=True)
        r += 1
    S.note(ws, r + 1, "Both records keep their source values. Their combined effect on Total Sales, Units and "
                      "Operating Profit is zero, and they are excluded from margin calculations by construction.",
           last_col=6)
    ws.sheet_properties.tabColor = "FFDC2626"


# ---------------------------------------------------------------- 05_KPI_Summary
def s_kpi(wb):
    ws = wb.create_sheet("05_KPI_Summary")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 36, "B": 18, "C": 20, "D": 78, "E": 30})
    S.title(ws, "KPI SUMMARY  |  ratios of sums, not averages of ratios",
            "All six brands, full 2021. Every value is a live formula over tblSales.", last_col=5)

    S.section(ws, 4, "PORTFOLIO KPIs  |  scope: all six brands, 3,888 records, 2021-01-02 to 2021-12-25", 5)
    S.headers(ws, 5, ["Metric", "Value", "Format", "Excel formula / definition", "Scope note"], height=22)
    rows = [
        ("Total Sales", "=SUM(tblSales[Total Sales])", S.MONEY, "SUM(tblSales[Total Sales])", "Additive revenue"),
        ("Total Units", "=SUM(tblSales[Units Sold])", S.INT, "SUM(tblSales[Units Sold])", "Additive volume"),
        ("Operating Profit", "=SUM(tblSales[Operating Profit])", S.MONEY, "SUM(tblSales[Operating Profit])", "Operating, not net, profit"),
        ("Weighted Operating Margin",
         "=IF(SUM(tblSales[Total Sales])=0,NA(),SUM(tblSales[Operating Profit])/SUM(tblSales[Total Sales]))",
         S.PCT2, "Operating Profit / Total Sales", "Ratio of sums"),
        ("Realised Sales per Unit",
         "=IF(SUM(tblSales[Units Sold])=0,NA(),SUM(tblSales[Total Sales])/SUM(tblSales[Units Sold]))",
         S.MONEY4, "Total Sales / Total Units", "Not the listed price"),
        ("Unit-weighted Listed Price",
         "=IF(SUM(tblSales[Units Sold])=0,NA(),SUMPRODUCT(tblSales[Price per Unit],tblSales[Units Sold])/SUM(tblSales[Units Sold]))",
         S.MONEY4, "SUMPRODUCT(Price, Units) / SUM(Units)", "Weighted by volume"),
        ("Operating Profit per Unit",
         "=IF(SUM(tblSales[Units Sold])=0,NA(),SUM(tblSales[Operating Profit])/SUM(tblSales[Units Sold]))",
         S.MONEY4, "Operating Profit / Total Units", "Contribution per unit"),
        ("Records analysed", "=COUNTA(tblSales[Record ID])", S.INT, "COUNTA(tblSales[Record ID])", "Not an order count"),
        ("Cities covered", "=ROUND(SUMPRODUCT(1/COUNTIF(tblSales[Location],tblSales[Location])),0)", S.INT,
         "Distinct Location", "State | City"),
        ("Brands covered", "=ROUND(SUMPRODUCT(1/COUNTIF(tblSales[Beverage Brand],tblSales[Beverage Brand])),0)", S.INT,
         "Distinct Beverage Brand", "Full portfolio"),
    ]
    r = 6
    for label, formula, fmt, defn, sc in rows:
        S.put(ws, r, 1, label, bold=True)
        S.put(ws, r, 2, formula, fmt=fmt, bold=True, color=S.DARK)
        S.put(ws, r, 3, "%" if fmt in (S.PCT, S.PCT2) else "$" if "MONEY" in str(fmt) else "count", align="center")
        S.put(ws, r, 4, defn, align="left", wrap=True)
        S.put(ws, r, 5, sc, align="left", wrap=True)
        r += 1

    r += 1
    S.section(ws, r, "METHOD CHECK  |  why the average of row margins is wrong", 5); r += 1
    S.headers(ws, r, ["Metric", "Value", "", "What it shows", ""], height=18); r += 1
    a = r
    S.put(ws, a, 1, "Simple average of row margins", bold=True)
    S.put(ws, a, 2, "=AVERAGE(tblSales[Operating Margin (source)])", fmt=S.PCT2)
    S.put(ws, a, 4, "Every row counts equally, so a 4,000-unit city weighs the same as a 12,000-unit city.", align="left", wrap=True)
    S.put(ws, a + 1, 1, "Weighted (correct) margin", bold=True)
    S.put(ws, a + 1, 2, "=B9", fmt=S.PCT2)
    S.put(ws, a + 1, 4, "Total operating profit / total sales. Used everywhere in this workbook.", align="left", wrap=True)
    S.put(ws, a + 2, 1, "Overstatement from averaging (pp)", bold=True)
    S.put(ws, a + 2, 2, f"=(B{a}-B{a+1})*100", fmt=S.PP)
    S.put(ws, a + 2, 4, "The shortcut differs from the ratio of sums, so it is never used here.", align="left", wrap=True)

    r = a + 4
    S.section(ws, r, "COCA-COLA BRAND VIEW  |  scope: Beverage Brand = 'Coca-Cola' only (one of six brands)", 5); r += 1
    S.headers(ws, r, ["Metric", "Coca-Cola brand", "Share of portfolio", "Formula", "Reading"], height=22); r += 1
    brand_rows = [
        ("Total Sales", "=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],\"Coca-Cola\")", S.MONEY,
         "SUMIFS on Beverage Brand", "Brand revenue, not portfolio revenue"),
        ("Total Units", "=SUMIFS(tblSales[Units Sold],tblSales[Beverage Brand],\"Coca-Cola\")", S.INT,
         "SUMIFS on Beverage Brand", "Brand volume"),
        ("Operating Profit", "=SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],\"Coca-Cola\")", S.MONEY,
         "SUMIFS on Beverage Brand", "Brand profit"),
        ("Weighted Operating Margin", None, S.PCT2, "Profit / Sales for the brand", "Compare with the 36.5% portfolio"),
        ("Operating Profit per Unit", None, S.MONEY4, "Profit / Units for the brand", "Brand contribution per unit"),
    ]
    start = r
    for i, (label, formula, fmt, defn, read) in enumerate(brand_rows):
        row = r + i
        S.put(ws, row, 1, label, bold=True)
        if label == "Weighted Operating Margin":
            formula = f"=IF(B{start}=0,NA(),B{start+2}/B{start})"
        if label == "Operating Profit per Unit":
            formula = f"=IF(B{start+1}=0,NA(),B{start+2}/B{start+1})"
        S.put(ws, row, 2, formula, fmt=fmt, bold=True, color=S.DARK)
        share_row = {"Total Sales": 6, "Total Units": 7, "Operating Profit": 8}.get(label)
        if share_row:
            S.put(ws, row, 3, f"=IF($B${share_row}=0,NA(),B{row}/$B${share_row})", fmt=S.PCT)
        else:
            S.put(ws, row, 3, "", align="center")
        S.put(ws, row, 4, defn, align="left", wrap=True)
        S.put(ws, row, 5, read, align="left", wrap=True)

    r = start + len(brand_rows) + 1
    S.note(ws, r, "Plain SUM and SUMIFS formulas do not become slicer-aware because a slicer sits nearby. The "
                  "dashboard KPI cards use GETPIVOTDATA against the filtered PivotTable instead; this sheet is the "
                  "full-year auditable baseline.", last_col=5)
    r += 2
    S.note(ws, r, "Currency is assumed USD (not documented in the source). 'Operating profit' is the supplied field; "
                  "no net-profit figure exists in this dataset.", last_col=5)
    ws.sheet_properties.tabColor = S.BLUE


# ---------------------------------------------------------------- 06_Retailer
def s_retailer(wb):
    ws = wb.create_sheet("06_Retailer_Analysis")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 18, "B": 14, "C": 12, "D": 15, "E": 13, "F": 13, "G": 12, "H": 11, "I": 11,
                  "J": 14, "K": 10, "L": 9})
    S.title(ws, "RETAILER ANALYSIS  |  scale, profitability and share",
            "Scope: all six brands, full 2021. Shares are of the all-brand portfolio.", last_col=12)
    S.headers(ws, 4, ["Retailer", "Sales", "Units", "Operating Profit", "Weighted Margin", "Realised $/Unit",
                      "Profit/Unit", "Sales Share", "Profit Share", "Margin vs portfolio (pp)", "Locations",
                      "Records"], height=34)
    names = [r["name"] for r in F["levels"]["retailer"]]
    r = 5
    for name in names:
        S.put(ws, r, 1, name, bold=True, align="left")
        S.put(ws, r, 2, f"=SUMIFS(tblSales[Total Sales],tblSales[Retailer],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 3, f"=SUMIFS(tblSales[Units Sold],tblSales[Retailer],$A{r})", fmt=S.INT)
        S.put(ws, r, 4, f"=SUMIFS(tblSales[Operating Profit],tblSales[Retailer],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 5, f"=IF(B{r}=0,NA(),D{r}/B{r})", fmt=S.PCT2)
        S.put(ws, r, 6, f"=IF(C{r}=0,NA(),B{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 7, f"=IF(C{r}=0,NA(),D{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 8, f"=IF($B$9=0,NA(),B{r}/$B$9)", fmt=S.PCT)
        S.put(ws, r, 9, f"=IF($D$9=0,NA(),D{r}/$D$9)", fmt=S.PCT)
        S.put(ws, r, 10, f"=(E{r}-$E$9)*100", fmt=S.PP)
        S.put(ws, r, 11, f"=ROUND(SUMPRODUCT((tblSales[Retailer]=$A{r})/COUNTIFS(tblSales[Retailer],"
                         f"tblSales[Retailer],tblSales[Location],tblSales[Location])),0)", fmt=S.INT)
        S.put(ws, r, 12, f"=COUNTIFS(tblSales[Retailer],$A{r})", fmt=S.INT)
        r += 1
    S.put(ws, 9, 1, "Portfolio total", bold=True, align="left")
    for col, f in ((2, "=SUM(B5:B8)"), (3, "=SUM(C5:C8)"), (4, "=SUM(D5:D8)"), (11, "=SUM(K5:K8)"),
                   (12, "=SUM(L5:L8)")):
        S.put(ws, 9, col, f, fmt=S.INT if col in (3, 11, 12) else S.MONEY, bold=True)
    S.put(ws, 9, 5, "=IF(B9=0,NA(),D9/B9)", fmt=S.PCT2, bold=True)
    S.put(ws, 9, 6, "=IF(C9=0,NA(),B9/C9)", fmt=S.MONEY4, bold=True)
    S.put(ws, 9, 7, "=IF(C9=0,NA(),D9/C9)", fmt=S.MONEY4, bold=True)
    S.put(ws, 9, 8, "=IF(B9=0,NA(),B9/B9)", fmt=S.PCT, bold=True)
    S.put(ws, 9, 9, "=IF(D9=0,NA(),D9/D9)", fmt=S.PCT, bold=True)
    S.put(ws, 9, 10, "=(E9-$E$9)*100", fmt=S.PP, bold=True)
    S.total_row(ws, 9, 1, 12)

    ws.conditional_formatting.add("E5:E8", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))
    ws.conditional_formatting.add("J5:J8", CellIsRule(operator="lessThan", formula=["0"],
                                                      font=Font(color=S.RED, bold=True)))
    ws.conditional_formatting.add("B5:B8", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                       color=S.BLUE))

    r = 12
    S.section(ws, r, "RETAILER x REGION  |  sales by retailer and region (USD)", 12); r += 1
    regions = [x["name"] for x in F["levels"]["region"]]
    S.headers(ws, r, ["Retailer"] + regions + ["Total", "Regions served"], height=22); r += 1
    first_rr = r
    for name in names:
        S.put(ws, r, 1, name, bold=True, align="left")
        for j, rg in enumerate(regions, start=2):
            S.put(ws, r, j, f'=SUMIFS(tblSales[Total Sales],tblSales[Retailer],$A{r},tblSales[Region],{openpyxl.utils.get_column_letter(j)}$%d)' % (first_rr - 1),
                  fmt=S.MONEY)
        S.put(ws, r, 7, f"=SUM(B{r}:F{r})", fmt=S.MONEY, bold=True)
        S.put(ws, r, 8, f"=COUNTIF(B{r}:F{r},\">0\")", fmt=S.INT, align="center")
        r += 1
    S.put(ws, r, 1, "Total", bold=True)
    for j in range(2, 8):
        col = openpyxl.utils.get_column_letter(j)
        S.put(ws, r, j, f"=SUM({col}{first_rr}:{col}{r-1})", fmt=S.MONEY, bold=True)
    S.total_row(ws, r, 1, 8)

    r += 2
    S.section(ws, r, "HOW TO READ THIS", 12); r += 1
    notes = [
        "Scale and margin are different questions. Sodapop has the largest recorded sales; BevCo has the highest "
        "weighted margin. A larger sales figure is not evidence of greater efficiency - no store count, resource or "
        "cost measure is supplied.",
        "Retailer and geography are almost the same variable here: BevCo operates only in the South, FizzySip only in "
        "the West, and Sodapop operates in the Northeast, Southeast and Midwest. Retailer effects and region effects "
        "cannot be separated with this dataset.",
        "FizzySip's profit share is lower than its sales share because its margin is roughly three points below the "
        "portfolio, not because its prices are low - its realised sales per unit is the highest of the four accounts.",
    ]
    for n in notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        c = ws.cell(r, 1, "- " + n)
        c.font = Font(size=9, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        ws.row_dimensions[r].height = 30
        r += 1
    ws.freeze_panes = "B5"
    ws.sheet_properties.tabColor = S.BLUE


# ---------------------------------------------------------------- 07_Region
def s_region(wb):
    ws = wb.create_sheet("07_Region_Analysis")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 16, "B": 14, "C": 12, "D": 15, "E": 13, "F": 13, "G": 12, "H": 11, "I": 11,
                  "J": 14, "K": 10, "L": 8})
    S.title(ws, "REGION AND STATE ANALYSIS  |  where the sales are, and where the margin goes",
            "Scope: all six brands, full 2021. Materiality rule: a city counts as material at >= $100,000 annual sales.",
            last_col=12)
    S.headers(ws, 4, ["Region", "Sales", "Units", "Operating Profit", "Weighted Margin", "Realised $/Unit",
                      "Profit/Unit", "Sales Share", "Profit Share", "Margin vs portfolio (pp)", "Cities", "Records"],
              height=34)
    names = [x["name"] for x in F["levels"]["region"]]
    r = 5
    for name in names:
        S.put(ws, r, 1, name, bold=True, align="left")
        S.put(ws, r, 2, f"=SUMIFS(tblSales[Total Sales],tblSales[Region],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 3, f"=SUMIFS(tblSales[Units Sold],tblSales[Region],$A{r})", fmt=S.INT)
        S.put(ws, r, 4, f"=SUMIFS(tblSales[Operating Profit],tblSales[Region],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 5, f"=IF(B{r}=0,NA(),D{r}/B{r})", fmt=S.PCT2)
        S.put(ws, r, 6, f"=IF(C{r}=0,NA(),B{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 7, f"=IF(C{r}=0,NA(),D{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 8, f"=IF($B$10=0,NA(),B{r}/$B$10)", fmt=S.PCT)
        S.put(ws, r, 9, f"=IF($D$10=0,NA(),D{r}/$D$10)", fmt=S.PCT)
        S.put(ws, r, 10, f"=(E{r}-$E$10)*100", fmt=S.PP)
        S.put(ws, r, 11, f"=ROUND(SUMPRODUCT((tblSales[Region]=$A{r})/COUNTIFS(tblSales[Region],tblSales[Region],"
                         f"tblSales[Location],tblSales[Location])),0)", fmt=S.INT)
        S.put(ws, r, 12, f"=COUNTIFS(tblSales[Region],$A{r})", fmt=S.INT)
        r += 1
    S.put(ws, 10, 1, "Portfolio total", bold=True, align="left")
    S.put(ws, 10, 2, "=SUM(B5:B9)", fmt=S.MONEY, bold=True)
    S.put(ws, 10, 3, "=SUM(C5:C9)", fmt=S.INT, bold=True)
    S.put(ws, 10, 4, "=SUM(D5:D9)", fmt=S.MONEY, bold=True)
    S.put(ws, 10, 5, "=IF(B10=0,NA(),D10/B10)", fmt=S.PCT2, bold=True)
    S.put(ws, 10, 6, "=IF(C10=0,NA(),B10/C10)", fmt=S.MONEY4, bold=True)
    S.put(ws, 10, 7, "=IF(C10=0,NA(),D10/C10)", fmt=S.MONEY4, bold=True)
    S.put(ws, 10, 8, "=1", fmt=S.PCT, bold=True)
    S.put(ws, 10, 9, "=1", fmt=S.PCT, bold=True)
    S.put(ws, 10, 10, "=0", fmt=S.PP, bold=True)
    S.put(ws, 10, 11, "=SUM(K5:K9)", fmt=S.INT, bold=True)
    S.put(ws, 10, 12, "=SUM(L5:L9)", fmt=S.INT, bold=True)
    S.total_row(ws, 10, 1, 12)
    ws.conditional_formatting.add("E5:E9", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))

    r = 12
    S.section(ws, r, "STATE RANKING  |  50 states, ranked by sales (materiality flag uses annual sales)", 12); r += 1
    S.headers(ws, r, ["State", "Region", "Sales", "Units", "Operating Profit", "Weighted Margin", "Profit/Unit",
                      "Share of sales", "Materiality", "Rank", "", ""], height=30)
    r += 1
    first_state = r
    states = F["levels"]["state"]
    for st in states:
        S.put(ws, r, 1, st["name"], align="left")
        S.put(ws, r, 2, st["state"] if "state" in st else "", align="left")
        r += 1
    ws.sheet_properties.tabColor = S.LIGHT
    return ws, first_state, states


def s_region_states(ws, first_state, states):
    # region attribute per state comes from the source mapping
    src = openpyxl.load_workbook(SOURCE, data_only=True)["Data"]
    mapping = {}
    for rr in range(6, 3894):
        mapping[src.cell(rr, 6).value] = src.cell(rr, 5).value
    r = first_state
    for st in states:
        S.put(ws, r, 2, mapping.get(st["name"], ""), align="left")
        S.put(ws, r, 3, f"=SUMIFS(tblSales[Total Sales],tblSales[State],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 4, f"=SUMIFS(tblSales[Units Sold],tblSales[State],$A{r})", fmt=S.INT)
        S.put(ws, r, 5, f"=SUMIFS(tblSales[Operating Profit],tblSales[State],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 6, f"=IF(C{r}=0,NA(),E{r}/C{r})", fmt=S.PCT2)
        S.put(ws, r, 7, f"=IF(D{r}=0,NA(),E{r}/D{r})", fmt=S.MONEY4)
        S.put(ws, r, 8, f"=IF($C${first_state + len(states)}=0,NA(),C{r}/$C${first_state + len(states)})", fmt=S.PCT)
        S.put(ws, r, 9, f'=IF(C{r}>=100000,"Material","Review")', align="center")
        S.put(ws, r, 10, f"=RANK(C{r},$C${first_state}:$C${first_state + len(states) - 1})", fmt=S.INT, align="center")
        r += 1
    last = r - 1
    S.put(ws, r, 1, "Total", bold=True)
    S.put(ws, r, 3, f"=SUM(C{first_state}:C{last})", fmt=S.MONEY, bold=True)
    S.put(ws, r, 4, f"=SUM(D{first_state}:D{last})", fmt=S.INT, bold=True)
    S.put(ws, r, 5, f"=SUM(E{first_state}:E{last})", fmt=S.MONEY, bold=True)
    S.put(ws, r, 6, f"=IF(C{r}=0,NA(),E{r}/C{r})", fmt=S.PCT2, bold=True)
    S.total_row(ws, r, 1, 10)
    ws.conditional_formatting.add(f"F{first_state}:F{last}", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))
    ws.conditional_formatting.add(f"C{first_state}:C{last}", DataBarRule(start_type="num", start_value=0,
                                                                        end_type="max", color=S.BLUE))
    ws.conditional_formatting.add(f"I{first_state}:I{last}", CellIsRule(
        operator="equal", formula=['"Review"'], font=Font(color=S.GREY, italic=True)))
    r = last + 3
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(r, 1, "Materiality rule: >= $100,000 of annual sales marks a state or city as material. It is an "
                      "analyst convention chosen so that small-volume markets are not read as weak demand, not a "
                      "company policy. 37 of 54 cities clear it and produce 83.7% of portfolio sales.")
    c.font = Font(size=9, italic=True, color=S.GREY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30


# ---------------------------------------------------------------- 08_Brand
def s_brand(wb):
    ws = wb.create_sheet("08_Brand_Analysis")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 18, "B": 14, "C": 12, "D": 15, "E": 13, "F": 13, "G": 12, "H": 11, "I": 11,
                  "J": 13, "K": 12, "L": 10})
    S.title(ws, "BRAND ANALYSIS  |  the full six-brand portfolio, and the Coca-Cola brand on its own",
            "Scope warning: portfolio totals include five brands other than Coca-Cola. Brand shares are of the "
            "portfolio; they are not Coca-Cola-brand-only sales.", last_col=12)
    S.headers(ws, 4, ["Beverage Brand", "Sales", "Units", "Operating Profit", "Weighted Margin", "Realised $/Unit",
                      "Profit/Unit", "Sales Share", "Profit Share", "Margin vs portfolio (pp)", "Rank", "Records"],
              height=34)
    names = [x["name"] for x in F["levels"]["brand"]]
    r = 5
    for name in names:
        S.put(ws, r, 1, name, bold=True, align="left")
        S.put(ws, r, 2, f"=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 3, f"=SUMIFS(tblSales[Units Sold],tblSales[Beverage Brand],$A{r})", fmt=S.INT)
        S.put(ws, r, 4, f"=SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],$A{r})", fmt=S.MONEY)
        S.put(ws, r, 5, f"=IF(B{r}=0,NA(),D{r}/B{r})", fmt=S.PCT2)
        S.put(ws, r, 6, f"=IF(C{r}=0,NA(),B{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 7, f"=IF(C{r}=0,NA(),D{r}/C{r})", fmt=S.MONEY4)
        S.put(ws, r, 8, f"=IF($B$11=0,NA(),B{r}/$B$11)", fmt=S.PCT)
        S.put(ws, r, 9, f"=IF($D$11=0,NA(),D{r}/$D$11)", fmt=S.PCT)
        S.put(ws, r, 10, f"=(E{r}-$E$11)*100", fmt=S.PP)
        S.put(ws, r, 11, f"=RANK(B{r},$B$5:$B$10)", fmt=S.INT, align="center")
        S.put(ws, r, 12, f"=COUNTIFS(tblSales[Beverage Brand],$A{r})", fmt=S.INT)
        r += 1
    S.put(ws, 11, 1, "Portfolio total", bold=True, align="left")
    S.put(ws, 11, 2, "=SUM(B5:B10)", fmt=S.MONEY, bold=True)
    S.put(ws, 11, 3, "=SUM(C5:C10)", fmt=S.INT, bold=True)
    S.put(ws, 11, 4, "=SUM(D5:D10)", fmt=S.MONEY, bold=True)
    S.put(ws, 11, 5, "=IF(B11=0,NA(),D11/B11)", fmt=S.PCT2, bold=True)
    S.put(ws, 11, 6, "=IF(C11=0,NA(),B11/C11)", fmt=S.MONEY4, bold=True)
    S.put(ws, 11, 7, "=IF(C11=0,NA(),D11/C11)", fmt=S.MONEY4, bold=True)
    S.put(ws, 11, 8, "=1", fmt=S.PCT, bold=True)
    S.put(ws, 11, 9, "=1", fmt=S.PCT, bold=True)
    S.put(ws, 11, 12, "=SUM(L5:L10)", fmt=S.INT, bold=True)
    S.total_row(ws, 11, 1, 12)
    ws.conditional_formatting.add("E5:E10", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))

    r = 13
    S.section(ws, r, "BRAND x REGION  |  weighted operating margin (blank where the brand is not sold in that region)", 12)
    r += 1
    regions = [x["name"] for x in F["levels"]["region"]]
    S.headers(ws, r, ["Beverage Brand"] + regions + ["Portfolio"], height=22)
    r += 1
    hdr_row = r - 1
    for name in names:
        S.put(ws, r, 1, name, bold=True, align="left")
        for j, rg in enumerate(regions, start=2):
            col = openpyxl.utils.get_column_letter(j)
            S.put(ws, r, j,
                  f'=IFERROR(SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],$A{r},'
                  f'tblSales[Region],{col}${hdr_row})/SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],$A{r},'
                  f'tblSales[Region],{col}${hdr_row}),"")', fmt=S.PCT2)
        S.put(ws, r, 7, f"=IF(B{r-0}=0,NA(),SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],$A{r})/"
                        f"SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],$A{r}))", fmt=S.PCT2)
        r += 1
    ws.conditional_formatting.add(f"B{hdr_row+1}:G{r-1}", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))

    r += 1
    S.section(ws, r, "COCA-COLA BRAND vs THE OTHER FIVE BRANDS  |  scope stated for each row", 12); r += 1
    S.headers(ws, r, ["Scope", "Sales", "Units", "Operating Profit", "Weighted Margin", "Share of portfolio sales",
                      "Share of portfolio profit", "Profit/Unit"], height=30)
    r += 1
    rows = [
        ("Coca-Cola brand only", '"Coca-Cola"', f"=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],A{r})"),
        ("Other five brands", '<>"Coca-Cola"', f"=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],A{r})"),
    ]
    # build explicitly to keep formulas readable
    S.put(ws, r, 1, "Coca-Cola brand only", bold=True, align="left")
    S.put(ws, r, 2, '=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],"Coca-Cola")', fmt=S.MONEY)
    S.put(ws, r, 3, '=SUMIFS(tblSales[Units Sold],tblSales[Beverage Brand],"Coca-Cola")', fmt=S.INT)
    S.put(ws, r, 4, '=SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],"Coca-Cola")', fmt=S.MONEY)
    S.put(ws, r, 5, f"=IF(B{r}=0,NA(),D{r}/B{r})", fmt=S.PCT2)
    S.put(ws, r, 6, f"=IF($B$11=0,NA(),B{r}/$B$11)", fmt=S.PCT)
    S.put(ws, r, 7, f"=IF($D$11=0,NA(),D{r}/$D$11)", fmt=S.PCT)
    S.put(ws, r, 8, f"=IF(C{r}=0,NA(),D{r}/C{r})", fmt=S.MONEY4)
    S.put(ws, r + 1, 1, "Other five brands combined", bold=True, align="left")
    S.put(ws, r + 1, 2, f"=$B$11-B{r}", fmt=S.MONEY)
    S.put(ws, r + 1, 3, f"=$C$11-C{r}", fmt=S.INT)
    S.put(ws, r + 1, 4, f"=$D$11-D{r}", fmt=S.MONEY)
    S.put(ws, r + 1, 5, f"=IF(B{r+1}=0,NA(),D{r+1}/B{r+1})", fmt=S.PCT2)
    S.put(ws, r + 1, 6, f"=IF($B$11=0,NA(),B{r+1}/$B$11)", fmt=S.PCT)
    S.put(ws, r + 1, 7, f"=IF($D$11=0,NA(),D{r+1}/$D$11)", fmt=S.PCT)
    S.put(ws, r + 1, 8, f"=IF(C{r+1}=0,NA(),D{r+1}/C{r+1})", fmt=S.MONEY4)
    S.total_row(ws, r + 2, 1, 8)
    S.put(ws, r + 2, 1, "Portfolio total", bold=True)
    S.put(ws, r + 2, 2, "=$B$11", fmt=S.MONEY, bold=True)
    S.put(ws, r + 2, 3, "=$C$11", fmt=S.INT, bold=True)
    S.put(ws, r + 2, 4, "=$D$11", fmt=S.MONEY, bold=True)
    S.put(ws, r + 2, 5, "=$E$11", fmt=S.PCT2, bold=True)

    r = r + 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(r, 1, "Reading note: Dasani Water earns the highest operating profit per unit of the six brands, while "
                      "Diet Coke and Sprite sit at the bottom. Coca-Cola is the largest brand by sales and by profit, "
                      "but it is one brand inside a six-brand portfolio - no portfolio total in this workbook is "
                      "described as Coca-Cola-brand-only revenue.")
    c.font = Font(size=9, italic=True, color=S.GREY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 34
    ws.freeze_panes = "B5"
    ws.sheet_properties.tabColor = S.BLUE


# ---------------------------------------------------------------- 09_Monthly
def s_monthly(wb):
    ws = wb.create_sheet("09_Monthly_Analysis")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 13, "B": 11, "C": 10, "D": 14, "E": 12, "F": 15, "G": 13, "H": 13, "I": 11,
                  "J": 13, "K": 14, "L": 12})
    S.title(ws, "MONTHLY ANALYSIS  |  one year, twelve months, sorted by date",
            "Scope: all six brands, 2021. No prior-year actuals exist in the dataset, so no year-on-year comparison is "
            "made and no forecast is included in these totals.", last_col=12)
    S.headers(ws, 4, ["Month Start", "Year-Month", "Month", "Sales", "Units", "Operating Profit", "Weighted Margin",
                      "Realised $/Unit", "Sales Share", "MoM Sales Growth", "Margin Change (pp)", "Records"], height=34)
    r = 5
    first = r
    for m in MONTHS:
        S.put(ws, r, 1, m["name"] + "-01", fmt=S.DATE, align="center")
        S.put(ws, r, 2, m["name"], align="left")
        S.put(ws, r, 3, int(m["name"][5:]), fmt="0", align="center")
        S.put(ws, r, 4, f'=SUMIFS(tblSales[Total Sales],tblSales[Year-Month],$B{r})', fmt=S.MONEY)
        S.put(ws, r, 5, f'=SUMIFS(tblSales[Units Sold],tblSales[Year-Month],$B{r})', fmt=S.INT)
        S.put(ws, r, 6, f'=SUMIFS(tblSales[Operating Profit],tblSales[Year-Month],$B{r})', fmt=S.MONEY)
        S.put(ws, r, 7, f"=IF(D{r}=0,NA(),F{r}/D{r})", fmt=S.PCT2)
        S.put(ws, r, 8, f"=IF(E{r}=0,NA(),D{r}/E{r})", fmt=S.MONEY4)
        S.put(ws, r, 9, f"=IF($D$17=0,NA(),D{r}/$D$17)", fmt=S.PCT)
        if r == first:
            S.put(ws, r, 10, "", align="center")
            S.put(ws, r, 11, "", align="center")
        else:
            S.put(ws, r, 10, f"=D{r}/D{r-1}-1", fmt=S.PCT)
            S.put(ws, r, 11, f"=(G{r}-G{r-1})*100", fmt=S.PP)
        S.put(ws, r, 12, f'=COUNTIFS(tblSales[Year-Month],$B{r})', fmt=S.INT)
        r += 1
    S.put(ws, 17, 1, "Total", bold=True, align="left")
    S.put(ws, 17, 4, "=SUM(D5:D16)", fmt=S.MONEY, bold=True)
    S.put(ws, 17, 5, "=SUM(E5:E16)", fmt=S.INT, bold=True)
    S.put(ws, 17, 6, "=SUM(F5:F16)", fmt=S.MONEY, bold=True)
    S.put(ws, 17, 7, "=IF(D17=0,NA(),F17/D17)", fmt=S.PCT2, bold=True)
    S.put(ws, 17, 8, "=IF(E17=0,NA(),D17/E17)", fmt=S.MONEY4, bold=True)
    S.put(ws, 17, 9, "=1", fmt=S.PCT, bold=True)
    S.put(ws, 17, 12, "=SUM(L5:L16)", fmt=S.INT, bold=True)
    S.total_row(ws, 17, 1, 12)

    ws.conditional_formatting.add("G5:G16", ColorScaleRule(
        start_type="min", start_color="FFF8696B", mid_type="percentile", mid_value=50, mid_color="FFFFEB84",
        end_type="max", end_color="FF63BE7B"))
    ws.conditional_formatting.add("D5:D16", DataBarRule(start_type="num", start_value=0, end_type="max", color=S.BLUE))

    r = 19
    S.section(ws, r, "PEAK AND TROUGH  |  live lookups over the monthly block", 12); r += 1
    S.headers(ws, r, ["Measure", "Month", "Value", "How it is calculated"], height=20); r += 1
    stats = [
        ("Highest sales month", f"=INDEX($B$5:$B$16,MATCH(MAX($D$5:$D$16),$D$5:$D$16,0))",
         f"=MAX($D$5:$D$16)", "INDEX/MATCH on monthly sales"),
        ("Lowest sales month", f"=INDEX($B$5:$B$16,MATCH(MIN($D$5:$D$16),$D$5:$D$16,0))",
         f"=MIN($D$5:$D$16)", "INDEX/MATCH on monthly sales"),
        ("Difference (highest - lowest)", "", "=C21-C22", "Absolute dollar gap"),
        ("Highest month vs lowest month", "", "=C21/C22-1", "Relative, denominator = the lowest month"),
        ("Highest units month", f"=INDEX($B$5:$B$16,MATCH(MAX($E$5:$E$16),$E$5:$E$16,0))",
         f"=MAX($E$5:$E$16)", "INDEX/MATCH on units"),
        ("Highest margin month", f"=INDEX($B$5:$B$16,MATCH(MAX($G$5:$G$16),$G$5:$G$16,0))",
         f"=MAX($G$5:$G$16)", "INDEX/MATCH on weighted margin"),
        ("Lowest margin month", f"=INDEX($B$5:$B$16,MATCH(MIN($G$5:$G$16),$G$5:$G$16,0))",
         f"=MIN($G$5:$G$16)", "INDEX/MATCH on weighted margin"),
        ("Margin range (pp)", "", "=(MAX($G$5:$G$16)-MIN($G$5:$G$16))*100", "Highest minus lowest, in pp"),
    ]
    for label, month, value, how in stats:
        S.put(ws, r, 1, label, bold=True)
        S.put(ws, r, 2, month, align="center")
        fmt = S.PCT if "high" in label.lower() and "month vs" in label else (
            S.PCT2 if "margin" in label.lower() and "pp" not in label.lower() else S.MONEY)
        S.put(ws, r, 3, value, fmt=fmt)
        S.put(ws, r, 4, how, align="left")
        r += 1

    r += 1
    S.section(ws, r, "REGION x MONTH  |  sales (USD), for the heatmap and for peak-timing checks", 12); r += 1
    S.headers(ws, r, ["Region"] + [f"2021-{m:02d}" for m in range(1, 13)], height=20)
    r += 1
    hdr = r - 1
    for reg in [x["name"] for x in F["levels"]["region"]]:
        S.put(ws, r, 1, reg, bold=True, align="left")
        for j in range(2, 14):
            col = openpyxl.utils.get_column_letter(j)
            S.put(ws, r, j,
                  f'=SUMIFS(tblSales[Total Sales],tblSales[Region],$A{r},tblSales[Year-Month],{col}${hdr})',
                  fmt=S.MONEY)
        r += 1
    ws.conditional_formatting.add(f"B{hdr+1}:M{r-1}", ColorScaleRule(
        start_type="min", start_color="FFFFFFFF", end_type="max", end_color="FF0070C0"))
    S.note(ws, r + 1, "Heatmap scale: white = lowest sales cell, dark blue = highest. Because only one year is "
                      "observed, these are within-year patterns, not proven recurring seasonality.", last_col=12)
    ws.freeze_panes = "B5"
    ws.sheet_properties.tabColor = S.BLUE


# ---------------------------------------------------------------- 10_Price_Volume
def s_price(wb):
    ws = wb.create_sheet("10_Price_Volume")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 20, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 14, "I": 12,
                  "N": 16, "O": 26, "P": 12, "Q": 11, "R": 12, "S": 10})
    S.title(ws, "PRICE, VOLUME AND PROFITABILITY  |  association, not elasticity",
            "Scope: all six brands, full 2021. A positive relationship here means higher-priced markets also sell "
            "more units; it does not mean price caused volume.", last_col=13)

    S.section(ws, 4, "PRICE-VOLUME ASSOCIATION  |  Pearson correlation, row-level observations", 13)
    S.headers(ws, 5, ["Scope", "n", "Sum x", "Sum y", "Sum xy", "Sum x^2", "Sum y^2", "Correlation r",
                      "Read as"], height=30)
    r = 6
    scopes = [("All records", None)] + [(b, b) for b in [x["name"] for x in F["levels"]["brand"]]]
    for label, brand in scopes:
        S.put(ws, r, 1, label, bold=True, align="left")
        if brand is None:
            S.put(ws, r, 2, "=COUNT(tblSales[Price per Unit])", fmt=S.INT)
            S.put(ws, r, 3, "=SUM(tblSales[Price per Unit])", fmt="0.000")
            S.put(ws, r, 4, "=SUM(tblSales[Units Sold])", fmt=S.INT)
            S.put(ws, r, 5, "=SUMPRODUCT(tblSales[Price per Unit],tblSales[Units Sold])", fmt=S.INT)
            S.put(ws, r, 6, "=SUMPRODUCT(tblSales[Price per Unit]^2)", fmt="0.000")
            S.put(ws, r, 7, "=SUMPRODUCT(tblSales[Units Sold]^2)", fmt="0")
        else:
            S.put(ws, r, 2, f"=COUNTIFS(tblSales[Beverage Brand],$A{r})", fmt=S.INT)
            S.put(ws, r, 3, f"=SUMIFS(tblSales[Price per Unit],tblSales[Beverage Brand],$A{r})", fmt="0.000")
            S.put(ws, r, 4, f"=SUMIFS(tblSales[Units Sold],tblSales[Beverage Brand],$A{r})", fmt=S.INT)
            S.put(ws, r, 5, f"=SUMPRODUCT((tblSales[Beverage Brand]=$A{r})*tblSales[Price per Unit]*tblSales[Units Sold])",
                  fmt=S.INT)
            S.put(ws, r, 6, f"=SUMPRODUCT((tblSales[Beverage Brand]=$A{r})*tblSales[Price per Unit]^2)", fmt="0.000")
            S.put(ws, r, 7, f"=SUMPRODUCT((tblSales[Beverage Brand]=$A{r})*tblSales[Units Sold]^2)", fmt="0")
        S.put(ws, r, 8, f"=(B{r}*E{r}-C{r}*D{r})/SQRT((B{r}*F{r}-C{r}^2)*(B{r}*G{r}-D{r}^2))", fmt="0.000")
        S.put(ws, r, 9, "Association only - confounded by retailer, city and month", align="left", wrap=True, size=8)
        r += 1
    S.note(ws, r, "The all-records row equals =CORREL(tblSales[Price per Unit],tblSales[Units Sold]). "
                  "Brand rows use the equivalent Pearson formula written with SUMPRODUCT so the scope stays explicit.",
           last_col=13)

    r += 2
    S.section(ws, r, "PRICE BAND TABLE  |  unit-weighted outcomes by listed price", 13); r += 1
    S.headers(ws, r, ["Price band", "Lower", "Upper", "Sales", "Units", "Operating Profit", "Weighted Margin",
                      "Realised $/Unit", "Share of sales"], height=30)
    r += 1
    bands = [("<= $0.30", 0, 0.30), ("$0.31-$0.45", 0.30001, 0.45), ("$0.46-$0.60", 0.45001, 0.60),
             ("$0.61-$0.75", 0.60001, 0.75), ("$0.76-$0.90", 0.75001, 0.90), ("> $0.90", 0.90001, 9)]
    first_band = r
    for label, lo, hi in bands:
        S.put(ws, r, 1, label, bold=True, align="left")
        S.put(ws, r, 2, lo, fmt=S.MONEY2)
        S.put(ws, r, 3, hi if hi < 9 else 99, fmt=S.MONEY2)
        S.put(ws, r, 4, f'=SUMIFS(tblSales[Total Sales],tblSales[Price per Unit],">="&B{r},'
                        f'tblSales[Price per Unit],"<="&C{r})', fmt=S.MONEY)
        S.put(ws, r, 5, f'=SUMIFS(tblSales[Units Sold],tblSales[Price per Unit],">="&B{r},'
                        f'tblSales[Price per Unit],"<="&C{r})', fmt=S.INT)
        S.put(ws, r, 6, f'=SUMIFS(tblSales[Operating Profit],tblSales[Price per Unit],">="&B{r},'
                        f'tblSales[Price per Unit],"<="&C{r})', fmt=S.MONEY)
        S.put(ws, r, 7, f"=IF(D{r}=0,NA(),F{r}/D{r})", fmt=S.PCT2)
        S.put(ws, r, 8, f"=IF(E{r}=0,NA(),D{r}/E{r})", fmt=S.MONEY4)
        S.put(ws, r, 9, f"=IF($D${first_band + len(bands)}=0,NA(),D{r}/$D${first_band + len(bands)})", fmt=S.PCT)
        r += 1
    S.put(ws, r, 1, "Total", bold=True)
    S.put(ws, r, 4, f"=SUM(D{first_band}:D{r-1})", fmt=S.MONEY, bold=True)
    S.put(ws, r, 5, f"=SUM(E{first_band}:E{r-1})", fmt=S.INT, bold=True)
    S.put(ws, r, 6, f"=SUM(F{first_band}:F{r-1})", fmt=S.MONEY, bold=True)
    S.put(ws, r, 7, f"=IF(D{r}=0,NA(),F{r}/D{r})", fmt=S.PCT2, bold=True)
    S.put(ws, r, 8, f"=IF(E{r}=0,NA(),D{r}/E{r})", fmt=S.MONEY4, bold=True)
    S.total_row(ws, r, 1, 9)

    # scatter source block, columns N:S
    hdr = 4
    S.headers(ws, hdr, ["Beverage Brand", "Location", "Sales", "Units", "Realised $/Unit", "Weighted Margin"],
              start=14, height=30)
    rr = hdr + 1
    # real brand-location pairs, taken from the source data
    real = set()
    src = openpyxl.load_workbook(SOURCE, data_only=True)["Data"]
    for row in range(6, 3894):
        real.add((src.cell(row, 8).value, src.cell(row, 6).value + " | " + src.cell(row, 7).value))
    pairs = [p for p in real]
    pairs.sort()
    for b, loc in pairs:
        S.put(ws, rr, 14, b, align="left")
        S.put(ws, rr, 15, loc, align="left")
        S.put(ws, rr, 16, f'=SUMIFS(tblSales[Total Sales],tblSales[Beverage Brand],$N{rr},tblSales[Location],$O{rr})',
              fmt=S.MONEY)
        S.put(ws, rr, 17, f'=SUMIFS(tblSales[Units Sold],tblSales[Beverage Brand],$N{rr},tblSales[Location],$O{rr})',
              fmt=S.INT)
        S.put(ws, rr, 18, f"=IF(Q{rr}=0,NA(),P{rr}/Q{rr})", fmt=S.MONEY4)
        S.put(ws, rr, 19, f'=IF(P{rr}=0,NA(),SUMIFS(tblSales[Operating Profit],tblSales[Beverage Brand],$N{rr},'
                          f'tblSales[Location],$O{rr})/P{rr})', fmt=S.PCT2)
        rr += 1
    last = rr - 1

    ch = ScatterChart()
    ch.title = "Realised price vs units sold - 324 city x brand cells"
    ch.style = 13
    ch.height = 9.5
    ch.width = 17
    ch.x_axis.title = "Realised sales per unit (USD)"
    ch.y_axis.title = "Units sold in 2021"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend = None
    xref = openpyxl.chart.Reference(ws, min_col=18, min_row=hdr + 1, max_row=last)
    yref = openpyxl.chart.Reference(ws, min_col=17, min_row=hdr, max_row=last)
    ser = Series(yref, xref, title="City x brand")
    ser.marker = Marker(symbol="circle", size=5)
    ser.trendline = Trendline(trendlineType="linear")
    ch.series.append(ser)
    ws.add_chart(ch, "A32")

    S.note(ws, 30, "Scatter source: every brand-city pair, aggregated over the 12 months. Aggregation hides "
                   "within-year price changes, which is why the correlation below is also computed at row level.",
           last_col=13)
    ws.sheet_properties.tabColor = S.BLUE
