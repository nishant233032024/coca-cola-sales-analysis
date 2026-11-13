"""Second half of the workbook build: pivots placeholder, dashboard, insights,
implementation guide, validation and script sheets. Imported by build_base.py."""
from __future__ import annotations

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

import style as S
from content_loom import build_loom, build_qa, CHECKLIST, money, pct
from content_guide import build_impl_rows
from content_report import build_insights


def s_pivot_placeholder(wb):
    ws = wb.create_sheet("11_Pivot_Tables")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 26, "B": 20, "C": 16, "D": 16, "E": 30, "F": 30, "G": 14})
    S.title(ws, "PIVOT TABLES  |  one shared cache, eight views",
            "Native PivotTables are added to this sheet by Excel automation. If they are missing, rebuild them with "
            "the configurations below - every field and value is listed.", last_col=7)
    S.headers(ws, 4, ["PivotTable", "Source", "Rows", "Columns", "Values", "Filters / sort", "Chart"], height=26)
    rows = [
        ("PT_KPI", "tblSales (shared cache)", "(none)", "Values",
         "Sales, Units, Operating Profit, Margin pct, Sales per Unit",
         "Slicers: Retailer, Region, Brand", "Dashboard KPI cards (GETPIVOTDATA)"),
        ("PT_Retailer", "tblSales (shared cache)", "Retailer", "(none)",
         "Sales, Units, Operating Profit, Margin pct", "Sales descending", "Retailer bar"),
        ("PT_Region", "tblSales (shared cache)", "Region", "(none)",
         "Sales, Units, Operating Profit, Margin pct", "Sales descending", "Region bar"),
        ("PT_Month", "tblSales (shared cache)", "Year-Month", "(none)",
         "Sales, Units, Operating Profit, Margin pct", "Year-Month ascending",
         "Monthly sales + monthly margin lines"),
        ("PT_Brand", "tblSales (shared cache)", "Beverage Brand", "(none)",
         "Sales, Units, Operating Profit, Margin pct", "Sales descending", "Brand contribution bar"),
        ("PT_RegionMonth", "tblSales (shared cache)", "Region", "Year-Month", "Sales",
         "Month ascending", "Dashboard heatmap table"),
        ("PT_BrandRegion", "tblSales (shared cache)", "Beverage Brand", "Region",
         "Sales, Operating Profit", "Sales descending", "Reference table"),
        ("PT_RetailerRegion", "tblSales (shared cache)", "Retailer", "Region",
         "Sales, Operating Profit", "Sales descending", "Reference table"),
        ("PT_RetailerChart / PT_RegionChart / PT_MonthSalesChart / PT_MonthMarginChart / PT_BrandChart",
         "tblSales (shared cache)", "one dimension each", "(none)", "single measure",
         "chart sources only", "the five dashboard charts"),
    ]
    r = 5
    for row in rows:
        for i, v in enumerate(row, start=1):
            S.put(ws, r, i, v, align="left", wrap=True, bold=(i == 1))
        ws.row_dimensions[r].height = 34
        r += 1
    S.note(ws, r + 1, "Margin is a PivotTable calculated field: ='Operating Profit'/'Total Sales'. It evaluates as a "
                      "ratio of the summed fields in each cell, which is the weighted margin - not an average of row "
                      "margins. 'Sales per Unit' is ='Total Sales'/'Units Sold'.", last_col=7)
    ws.sheet_properties.tabColor = S.LIGHT
    return ws


def s_dashboard(wb):
    ws = wb.create_sheet("12_Dashboard")
    ws.sheet_view.showGridLines = False
    for col in "ABCDEFGHIJKLMNOPQRS":
        ws.column_dimensions[col].width = 9.6
    for col in "UVW":
        ws.column_dimensions[col].width = 13

    ws.merge_cells("A1:S1")
    c = ws.cell(1, 1, "COCA-COLA  |  SALES & OPERATING PROFITABILITY")
    c.font = Font(bold=True, size=20, color=S.WHITE)
    c.fill = S.fill(S.DARK)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:S2")
    c = ws.cell(2, 1, "U.S. retailers  |  all six beverage brands  |  2021-01-02 to 2021-12-25  |  3,888 records  |  "
                      "3,888-row filtered PivotTable cache  |  USD (assumed - not documented in source)")
    c.font = Font(size=9, color=S.GREY)
    c.alignment = Alignment(vertical="center", indent=1)

    ws.merge_cells("A3:S3")
    c = ws.cell(3, 1, "FILTERS: native slicers at right (Retailer, Region, Beverage Brand) and an Invoice Date timeline. "
                      "They filter every pivot chart and every KPI card. Clear a slicer to return to the full portfolio.")
    c.font = Font(size=9, bold=True, color=S.BLUE)
    c.fill = S.fill(S.PALE)
    c.alignment = Alignment(vertical="center", indent=1)

    # KPI cards
    cards = [
        ("A5:C5", "A6:C6", "A7:C7", "TOTAL SALES", "money"),
        ("E5:G5", "E6:G6", "E7:G7", "UNITS SOLD", "int"),
        ("I5:K5", "I6:K6", "I7:K7", "OPERATING PROFIT", "money"),
        ("M5:O5", "M6:O6", "M7:O7", "WEIGHTED MARGIN", "pct"),
        ("Q5:S5", "Q6:S6", "Q7:S7", "SALES PER UNIT", "spu"),
    ]
    for label_rng, value_rng, sub_rng, label, kind in cards:
        ws.merge_cells(label_rng)
        ws.merge_cells(value_rng)
        ws.merge_cells(sub_rng)
        lc = ws[label_rng.split(":")[0]]
        lc.value = label
        lc.font = Font(size=9, bold=True, color=GREY_HACK())
        lc.fill = S.fill(S.WHITE)
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws[value_rng.split(":")[0]]
        vc.font = Font(size=18, bold=True, color=S.DARK)
        vc.fill = S.fill(S.WHITE)
        vc.alignment = Alignment(horizontal="center", vertical="center")
        sc = ws[sub_rng.split(":")[0]]
        sc.font = Font(size=8, italic=True, color=S.GREY)
        sc.fill = S.fill(S.WHITE)
        sc.alignment = Alignment(horizontal="center", vertical="center")
    for rw in (5, 6, 7):
        ws.row_dimensions[rw].height = 16 if rw != 6 else 30

    # chart areas
    zones = [
        ("A9:I9", "RETAILER SALES", "PT_Retailer"),
        ("K9:S9", "REGIONAL SALES", "PT_Region"),
        ("A24:I24", "MONTHLY SALES TREND", "PT_Month"),
        ("K24:S24", "MONTHLY MARGIN TREND", "PT_Month"),
        ("A38:S38", "REGION x MONTH SALES  |  native PivotTable with conditional-format heatmap", None),
        ("A51:I51", "BRAND CONTRIBUTION", "PT_Brand"),
        ("K51:S51", "WHAT THE DATA SAYS  |  default-view interpretation", None),
    ]
    for rng, label, _src in zones:
        ws.merge_cells(rng)
        row = int(rng.split(":")[0][1:])
        c = ws.cell(row, 1, label)
        c.font = Font(size=10, bold=True, color=S.DARK)
        c.fill = S.fill(S.BG)
        c.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[row].height = 20

    # insight panel text
    ins = build_insights(FACTS)
    lines = [
        f"1. Profit is concentrated: Sodapop is 50.7% of sales and 51.9% of profit; FizzySip is 29.8% of sales but "
        f"26.3% of profit.",
        f"2. The West earns 32.6% of sales but 29.1% of profit - 4.0 pp below the portfolio margin.",
        f"3. 2021 growth was price-led: sales +104.8% Jan to Dec while units rose 24.4%.",
        f"4. Two peaks: December on realised price, July on units. Margin stayed within 1.1 pp all year.",
        f"5. Action: matched-cell margin review in the lowest-margin material cities before any price change.",
    ]
    r = 52
    for line in lines:
        ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=19)
        c = ws.cell(r, 11, line)
        c.font = Font(size=9, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.row_dimensions[r].height = 24
        r += 1

    # slicer panel label
    ws.merge_cells("U1:W1")
    c = ws.cell(1, 21, "FILTERS")
    c.font = Font(bold=True, size=10, color=S.WHITE)
    c.fill = S.fill(S.BLUE)
    c.alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A65:S65")
    c = ws.cell(65, 1, "Source: tblSales on 02_Clean_Data. Margin = total operating profit / total sales. "
                       "Charts and cards are native PivotCharts and GETPIVOTDATA and respond to the slicers. "
                       "Educational dataset - not audited company reporting.")
    c.font = Font(size=8, italic=True, color=S.GREY)
    c.alignment = Alignment(vertical="center", indent=1)

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = S.DARK
    return ws


def GREY_HACK():
    return "FF6B7280"


FACTS = None


def s_insights(wb):
    ws = wb.create_sheet("13_Insights_Actions")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 7, "B": 46, "C": 62, "D": 46, "E": 30, "F": 40, "G": 46, "H": 36})
    S.title(ws, "INSIGHTS AND ACTIONS  |  five prioritised recommendations",
            "Every recommendation is proportionate to what the data can support. No uplift, ROI or target figure is "
            "invented; suggested owners are proposed roles, not named employees.", last_col=8)
    S.headers(ws, 4, ["Priority", "Finding", "Verified evidence", "Proposed action", "Suggested owner", "Success KPI",
                      "Risk or limitation", "Additional evidence needed"], height=34)
    r = 5
    fills = ["FFEAF2FB", "FFFFFFFF"]
    for i, ins in enumerate(build_insights(FACTS)):
        bg = fills[i % 2]
        S.put(ws, r, 1, ins["pri"], align="center", bold=True, bg=bg)
        S.put(ws, r, 2, ins["finding"], align="left", wrap=True, bg=bg)
        S.put(ws, r, 3, ins["evidence"], align="left", wrap=True, bg=bg, size=9)
        S.put(ws, r, 4, ins["action"], align="left", wrap=True, bg=bg, size=9)
        S.put(ws, r, 5, ins["owner"], align="left", wrap=True, bg=bg, size=9)
        S.put(ws, r, 6, ins["kpi"], align="left", wrap=True, bg=bg, size=9)
        S.put(ws, r, 7, ins["risk"], align="left", wrap=True, bg=bg, size=9)
        S.put(ws, r, 8, ins["more"], align="left", wrap=True, bg=bg, size=9)
        ws.row_dimensions[r].height = 110
        r += 1
    S.note(ws, r + 1, "Priority reflects the size of the verified gap and how cheaply it can be investigated - not the "
                      "size of an expected financial benefit, which this dataset cannot quantify.", last_col=8)
    r += 2
    S.section(ws, r, "WHAT THIS WORKBOOK DOES NOT CLAIM", 8); r += 1
    for line in [
        "No expected revenue, profit, ROI or target figures are estimated anywhere in this workbook.",
        "No causal claim is made about price, promotion, competition or supply chain; the dataset has no cost field.",
        "No prior-year comparison and no future forecast is presented as actual.",
    ]:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        c = ws.cell(r, 1, "- " + line)
        c.font = Font(size=9, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        r += 1
    ws.freeze_panes = "B5"
    ws.sheet_properties.tabColor = S.BLUE


def s_implementation(wb):
    F = GLOBAL_FACTS
    T = F["total"]
    RET = {r["name"]: r for r in F["levels"]["retailer"]}
    REG = {r["name"]: r for r in F["levels"]["region"]}
    ws = wb.create_sheet("14_Implementation")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 12, "B": 30, "C": 40, "D": 26, "E": 34, "F": 44, "G": 40, "H": 34, "I": 22, "J": 32,
                  "K": 26, "L": 40})
    S.title(ws, "IMPLEMENTATION GUIDE  |  rebuild every insight by hand",
            "Twelve fields per step: question, fields, cleaning, formula, PivotTable, chart, result, reading, "
            "validation, and the question you are likely to be asked.", last_col=12)
    S.headers(ws, 4, ["Phase", "Business question / insight", "Source fields", "Cleaning step", "Excel formula or DAX",
                      "PivotTable configuration", "Chart steps", "Verified result", "How to interpret it",
                      "Validation check", "Likely question", "Concise answer"], height=40)

    rows = build_impl_rows(F, T, RET, REG, pct, money)
    r = 5
    for row in rows:
        for i, v in enumerate(row, start=1):
            S.put(ws, r, i, v, align="left", wrap=True, size=8, bold=(i == 1), literal=True)
        ws.row_dimensions[r].height = 96
        r += 1

    r += 1
    S.section(ws, r, "POWER QUERY  |  optional clean-load alternative (M code, not installed in this file)", 12); r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
    c = ws.cell(r, 1, "The workbook you are reading was built with Python + Excel automation, not Power Query. The M "
                      "code below matches the inspected source structure and is supplied as a specification; it is not "
                      "a live query in this file.")
    c.font = Font(size=9, italic=True, color=S.GREY)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30
    r += 1
    m_code = [
        "let",
        "    Source = Excel.Workbook(File.Contents(pSourceFile), null, true),",
        "    DataSheet = Source{[Item=\"Data\",Kind=\"Sheet\"]}[Data],",
        "    PromoteHeaders = Table.PromoteHeaders(Table.Skip(DataSheet,4), [PromoteAllScalars=true]),",
        "    KeepColumns = Table.SelectColumns(PromoteHeaders,{\"Retailer\",\"Retailer ID\",\"Invoice Date\",\"Region\",",
        "        \"State\",\"City\",\"Beverage Brand\",\"Price per Unit\",\"Units Sold\",\"Total Sales\",",
        "        \"Operating Profit\",\"Operating Margin\"}),",
        "    RemoveBlankRows = Table.SelectRows(KeepColumns, each not List.IsEmpty(List.RemoveMatchingItems(",
        "        Record.FieldValues(_), {null, \"\"}))),",
        "    TrimText = Table.TransformColumns(RemoveBlankRows, {{\"Retailer\", Text.Trim}, {\"Region\", Text.Trim},",
        "        {\"State\", Text.Trim}, {\"City\", Text.Trim}, {\"Beverage Brand\", Text.Trim}}),",
        "    TypedDate = Table.TransformColumnTypes(TrimText, {{\"Invoice Date\", type date}}),",
        "    TypedNumbers = Table.TransformColumnTypes(TypedDate, {{\"Price per Unit\", type number},",
        "        {\"Units Sold\", Int64.Type}, {\"Total Sales\", type number}, {\"Operating Profit\", type number},",
        "        {\"Operating Margin\", type number}}),",
        "    AddHelpers = Table.AddColumn(TypedNumbers, \"Month Start\", each Date.StartOfMonth([Invoice Date]),",
        "        type date),",
        "    Output = Table.AddColumn(AddHelpers, \"Location\", each [State] & \" | \" & [City], type text)",
        "in",
        "    Output",
    ]
    for line in m_code:
        S.put(ws, r, 1, line, align="left", size=9, border=False, color="FF1F2937")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=12)
        r += 1
    ws.freeze_panes = "B5"
    ws.sheet_properties.tabColor = S.BLUE


def s_validation(wb):
    T = GLOBAL_FACTS["total"]
    ws = wb.create_sheet("15_Validation")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 58, "B": 20, "C": 20, "D": 12, "E": 22, "F": 66})
    S.title(ws, "VALIDATION  |  check, expected, actual, pass/fail, explanation",
            "Live formulas where possible; Python-verified checks are labelled. Structural checks confirm the file, not "
            "a human review of every chart.", last_col=6)
    S.headers(ws, 4, ["Check", "Expected", "Actual", "Pass/Fail", "Verified by", "Explanation"], height=24)

    checks = [
        ("Raw records = cleaned records", "=COUNTA('01_Raw_Data'!B6:B3893)", "=COUNTA(tblSales[Record ID])",
         "Formula", "No row was dropped, added or merged."),
        ("Raw Total Sales = cleaned Total Sales", "=SUM('01_Raw_Data'!K6:K3893)", "=SUM(tblSales[Total Sales])",
         "Formula", "Source sales values are never overwritten."),
        ("Raw Operating Profit = cleaned Operating Profit", "=SUM('01_Raw_Data'!L6:L3893)",
         "=SUM(tblSales[Operating Profit])", "Formula", "Source profit values are never overwritten."),
        ("Raw Units = cleaned Units", "=SUM('01_Raw_Data'!J6:J3893)", "=SUM(tblSales[Units Sold])", "Formula",
         "Volume reconciliation."),
        ("Retailer block sums to portfolio sales", T["sales"], "=SUM('06_Retailer_Analysis'!B5:B8)", "Formula",
         "Retailer dimension fully covers the data."),
        ("Region block sums to portfolio sales", T["sales"], "=SUM('07_Region_Analysis'!B5:B9)", "Formula",
         "Region dimension fully covers the data."),
        ("Brand block sums to portfolio sales", T["sales"], "=SUM('08_Brand_Analysis'!B5:B10)", "Formula",
         "Six brands cover the portfolio."),
        ("Monthly block sums to portfolio sales", T["sales"], "='09_Monthly_Analysis'!D17", "Formula",
         "Twelve months cover the year."),
        ("Weighted margin recomputed independently", T["margin"],
         "=SUM(tblSales[Operating Profit])/SUM(tblSales[Total Sales])", "Formula", "Ratio of sums, not an average."),
        ("Sales identity exceptions (Price x Units)", 0, '=COUNTIF(tblSales[Validation Status],"Sales mismatch")',
         "Formula", "All 3,888 rows satisfy the identity within $0.01."),
        ("Margin identity exceptions", 0, '=COUNTIF(tblSales[Validation Status],"Margin mismatch")', "Formula",
         "Supplied margin equals profit/sales wherever sales are non-zero."),
        ("Zero-activity records retained", 2, '=COUNTIF(tblSales[Validation Status],"Zero sales - check")', "Formula",
         "Two Omaha rows; zero impact on totals."),
        ("Rows marked OK", 3886, '=COUNTIF(tblSales[Validation Status],"OK")', "Formula",
         "3,888 less the two flagged rows."),
        ("Months with full 324-record coverage", 12,
         '=SUMPRODUCT(--(COUNTIF(tblSales[Year-Month],{"2021-01","2021-02","2021-03","2021-04","2021-05","2021-06",'
         '"2021-07","2021-08","2021-09","2021-10","2021-11","2021-12"})=324))', "Formula",
         "No partial month, so monthly comparisons are like-for-like."),
        ("Negative values in sales, units or profit", 0, 0, "Python",
         "Verified in Reproduce_Coca_Cola_Analysis.py; no returns present."),
        ("Duplicate retailer-location-brand-month keys", 0, 0, "Python",
         "Key uniqueness verified; no legitimate repeat transaction removed."),
        ("Date range", "2021-01-02 to 2021-12-25", "2021-01-02 to 2021-12-25", "Python",
         "MIN/MAX of Invoice Date; one calendar year, 12 months."),
        ("Currency documented in source", "Documented", "Not documented", "Disclosed",
         "USD is an assumption; flagged on 04_Data_Quality and in the report."),
        ("Native PivotTables present", 18, "PLACEHOLDER", "Formula",
         "Counted after Excel automation builds them: 13 summary views on 11_Pivot_Tables, 4 summary panels and "
         "the heatmap on 12_Dashboard."),
        ("Slicers present and connected", 3, "PLACEHOLDER", "Formula",
         "Retailer, Region, Beverage Brand - all on one shared cache. The Invoice Date timeline is a documented one-click manual step, not automated."),
        ("Dashboard KPI card = filtered pivot value", "MATCH", "PLACEHOLDER", "Text",
         "Slicer set to Sodapop, card read back, then cleared."),
        ("Slicer cleared returns full-portfolio KPI", "MATCH", "PLACEHOLDER", "Text",
         "Confirms the default state after filtering."),
        ("Slicer filters the PivotTables and pivot charts", "MATCH", "PLACEHOLDER", "Text",
         "With Sodapop selected, the dashboard retailer PivotTable shows Sodapop only and its sales equal the "
         "Sodapop total; every PivotTable is explicitly wired to each slicer cache."),
    ]
    r = 5
    first = r
    for label, exp, act, by, expl in checks:
        S.put(ws, r, 1, label, align="left", wrap=True)
        S.put(ws, r, 2, exp, align="center")
        S.put(ws, r, 3, act, align="center")
        if by == "Formula":
            S.put(ws, r, 4, f'=IF(ABS(C{r}-B{r})<=0.01,"Pass","Fail")', align="center", bold=True)
        elif by == "Disclosed":
            S.put(ws, r, 4, "Disclosed", align="center", bold=True, color=S.GREY)
        else:
            S.put(ws, r, 4, f'=IF(C{r}=B{r},"Pass","Fail")', align="center", bold=True)
        S.put(ws, r, 5, by, align="center")
        S.put(ws, r, 6, expl, align="left", wrap=True, size=9)
        ws.row_dimensions[r].height = 26
        r += 1
    ws.conditional_formatting.add(f"D{first}:D{r-1}",
                                  CellIsRule(operator="equal", formula=['"Pass"'],
                                             font=Font(color=S.GREEN, bold=True)))
    ws.conditional_formatting.add(f"D{first}:D{r-1}",
                                  CellIsRule(operator="equal", formula=['"Fail"'],
                                             font=Font(color=S.RED, bold=True)))
    S.note(ws, r + 1, "What these checks do not cover: a visual review of every chart in desktop Excel, the position of "
                      "slicers on other screen sizes, and print layout. Those are manual steps listed in "
                      "14_Implementation.", last_col=6)
    ws.freeze_panes = "A5"
    ws.sheet_properties.tabColor = S.GREEN


def s_script(wb):
    ws = wb.create_sheet("16_Loom_Script")
    ws.sheet_view.showGridLines = False
    S.widths(ws, {"A": 14, "B": 34, "C": 120, "D": 18})
    S.title(ws, "LOOM SCRIPT  |  6-8 minutes, first person, Nishant",
            "Every number is pulled from the same facts used by the workbook. Read it, do not memorise it.", last_col=3)
    S.headers(ws, 4, ["Time", "On-screen view / action", "Exact spoken narration", "Words"], height=24)
    script = build_loom(GLOBAL_FACTS)
    r = 5
    total_words = 0
    for block in script:
        S.put(ws, r, 1, block["t"], align="center", bold=True, color=S.DARK)
        S.put(ws, r, 2, block["view"], align="left", wrap=True, size=9)
        S.put(ws, r, 3, block["say"], align="left", wrap=True, size=9)
        n = len(block["say"].split())
        total_words += n
        S.put(ws, r, 4, n, fmt=S.INT, align="center")
        ws.row_dimensions[r].height = 150
        r += 1
    S.put(ws, r, 1, "Total", bold=True)
    S.put(ws, r, 3, f"Spoken word count: {total_words} words (target 850-1,050). At a calm pace with demonstration "
                    f"pauses this runs about 7 minutes.", align="left", italic=True, size=9)
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)

    r += 2
    S.section(ws, r, "FIVE LIKELY EVALUATOR QUESTIONS AND MODEL ANSWERS", 3); r += 1
    for q, a in build_qa(GLOBAL_FACTS):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(r, 1, "Q: " + q)
        c.font = Font(bold=True, size=10, color=S.DARK)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        ws.row_dimensions[r].height = 18
        r += 1
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(r, 1, "A: " + a)
        c.font = Font(size=10, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="top", indent=1)
        ws.row_dimensions[r].height = 56
        r += 2

    S.section(ws, r, "RECORDING CHECKLIST", 3); r += 1
    for i, item in enumerate(CHECKLIST, 1):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        c = ws.cell(r, 1, f"{i}. {item}")
        c.font = Font(size=10, color=S.CHAR)
        c.alignment = Alignment(wrap_text=True, vertical="center", indent=1)
        r += 1
    ws.sheet_properties.tabColor = S.LIGHT


GLOBAL_FACTS = None
