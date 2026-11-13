"""Create the Word/PDF deliverables from the same facts as the workbook."""
from __future__ import annotations

import json
import os

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

import content_guide
from content_loom import build_loom, build_qa, CHECKLIST, money, pct
from content_report import build_insights, build_report

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
OUT = os.path.join(ROOT, "deliverables")

F = json.load(open(os.path.join(HERE, "..", "facts.json"), encoding="utf-8"))
T = F["total"]
RET = {r["name"]: r for r in F["levels"]["retailer"]}
REG = {r["name"]: r for r in F["levels"]["region"]}
BRD = {r["name"]: r for r in F["levels"]["brand"]}

BLUE = RGBColor(0x00, 0x70, 0xC0)
DARK = RGBColor(0x17, 0x36, 0x5D)
GREY = RGBColor(0x6B, 0x72, 0x80)


def new_doc(landscape=False):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Calibri"
    st.font.size = Pt(10.5)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.line_spacing = 1.12
    for s in doc.sections:
        if landscape:
            s.orientation = WD_ORIENT.LANDSCAPE
            s.page_width, s.page_height = s.page_height, s.page_width
        s.top_margin = Cm(1.8)
        s.bottom_margin = Cm(1.8)
        s.left_margin = Cm(2.0)
        s.right_margin = Cm(2.0)
    return doc


def h1(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = DARK
    p.paragraph_format.space_after = Pt(2)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = BLUE
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    return p


def h3(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = DARK
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    return p


def body(doc, text, italic=False, size=10.5, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic = italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def label_para(doc, label, text):
    p = doc.add_paragraph()
    r = p.add_run(label + " ")
    r.bold = True
    r.font.size = Pt(10)
    r.font.color.rgb = DARK
    r2 = p.add_run(text)
    r2.font.size = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    return p


def kv_table(doc, rows, widths=None, header=True):
    t = doc.add_table(rows=0, cols=len(rows[0]))
    t.style = "Light Grid Accent 1"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for j, v in enumerate(row):
            cells[j].text = ""
            p = cells[j].paragraphs[0]
            r = p.add_run(str(v))
            r.font.size = Pt(9)
            if i == 0 and header:
                r.bold = True
    if widths:
        for row in t.rows:
            for j, w in enumerate(widths):
                row.cells[j].width = Cm(w)
    return t


def footer_note(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.size = Pt(8)
    r.italic = True
    r.font.color.rgb = GREY


# ------------------------------------------------------------------ report
def build_report_doc():
    doc = new_doc()
    h1(doc, "Coca-Cola Sales Analysis Using Excel")
    p = doc.add_paragraph()
    r = p.add_run("U.S. retailers, six beverage brands, calendar 2021  |  Prepared by Nishant")
    r.font.size = Pt(10)
    r.font.color.rgb = GREY

    h2(doc, "At a glance")
    rows = [["Metric", "Value", "Scope"],
            ["Total sales", money(T["sales"]), "All six brands, 3,888 records, 2021"],
            ["Units sold", f"{T['units']:,}", "All six brands"],
            ["Operating profit", money(T["profit"]), "Operating profit only - no net profit exists in the data"],
            ["Weighted operating margin", pct(T["margin"]), "Total operating profit / total sales"],
            ["Realised sales per unit", money(T["sales_per_unit"], 2), "Total sales / total units"],
            ["Operating profit per unit", money(T["profit_per_unit"], 3), "Total operating profit / total units"],
            ["Retailers / regions / cities / brands", "4 / 5 / 54 / 6", "As supplied"],
            ["Period", f"{F['date_min']} to {F['date_max']}", "One year, 12 full months"]]
    kv_table(doc, rows, widths=[6.2, 4.0, 6.4])

    for heading, paras in build_report(F):
        h2(doc, heading)
        for text in paras:
            body(doc, text)

    h2(doc, "Document control")
    body(doc, "This report describes an educational assignment dataset and does not represent audited Coca-Cola "
              "company reporting. All figures are calculated from the supplied workbook; the live reconciliation is on "
              "worksheet 15_Validation of Coca_Cola_Sales_Analysis.xlsx. Currency is assumed to be USD because the "
              "source file documents no currency.", size=9, italic=True, color=GREY)
    path = os.path.join(OUT, "Coca_Cola_Project_Report.docx")
    doc.save(path)
    return path


# ------------------------------------------------------------------ guide
def build_guide_doc():
    doc = new_doc(landscape=True)
    h1(doc, "Excel Implementation Guide")
    body(doc, "Rebuild every view in Coca_Cola_Sales_Analysis.xlsx from the supplied source file. Each step lists the "
              "business question, the fields used, the cleaning step, the exact formula, the PivotTable configuration, "
              "the chart steps, the verified result, how to read it, the validation check, and the question you are "
              "likely to be asked.", size=10)

    h2(doc, "Environment and prerequisites")
    for line in [
        "Microsoft Excel 2016 or later (Microsoft 365 recommended). PivotTables, slicers and GETPIVOTDATA are native "
        "features; nothing here needs a macro.",
        "The source sheet is 'Data': headers on row 5, records on rows 6-3893, fields B:M.",
        "Create the analysis table on a new sheet, name it tblSales, and keep the original sheet unchanged.",
        "Before presenting: Data > Refresh All, then Ctrl+Alt+F9, then read 15_Validation.",
    ]:
        p = doc.add_paragraph(line, style="List Bullet")
        p.runs[0].font.size = Pt(9.5)

    steps = content_guide.build_impl_rows(F, T, RET, REG, pct, money)
    for row in steps:
        phase, question, fields, cleaning, formula, pivot, chart, result, reading, validation, lq, la = row
        h3(doc, f"{phase}  |  {question}")
        label_para(doc, "Source fields:", fields)
        label_para(doc, "Cleaning:", cleaning)
        label_para(doc, "Formula / measure:", formula)
        label_para(doc, "PivotTable:", pivot)
        label_para(doc, "Chart:", chart)
        label_para(doc, "Verified result:", result)
        label_para(doc, "How to interpret it:", reading)
        label_para(doc, "Validation check:", validation)
        label_para(doc, "Likely question:", lq)
        label_para(doc, "Concise answer:", la)

    h2(doc, "Slicers and the shared cache")
    for line in [
        "Click any cell in PT_KPI > PivotTable Analyze > Insert Slicer > Retailer. Repeat for Region and Beverage Brand.",
        "Because every PivotTable is built from the same cache, one slicer set filters all of them. Confirm this by "
        "clicking a retailer and watching the KPI cards and every chart move together.",
        "A date timeline is not included in the delivered build: select a cell in a PivotTable > PivotTable Analyze > "
        "Insert Timeline > Invoice Date. It appears as a draggable date range control.",
        "Clear a slicer with the funnel icon in its header to return to the full portfolio.",
    ]:
        p = doc.add_paragraph(line, style="List Bullet")
        p.runs[0].font.size = Pt(9.5)

    h2(doc, "Power Query alternative (optional, not installed in the delivered file)")
    body(doc, "The delivered workbook was built with Python plus Excel automation, not Power Query. The M code below "
              "matches the inspected structure and is a specification you can paste into a blank query after creating "
              "a text parameter pSourceFile with the full path to the original file.", size=9.5, italic=True)
    m_code = """let
    Source = Excel.Workbook(File.Contents(pSourceFile), null, true),
    DataSheet = Source{[Item="Data",Kind="Sheet"]}[Data],
    PromoteHeaders = Table.PromoteHeaders(Table.Skip(DataSheet,4), [PromoteAllScalars=true]),
    KeepColumns = Table.SelectColumns(PromoteHeaders,{"Retailer","Retailer ID","Invoice Date","Region",
        "State","City","Beverage Brand","Price per Unit","Units Sold","Total Sales",
        "Operating Profit","Operating Margin"}),
    TypedDate = Table.TransformColumnTypes(KeepColumns,{{"Invoice Date", type date}}),
    TypedNumbers = Table.TransformColumnTypes(TypedDate,{{"Price per Unit", type number},
        {"Units Sold", Int64.Type},{"Total Sales", type number},{"Operating Profit", type number},
        {"Operating Margin", type number}}),
    AddMonth = Table.AddColumn(TypedNumbers, "Month Start", each Date.StartOfMonth([Invoice Date]), type date),
    AddLocation = Table.AddColumn(AddMonth, "Location", each [State] & " | " & [City], type text)
in
    AddLocation"""
    p = doc.add_paragraph()
    r = p.add_run(m_code)
    r.font.name = "Consolas"
    r.font.size = Pt(8)

    h2(doc, "Manual steps that remain after the automated build")
    for line in [
        "Add the Invoice Date timeline (three clicks, described above) if you want date filtering on camera.",
        "Check the dashboard on your own screen at 100% zoom; slicer positions are fixed and may need a nudge on very "
        "small displays.",
        "Print or export the dashboard with File > Print > Landscape, fit to one page wide (print area A1:S94, "
        "which includes the four summary PivotTables below the charts).",
    ]:
        p = doc.add_paragraph(line, style="List Bullet")
        p.runs[0].font.size = Pt(9.5)

    path = os.path.join(OUT, "Coca_Cola_Excel_Implementation_Guide.docx")
    doc.save(path)
    return path


# ------------------------------------------------------------------ script
def build_script_doc():
    doc = new_doc(landscape=True)
    h1(doc, "Loom Presentation Script")
    body(doc, "Presenter: Nishant  |  Target length: 6-8 minutes  |  Prepare the workbook in its default state "
              "(all slicers cleared) before recording.", size=10, color=GREY)

    script = build_loom(F)
    words = sum(len(b["say"].split()) for b in script)
    body(doc, f"Spoken word count: {words} words. At a calm pace with demonstration pauses this runs about seven "
              f"minutes.", size=10, italic=True)

    t = doc.add_table(rows=1, cols=3)
    t.style = "Light Grid Accent 1"
    hdr = t.rows[0].cells
    for i, lab in enumerate(["Time", "On-screen view / action", "Exact spoken narration"]):
        hdr[i].text = ""
        r = hdr[i].paragraphs[0].add_run(lab)
        r.bold = True
        r.font.size = Pt(9)
    for b in script:
        cells = t.add_row().cells
        cells[0].text = b["t"]
        cells[1].text = b["view"]
        cells[2].text = b["say"]
        for c in cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(8.5)
    for row in t.rows:
        row.cells[0].width = Cm(3.2)
        row.cells[1].width = Cm(7.0)
        row.cells[2].width = Cm(15.0)

    h2(doc, "Five likely evaluator questions and model answers")
    for q, a in build_qa(F):
        h3(doc, q)
        body(doc, a, size=9.5)

    h2(doc, "Recording checklist")
    for item in CHECKLIST:
        p = doc.add_paragraph(item, style="List Bullet")
        p.runs[0].font.size = Pt(9.5)

    h2(doc, "Before you press record")
    body(doc, "Open 00_ReadMe and 12_Dashboard in separate windows. Set Excel zoom to 110-120% so the axis labels are "
              "readable, turn off the formula bar and gridlines, and keep the mouse still while you speak. If a slicer "
              "is clicked on camera, narrate only what actually changes - the test on 15_Validation confirms the cards "
              "and charts move together.", size=9.5)

    path = os.path.join(OUT, "Nishant_Loom_Presentation_Script.docx")
    doc.save(path)

    txt = [f"LOOM PRESENTATION SCRIPT - Nishant", "", f"Spoken word count: {words} words", ""]
    for b in script:
        txt += [f"[{b['t']}]  ON SCREEN: {b['view']}", b["say"], ""]
    txt += ["", "LIKELY EVALUATOR QUESTIONS", ""]
    for q, a in build_qa(F):
        txt += [f"Q: {q}", f"A: {a}", ""]
    txt += ["RECORDING CHECKLIST", ""] + [f"- {c}" for c in CHECKLIST]
    tpath = os.path.join(OUT, "Nishant_Loom_Presentation_Script.txt")
    open(tpath, "w", encoding="utf-8").write("\n".join(txt))
    return path, tpath


def to_pdf(docx_path):
    import win32com.client as win32

    pdf = docx_path.replace(".docx", ".pdf")
    word = win32.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = False
    try:
        d = word.Documents.Open(docx_path, ReadOnly=False)
        d.SaveAs(pdf, FileFormat=17)
        d.Close(False)
        return pdf
    finally:
        word.Quit()


if __name__ == "__main__":
    made = [build_report_doc(), build_guide_doc()]
    s_docx, s_txt = build_script_doc()
    made += [s_docx, s_txt]
    for p in list(made):
        if p.endswith(".docx"):
            try:
                made.append(to_pdf(p))
            except Exception as exc:  # noqa: BLE001
                print("pdf failed", p, exc)
    for p in made:
        print(os.path.basename(p), os.path.getsize(p))
