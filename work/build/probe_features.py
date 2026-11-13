"""Micro-probes: which openpyxl construct makes Excel repair the file?"""
import os

import openpyxl
import win32com.client as win32
from openpyxl.chart import ScatterChart, Series, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.trendline import Trendline
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

TMP = os.path.abspath(os.path.join("..", "tmp"))


def case_cellis(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "OK"
    ws["A2"] = "bad"
    ws.conditional_formatting.add("A1:A2", CellIsRule(operator="equal", formula=['"OK"'],
                                                      font=Font(color="FF15803D")))
    wb.save(path)


def case_cellis_fill(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "OK"
    ws.conditional_formatting.add("A1", CellIsRule(operator="notEqual", formula=['"OK"'],
                                                   font=Font(color="FFB91C1C", bold=True),
                                                   fill=openpyxl.styles.PatternFill("solid", fgColor="FFFDE7E7")))
    wb.save(path)


def case_table(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "B", "C"])
    for i in range(2, 50):
        ws.cell(i, 1, i)
        ws.cell(i, 2, i * 2)
        ws.cell(i, 3, f"=A{i}*B{i}")
    t = Table(displayName="tblSales", ref="A1:C49")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    wb.save(path)


def case_table_formula_na(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", "B", "C", "D"])
    for i in range(2, 30):
        ws.cell(i, 1, i)
        ws.cell(i, 2, 0 if i == 5 else i * 2)
        ws.cell(i, 3, f"=IF(B{i}=0,NA(),A{i}/B{i})")
        ws.cell(i, 4, f'=IF(B{i}=0,"Zero",IF(ISNA(C{i}),"Undefined","OK"))')
    t = Table(displayName="tblSales", ref="A1:D29")
    ws.add_table(t)
    wb.save(path)


def case_scatter(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(2, 30):
        ws.cell(i, 1, i * 0.1)
        ws.cell(i, 2, 1000 + i * 37)
    ch = ScatterChart()
    ch.title = "t"
    ch.x_axis.title = "x"
    ch.y_axis.title = "y"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend = None
    xref = Reference(ws, min_col=1, min_row=2, max_row=29)
    yref = Reference(ws, min_col=2, min_row=2, max_row=29)
    ser = Series(yref, xref, title="p")
    ser.marker = Marker(symbol="circle", size=5)
    ser.marker.graphicalProperties = GraphicalProperties(solidFill="FF0070C0")
    ser.marker.graphicalProperties.line.solidFill = "FF0070C0"
    ser.graphicalProperties.line.noFill = True
    ser.trendline = Trendline(trendlineType="linear")
    ch.series.append(ser)
    ws.add_chart(ch, "D2")
    wb.save(path)


def case_scatter_simple(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for i in range(2, 30):
        ws.cell(i, 1, i * 0.1)
        ws.cell(i, 2, 1000 + i * 37)
    ch = ScatterChart()
    ch.title = "t"
    xref = Reference(ws, min_col=1, min_row=2, max_row=29)
    yref = Reference(ws, min_col=2, min_row=2, max_row=29)
    ch.series.append(Series(yref, xref, title="p"))
    ws.add_chart(ch, "D2")
    wb.save(path)


def case_merge(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.merge_cells("A1:S1")
    ws["A1"] = "title"
    ws.merge_cells("A2:C2")
    ws["A2"] = "x"
    for c in "ABCDEFGHIJKLMNOPQRS":
        ws.column_dimensions[c].width = 9.6
    wb.save(path)


def case_pp(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = 0.05
    ws["A1"].number_format = "+0.0;-0.0;0.0"
    ws["A2"] = 0.3
    ws["A2"].number_format = '"$"#,##0.0000'
    ws["A3"] = 0.3
    ws["A3"].number_format = '"Q"0'
    wb.save(path)


CASES = [case_cellis, case_cellis_fill, case_table, case_table_formula_na, case_scatter,
         case_scatter_simple, case_merge, case_pp]


def can_open(path):
    app = win32.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Open(path)
        wb.Close(False)
        return "OK"
    except Exception:
        return "REPAIR"
    finally:
        app.Quit()


if __name__ == "__main__":
    for fn in CASES:
        p = os.path.join(TMP, fn.__name__ + ".xlsx")
        try:
            fn(p)
            print(fn.__name__, can_open(p))
        except Exception as exc:  # noqa: BLE001
            print(fn.__name__, "BUILD FAIL", exc)
