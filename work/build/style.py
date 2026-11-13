"""Shared visual system: clean blue theme."""
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BLUE = "FF0070C0"
DARK = "FF17365D"
LIGHT = "FF9DC3E6"
PALE = "FFDCE9F7"
WHITE = "FFFFFFFF"
CHAR = "FF111827"
GREY = "FF6B7280"
BG = "FFF3F4F6"
LINE = "FFD1D5DB"
RED = "FFB91C1C"
GREEN = "FF15803D"

MONEY = '"$"#,##0'
MONEY2 = '"$"#,##0.00'
MONEY3 = '"$"#,##0.000'
MONEY4 = '"$"#,##0.0000'
PCT = "0.0%"
PCT2 = "0.00%"
PP = '+0.0;-0.0;0.0'
INT = "#,##0"
DATE = "yyyy-mm-dd"

THIN = Side(style="thin", color=LINE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def fill(hex_):
    return PatternFill("solid", fgColor=hex_)


def title(ws, text, sub=None, last_col=12, height=30):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(1, 1, text)
    c.font = Font(bold=True, size=16, color=WHITE)
    c.fill = fill(DARK)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = height
    if sub:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
        s = ws.cell(2, 1, sub)
        s.font = Font(size=9, italic=True, color=GREY)
        s.alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[2].height = 16


def section(ws, row, text, last_col=12):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=last_col)
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, size=11, color=DARK)
    c.fill = fill(PALE)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 20


def headers(ws, row, labels, start=1, wrap=True, height=30):
    for i, lab in enumerate(labels):
        c = ws.cell(row, start + i, lab)
        c.font = Font(bold=True, size=9, color=WHITE)
        c.fill = fill(BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
        c.border = BOX
    ws.row_dimensions[row].height = height


def put(ws, row, col, value, fmt=None, bold=False, color=CHAR, size=10, align=None,
        border=True, wrap=False, bg=None, italic=False, literal=False):
    c = ws.cell(row, col, value)
    if literal and isinstance(value, str):
        c.data_type = "s"
    c.font = Font(bold=bold, size=size, color=color, italic=italic)
    if fmt:
        c.number_format = fmt
    c.alignment = Alignment(horizontal=align or ("left" if isinstance(value, str) else "right"),
                            vertical="center", wrap_text=wrap)
    if border:
        c.border = BOX
    if bg:
        c.fill = fill(bg)
    return c


def note(ws, row, text, col=1, last_col=None, size=9):
    if last_col:
        ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=last_col)
    c = ws.cell(row, col, text)
    c.font = Font(size=size, italic=True, color=GREY)
    c.alignment = Alignment(vertical="top", wrap_text=True)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def total_row(ws, row, col_from, col_to, fill_hex="FFE8EEF6"):
    for c in range(col_from, col_to + 1):
        cell = ws.cell(row, c)
        cell.font = Font(bold=True, size=10, color=CHAR)
        cell.fill = fill(fill_hex)
