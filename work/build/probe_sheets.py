"""Open each generated sheet on its own and report whether Excel repairs the file."""
import os
import sys

import openpyxl
import win32com.client as win32

import build_base as B
import build_base2 as B2

B2.FACTS = B.F
B2.GLOBAL_FACTS = B.F

SHEETS = {
    "readme": B.s_readme, "raw": B.s_raw, "clean": B.s_clean, "dict": B.s_dictionary,
    "quality": B.s_quality, "kpi": B.s_kpi, "retailer": B.s_retailer, "brand": B.s_brand,
    "monthly": B.s_monthly, "price": B.s_price, "pivot": B2.s_pivot_placeholder,
    "dash": B2.s_dashboard, "insights": B2.s_insights, "impl": B2.s_implementation,
    "valid": B2.s_validation, "script": B2.s_script,
}


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
    keys = sys.argv[1:] or list(SHEETS)
    for k in keys:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        p = os.path.abspath(os.path.join("..", "tmp", f"only_{k}.xlsx"))
        try:
            SHEETS[k](wb)
            wb.save(p)
            print(k, can_open(p))
        except Exception as exc:  # noqa: BLE001
            print(k, "BUILD FAIL", type(exc).__name__, exc)
