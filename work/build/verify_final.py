"""Open the delivered workbook, recalculate, and verify structure + numbers."""
from __future__ import annotations

import json
import os

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
FINAL = os.path.join(ROOT, "deliverables", "Coca_Cola_Sales_Analysis.xlsx")
F = json.load(open(os.path.join(HERE, "..", "facts.json"), encoding="utf-8"))
T = F["total"]
RET = {r["name"]: r for r in F["levels"]["retailer"]}
REG = {r["name"]: r for r in F["levels"]["region"]}
BRD = {r["name"]: r for r in F["levels"]["brand"]}

out = {}
app = win32.gencache.EnsureDispatch("Excel.Application")
app.Visible = False
app.DisplayAlerts = False
wb = app.Workbooks.Open(FINAL)
try:
    app.CalculateFullRebuild()
    out["sheets"] = [wb.Worksheets(i).Name for i in range(1, wb.Worksheets.Count + 1)]

    # error scan across used ranges
    errs = []
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        try:
            used = ws.UsedRange
            if used.Cells.Count > 600000:
                continue
            vals = used.Value
            if not isinstance(vals, tuple):
                vals = ((vals,),)
            for r, row in enumerate(vals, start=used.Row):
                if not isinstance(row, tuple):
                    row = (row,)
                for c, v in enumerate(row, start=used.Column):
                    if isinstance(v, int) and v in (-2146826281, -2146826246, -2146826259, -2146826288,
                                                     -2146826273, -2146826252, -2146826265):
                        errs.append((ws.Name, r, c, v))
        except Exception as exc:  # noqa: BLE001
            errs.append((ws.Name, 0, 0, str(exc)[:60]))
    out["error_cells"] = errs[:40]
    out["error_count"] = len(errs)

    # KPI / reconciliation
    k = wb.Worksheets("05_KPI_Summary")
    out["kpi"] = {
        "sales": k.Range("B6").Value, "units": k.Range("B7").Value, "profit": k.Range("B8").Value,
        "margin": k.Range("B9").Value, "spu": k.Range("B10").Value,
        "weighted_listed": k.Range("B11").Value, "ppu": k.Range("B12").Value,
        "records": k.Range("B13").Value, "cities": k.Range("B14").Value, "brands": k.Range("B15").Value,
        "simple_avg": k.Range("B19").Value, "overstatement_pp": k.Range("B21").Value,
    }
    out["coke_brand"] = {
        "sales": k.Range("B25").Value, "units": k.Range("B26").Value,
        "profit": k.Range("B27").Value, "margin": k.Range("B28").Value,
    }
    r6 = wb.Worksheets("06_Retailer_Analysis")
    out["retailer_block"] = [[r6.Cells(r, c).Value for c in range(1, 11)] for r in range(5, 10)]
    r9 = wb.Worksheets("09_Monthly_Analysis")
    out["month_total"] = [r9.Cells(17, c).Value for c in range(1, 13)]
    out["month_rows"] = [[r9.Cells(r, c).Value for c in range(2, 12)] for r in range(5, 17)]
    v = wb.Worksheets("15_Validation")
    out["validation"] = [[v.Cells(r, 1).Value, v.Cells(r, 2).Value, v.Cells(r, 3).Value,
                          v.Cells(r, 4).Value] for r in range(5, 28)]

    # pivots
    wsP = wb.Worksheets("11_Pivot_Tables")
    out["pivot_count"] = wsP.PivotTables().Count
    prev = {}
    dash_only = {"PT_DashHeatmap", "PT_D_Retailer", "PT_D_Region", "PT_D_Brand", "PT_D_Month"}
    for name in ("PT_KPI", "PT_Retailer", "PT_Region", "PT_Brand", "PT_RegionMonth",
                 "PT_BrandRegion", "PT_RetailerRegion", "PT_DashHeatmap",
                 "PT_D_Retailer", "PT_D_Region", "PT_D_Brand", "PT_D_Month"):
        try:
            pt = wb.Worksheets("12_Dashboard").PivotTables(name) if name in dash_only else \
                wsP.PivotTables(name)
            prev[name] = {"range": str(pt.TableRange1.Address),
                          "value": str(pt.TableRange1.Value)[:300]}
        except Exception as exc:  # noqa: BLE001
            prev[name] = f"ERR {exc}"
    out["pivots"] = prev
    out["slicer_count"] = wb.SlicerCaches.Count
    out["slicer_names"] = [wb.SlicerCaches(i).Name for i in range(1, wb.SlicerCaches.Count + 1)]

    wsD = wb.Worksheets("12_Dashboard")
    out["charts"] = []
    for i in range(1, wsD.ChartObjects().Count + 1):
        ch = wsD.ChartObjects(i)
        out["charts"].append({"name": ch.Name, "left": round(ch.Left), "top": round(ch.Top),
                              "w": round(ch.Width), "h": round(ch.Height),
                              "title": ch.Chart.ChartTitle.Text if ch.Chart.HasTitle else None})
    out["slicer_shapes"] = []
    for i in range(1, wsD.Shapes.Count + 1):
        sh = wsD.Shapes(i)
        if sh.Type == 25:
            out["slicer_shapes"].append({"name": sh.Name, "left": round(sh.Left), "top": round(sh.Top),
                                         "w": round(sh.Width), "h": round(sh.Height)})
    out["cards"] = {c: wsD.Range(c).Value for c in ("A6", "E6", "I6", "M6", "Q6")}
    out["clean_ok_rows"] = wb.Worksheets("02_Clean_Data").Evaluate(
        '=COUNTIF(\'02_Clean_Data\'!AA:AA,"OK")')
    out["clean_status_counts"] = wb.Worksheets("02_Clean_Data").Evaluate(
        '=COUNTIF(\'02_Clean_Data\'!AA:AA,"Zero sales - check")')
finally:
    wb.Close(SaveChanges=False)
    app.Quit()

json.dump(out, open(os.path.join(HERE, "..", "verify_final.json"), "w", encoding="utf-8"),
          indent=1, default=str)

print("sheets", len(out["sheets"]))
print("errors", out["error_count"], out["error_cells"][:5])
print("kpi", out["kpi"])
print("coke", out["coke_brand"])
print("retailer rows:")
for row in out["retailer_block"]:
    print("  ", row)
print("month total", out["month_total"])
print("pivots", out["pivot_count"], "slicers", out["slicer_count"], out["slicer_names"])
print("charts", json.dumps(out["charts"], indent=1))
print("slicer shapes", out["slicer_shapes"])
print("cards", out["cards"])
print("clean OK rows", out["clean_ok_rows"], "zero-check rows", out["clean_status_counts"])
print("VALIDATION")
for row in out["validation"]:
    print("  ", row)
