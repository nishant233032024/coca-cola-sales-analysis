"""Excel COM stage: native PivotTables, slicers, timeline, pivot charts, KPI cards."""
from __future__ import annotations

import json
import os
import time

import pythoncom
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
BASE = os.path.join(ROOT, "deliverables", "Coca_Cola_Sales_Analysis_base.xlsx")
FINAL = os.path.join(ROOT, "deliverables", "Coca_Cola_Sales_Analysis.xlsx")
REPORT = os.path.join(HERE, "..", "pivot_report.json")

F = json.load(open(os.path.join(HERE, "..", "facts.json"), encoding="utf-8"))
T = F["total"]
RET = {r["name"]: r for r in F["levels"]["retailer"]}

xlDatabase, xlRowField, xlColumnField, xlDataField = 1, 1, 2, 4
xlSum, xlTabularRow = -4157, 1
xlBarClustered, xlColumnClustered, xlLine = 57, 51, 4
xlOpenXMLWorkbook = 51
xlDescending, xlAscending = 2, 1
BLUE, DARK, GREY = 12611584, 6108695, 6710886
CALC = {"Margin %", "Sales per Unit"}
FMT = {"Sales": '"$"#,##0', "Units": "#,##0", "Profit": '"$"#,##0',
       "Margin pct": "0.0%", "Sales/Unit": '"$"0.0000'}

report = {"pivots": [], "dash_pivots": [], "slicers": [], "charts": [], "errors": [], "notes": []}
deferred = []
calc_ready = [False]


def step(msg):
    print(msg)
    report["notes"].append(msg)


def add_val(pt, src_field, label):
    df = pt.PivotFields(src_field)
    df.Orientation = xlDataField
    try:
        df.Function = xlSum
    except Exception as exc:  # noqa: BLE001
        step(f"function set failed {src_field}: {exc}")
    try:
        df.Name = label
    except Exception:
        try:
            df.Caption = label
        except Exception as exc:  # noqa: BLE001
            step(f"caption {label} kept Excel default: {exc}")
    try:
        df.NumberFormat = FMT.get(label, "General")
    except Exception:  # noqa: BLE001
        pass


def build():
    app = win32.gencache.EnsureDispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.AskToUpdateLinks = False
    wb = app.Workbooks.Open(BASE)
    try:
        app.CalculateFullRebuild()
        wsD = wb.Worksheets("12_Dashboard")
        wsP = wb.Worksheets("11_Pivot_Tables")
        wsData = wb.Worksheets("02_Clean_Data")
        wsVal = wb.Worksheets("15_Validation")
        step(f"recalc check: sales={wb.Worksheets('05_KPI_Summary').Range('B6').Value!r}")

        src = "'" + wsData.Name + "'!" + wsData.Range("A1:AA3889").Address
        pc = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=src)
        step(f"cache source={src} records={pc.RecordCount}")

        def make(name, dest, rows=None, cols=None, values=(), sort_by=None, sort_desc=True,
                 grand_rows=True):
            pt = pc.CreatePivotTable(TableDestination=dest[0].Range(dest[1]), TableName=name)
            try:
                pt.RowAxisLayout(xlTabularRow)
            except Exception:  # noqa: BLE001
                pass
            if rows:
                f = pt.PivotFields(rows)
                f.Orientation = xlRowField
                f.Position = 1
            if cols:
                f = pt.PivotFields(cols)
                f.Orientation = xlColumnField
                f.Position = 1
            for src_field, label in values:
                if src_field in CALC and not calc_ready[0]:
                    deferred.append((pt, src_field, label))
                else:
                    add_val(pt, src_field, label)
            try:
                pt.RowGrand = grand_rows
                # ColumnGrand=False also deletes the Grand Total row of a single-row-field
                # pivot, so only switch it off when there is a column field (or no total wanted)
                pt.ColumnGrand = bool(rows) and not cols and grand_rows
                pt.TableStyle2 = "PivotStyleMedium2"
            except Exception:  # noqa: BLE001
                pass
            if rows and sort_by:
                try:
                    pt.PivotFields(rows).AutoSort(xlDescending if sort_desc else xlAscending, sort_by)
                except Exception as exc:  # noqa: BLE001
                    step(f"sort failed {name}: {exc}")
            report["pivots"].append({"name": name, "dest": str(dest[1]), "range": str(pt.TableRange1.Address),
                                     "rows": rows, "cols": cols, "values": [v[1] for v in values]})
            step(f"pivot {name} -> {pt.TableRange1.Address}")
            return pt

        pt_kpi = make("PT_KPI", (wsP, "A16"), values=(
            ("Total Sales", "Sales"), ("Units Sold", "Units"),
            ("Operating Profit", "Profit"), ("Margin %", "Margin pct"),
            ("Sales per Unit", "Sales/Unit")))

        # calculated fields live on the cache but are reached through a PivotTable
        fields = []
        for fname, formula in (("Margin %", "='Operating Profit'/'Total Sales'"),
                               ("Sales per Unit", "='Total Sales'/'Units Sold'")):
            try:
                cf = pt_kpi.CalculatedFields().Add(fname, formula, True)
                fields.append(fname)
                report[f"calc_{fname}"] = True
                step(f"calculated field {fname} added")
            except Exception as exc:  # noqa: BLE001
                report["errors"].append(f"calc {fname}: {exc}")
                report[f"calc_{fname}"] = str(exc)
                step(f"calculated field {fname} failed: {exc}")

        calc_ready[0] = True
        for pt, src_field, label in deferred:
            try:
                add_val(pt, src_field, label)
            except Exception as exc:  # noqa: BLE001
                report["errors"].append(f"value {label} on {pt.Name}: {exc}")
                step(f"value field {label} failed: {exc}")
        deferred.clear()

        make("PT_Retailer", (wsP, "A24"), rows="Retailer", values=(
            ("Total Sales", "Sales"), ("Units Sold", "Units"), ("Operating Profit", "Profit"),
            ("Margin %", "Margin pct")), sort_by="Sales")
        make("PT_Region", (wsP, "A32"), rows="Region", values=(
            ("Total Sales", "Sales"), ("Units Sold", "Units"), ("Operating Profit", "Profit"),
            ("Margin %", "Margin pct")), sort_by="Sales")
        make("PT_Month", (wsP, "A40"), rows="Year-Month", values=(
            ("Total Sales", "Sales"), ("Units Sold", "Units"), ("Operating Profit", "Profit"),
            ("Margin %", "Margin pct")), sort_by="Year-Month", sort_desc=False)
        make("PT_Brand", (wsP, "A56"), rows="Beverage Brand", values=(
            ("Total Sales", "Sales"), ("Units Sold", "Units"), ("Operating Profit", "Profit"),
            ("Margin %", "Margin pct")), sort_by="Sales")
        make("PT_RegionMonth", (wsP, "H16"), rows="Region", cols="Year-Month", values=(
            ("Total Sales", "Sales"),))
        make("PT_BrandRegion", (wsP, "H32"), rows="Beverage Brand", cols="Region", values=(
            ("Total Sales", "Sales"),), sort_by="Sales")
        make("PT_RetailerRegion", (wsP, "H44"), rows="Retailer", cols="Region", values=(
            ("Total Sales", "Sales"), ("Operating Profit", "Profit")), sort_by="Sales")

        pt_ret_chart = make("PT_RetailerChart", (wsP, "A70"), rows="Retailer",
                            values=(("Total Sales", "Sales"),), sort_by="Sales", grand_rows=False)
        pt_reg_chart = make("PT_RegionChart", (wsP, "A78"), rows="Region",
                            values=(("Total Sales", "Sales"),), sort_by="Sales", grand_rows=False)
        pt_ms = make("PT_MonthSalesChart", (wsP, "A86"), rows="Year-Month",
                     values=(("Total Sales", "Sales"),), sort_by="Year-Month", sort_desc=False,
                     grand_rows=False)
        pt_mm = make("PT_MonthMarginChart", (wsP, "A100"), rows="Year-Month",
                     values=(("Margin %", "Margin pct"),), sort_by="Year-Month", sort_desc=False,
                     grand_rows=False)
        pt_brand_chart = make("PT_BrandChart", (wsP, "A120"), rows="Beverage Brand",
                              values=(("Total Sales", "Sales"),), sort_by="Sales", grand_rows=False)
        pt_heat = make("PT_DashHeatmap", (wsD, "A39"), rows="Region", cols="Year-Month",
                       values=(("Total Sales", "Sales"),))
        try:
            pt_heat.RowGrand = True
            pt_heat.ColumnGrand = False
        except Exception:  # noqa: BLE001
            pass

        # ------------------------------------------------------------ charts
        def zone(addr):
            rng = wsD.Range(addr)
            return rng.Left, rng.Top, rng.Width, rng.Height

        def add_chart(pt, addr, kind, title, fmt=None, reverse_cats=False):
            left, top, width, height = zone(addr)
            co = wsD.ChartObjects().Add(left + 2, top + 2, width - 4, height - 4)
            co.Chart.SetSourceData(pt.TableRange1)
            co.Chart.ChartType = kind
            co.Chart.HasTitle = True
            co.Chart.ChartTitle.Text = title
            for prop in ("Size",):
                pass
            try:
                co.Chart.ChartTitle.Font.Size = 11
                co.Chart.ChartTitle.Font.Color = DARK
                co.Chart.ChartTitle.Font.Bold = True
            except Exception:  # noqa: BLE001
                pass
            co.Chart.HasLegend = False
            try:
                co.Chart.ShowAllFieldButtons = False
                co.Chart.ShowAxisFieldButtons = False
                co.Chart.ShowLegendFieldButtons = False
                co.Chart.ShowReportFilterFieldButtons = False
                co.Chart.ShowValueFieldButtons = False
            except Exception:  # noqa: BLE001
                pass
            try:
                co.Chart.ChartArea.Format.Fill.Visible = False
                co.Chart.ChartArea.Format.Line.Visible = False
                co.Chart.PlotArea.Format.Fill.Visible = False
            except Exception:  # noqa: BLE001
                pass
            try:
                ser = co.Chart.SeriesCollection(1)
                ser.Format.Fill.ForeColor.RGB = BLUE
                ser.Format.Line.ForeColor.RGB = BLUE
                ser.Format.Line.Weight = 2.25
                if kind in (xlBarClustered, xlColumnClustered):
                    ser.Format.Line.Visible = False
            except Exception:  # noqa: BLE001
                pass
            try:
                ax = co.Chart.Axes(2)
                ax.HasMajorGridlines = True
                ax.MajorGridlines.Format.Line.ForeColor.RGB = 14277081
                ax.TickLabels.Font.Size = 8
                ax.TickLabels.Font.Color = GREY
                if fmt:
                    ax.TickLabels.NumberFormat = fmt
                elif kind == xlBarClustered:
                    ax.TickLabels.NumberFormat = '#,##0,"k"'
                    ax.HasTitle = True
                    ax.AxisTitle.Text = "USD thousands"
                    ax.AxisTitle.Font.Size = 8
            except Exception:  # noqa: BLE001
                pass
            try:
                cat = co.Chart.Axes(1)
                cat.TickLabels.Font.Size = 8
                cat.TickLabels.Font.Color = 3355443
                cat.Format.Line.ForeColor.RGB = 14277081
                if reverse_cats:
                    cat.ReversePlotOrder = True
            except Exception:  # noqa: BLE001
                pass
            report["charts"].append({"title": title, "anchor": addr, "source": pt.Name})
            step(f"chart '{title}' at {addr}")
            return co

        add_chart(pt_ret_chart, "A10:I23", xlBarClustered, "Sales by retailer (USD)", reverse_cats=True)
        add_chart(pt_reg_chart, "K10:S23", xlBarClustered, "Sales by region (USD)", reverse_cats=True)
        add_chart(pt_ms, "A25:I37", xlLine, "Monthly sales, 2021 (USD)")
        add_chart(pt_mm, "K25:S37", xlLine, "Monthly weighted operating margin", fmt="0.0%")
        add_chart(pt_brand_chart, "A52:I64", xlBarClustered, "Sales by beverage brand (USD)", reverse_cats=True)

        try:
            data = wsD.Range(pt_heat.TableRange1.Address)
            data.FormatConditions.Delete()
            cs = data.FormatConditions.AddColorScale(2)
            cs.ColorScaleCriteria(1).Type = 1
            cs.ColorScaleCriteria(1).FormatColor.Color = 16777215
            cs.ColorScaleCriteria(2).Type = 2
            cs.ColorScaleCriteria(2).FormatColor.Color = BLUE
            pt_heat.TableRange1.Font.Size = 8
            step("heatmap colour scale applied")
        except Exception as exc:  # noqa: BLE001
            report["errors"].append(f"heatmap: {exc}")
            step(f"heatmap formatting failed: {exc}")

        # ------------------------------------------- dashboard summary pivots
        # four native PivotTables on the dashboard itself, so every chart has its
        # numbers printed next to it and the slicers filter them at the same time
        wsD.Range("A67:S67").Merge()
        wsD.Range("A67").Value = ("SUMMARY PIVOTS  |  four native PivotTables on the same cache - the three slicers "
                                  "filter these panels and every chart together")
        wsD.Range("A67").Font.Bold = True
        wsD.Range("A67").Font.Size = 10
        wsD.Range("A67").Font.Color = DARK
        for cell, text in (("A69", "RETAILER"), ("G69", "REGION"), ("M69", "BEVERAGE BRAND"),
                           ("A79", "MONTH")):
            wsD.Range(cell).Value = text
            wsD.Range(cell).Font.Bold = True
            wsD.Range(cell).Font.Size = 9
            wsD.Range(cell).Font.Color = GREY

        wsD.Columns("G").ColumnWidth = 10.5  # room for "Northeast" in the region panel

        dash_vals = (("Total Sales", "Sales"), ("Units Sold", "Units"),
                     ("Operating Profit", "Profit"), ("Margin %", "Margin pct"))
        for name, addr, field, sort_by, desc in (("PT_D_Retailer", "A70", "Retailer", "Sales", True),
                                                ("PT_D_Region", "G70", "Region", "Sales", True),
                                                ("PT_D_Brand", "M70", "Beverage Brand", "Sales", True),
                                                ("PT_D_Month", "A80", "Year-Month", "Year-Month", False)):
            try:
                pt = make(name, (wsD, addr), rows=field, values=dash_vals,
                          sort_by=sort_by, sort_desc=desc)
                try:
                    pt.RowFields(1).Subtotals = [False] * 12
                except Exception:  # noqa: BLE001
                    pass
                # compact $k / k formats so every number fits the dashboard column widths
                for df in pt.DataFields:
                    try:
                        src = df.SourceName
                        if src == "Total Sales":
                            df.Caption, df.NumberFormat = "Sales $k", '"$"#,##0,"k"'
                        elif src == "Operating Profit":
                            df.Caption, df.NumberFormat = "Profit $k", '"$"#,##0,"k"'
                        elif src == "Units Sold":
                            df.Caption, df.NumberFormat = "Units k", '#,##0,"k"'
                        else:
                            df.NumberFormat = "0.0%"
                    except Exception:  # noqa: BLE001
                        pass
                pt.TableRange1.Font.Size = 9
                pt.TableRange1.Rows(1).Font.Size = 9
                report["dash_pivots"].append({"name": name, "range": str(pt.TableRange1.Address),
                                              "total_sales": pt.TableRange1.Cells(
                                                  pt.TableRange1.Rows.Count, 2).Value})
                step(f"dashboard summary pivot {name} -> {pt.TableRange1.Address}")
            except Exception as exc:  # noqa: BLE001
                report["errors"].append(f"dashboard pivot {name}: {exc}")
                step(f"dashboard pivot {name} failed: {exc}")

        try:
            ps = wsD.PageSetup
            ps.PrintArea = "$A$1:$S$94"
            ps.Orientation = 2  # xlLandscape
            ps.Zoom = False
            ps.FitToPagesWide = 1
            ps.FitToPagesTall = False
            step("dashboard print area A1:S94, landscape, fit to one page wide")
        except Exception as exc:  # noqa: BLE001
            step(f"dashboard page setup failed: {exc}")

        # ------------------------------------------------------------ slicers
        sc_ret = None
        for field, top, height in (("Retailer", 40, 92), ("Region", 140, 110), ("Beverage Brand", 258, 150)):
            try:
                sc = wb.SlicerCaches.Add2(pt_kpi, field)
                sc.Slicers.Add(wsD, pythoncom.Empty, field, field, top, wsD.Range('U1').Left + 6, 168, height)
                report["slicers"].append(field)
                if field == "Retailer":
                    sc_ret = sc
                step(f"slicer {field} added")
            except Exception as exc:  # noqa: BLE001
                report["errors"].append(f"slicer {field}: {exc}")
                step(f"slicer {field} failed: {exc}")

        # a slicer cache created through COM drives only the PivotTable it was built from,
        # so wire every PivotTable to every slicer cache explicitly - otherwise the charts,
        # the summary panels and the heatmap would ignore the slicers
        try:
            all_pivots = []
            for i in range(1, wb.Worksheets.Count + 1):
                ws = wb.Worksheets(i)
                for j in range(1, ws.PivotTables().Count + 1):
                    all_pivots.append(ws.PivotTables(j))
            n_conn = 0
            for k in range(1, wb.SlicerCaches.Count + 1):
                sc_k = wb.SlicerCaches(k)
                for pt in all_pivots:
                    try:
                        sc_k.PivotTables.AddPivotTable(pt)
                        n_conn += 1
                    except Exception:  # noqa: BLE001
                        pass
            report["slicer_connections"] = n_conn
            step(f"connected {len(all_pivots)} pivot tables to {wb.SlicerCaches.Count} slicer caches "
                 f"({n_conn} links)")
        except Exception as exc:  # noqa: BLE001
            report["errors"].append(f"slicer connect: {exc}")
            step(f"slicer connect failed: {exc}")

        report["notes"].append("Invoice Date timeline: not created by automation; one-click manual step, "
                               "documented on 14_Implementation.")

        # ------------------------------------------------------------ cards
        names = {}
        try:
            for d in pt_kpi.DataFields:
                try:
                    names[d.SourceName] = d.Name
                except Exception:  # noqa: BLE001
                    pass
        except Exception as exc:  # noqa: BLE001
            step(f"data-field readback failed: {exc}")
        step(f"data field names: {names}")
        ref = "'11_Pivot_Tables'!$A$16"
        for cell, field, fmt, caption in (
                ("A6", "Total Sales", '"$"#,##0', "all six brands"),
                ("E6", "Units Sold", "#,##0", "all six brands"),
                ("I6", "Operating Profit", '"$"#,##0', "all six brands"),
                ("M6", "Margin %", "0.0%", "profit / sales"),
                ("Q6", "Sales per Unit", '"$"0.0000', "realised")):
            try:
                wsD.Range(cell).Formula = f'=IFERROR(GETPIVOTDATA("{names.get(field, field)}",{ref}),"")'
                wsD.Range(cell).NumberFormat = fmt
                wsD.Range(cell.replace("6", "7")).Value = caption
                step(f"card {cell} <- GETPIVOTDATA({field})")
            except Exception as exc:  # noqa: BLE001
                report["errors"].append(f"card {cell}: {exc}")
                step(f"card {cell} failed: {exc}")

        app.CalculateFullRebuild()
        time.sleep(0.6)

        def card(cell):
            return wsD.Range(cell).Value

        report["card_full_sales"] = card("A6")
        report["card_full_margin"] = card("M6")
        report["expected_sales"] = T["sales"]
        report["expected_margin"] = T["margin"]
        step(f"cards: sales={card('A6')!r} units={card('E6')!r} profit={card('I6')!r} "
             f"margin={card('M6')!r} spu={card('Q6')!r}")

        # ------------------------------------------------------------ tests
        try:
            sc_ret.ClearManualFilter()
            for i in range(1, sc_ret.SlicerItems.Count + 1):
                it = sc_ret.SlicerItems(i)
                if it.Name != "Sodapop":
                    it.Selected = False
            app.CalculateFullRebuild()
            time.sleep(0.4)
            f_s, f_m, f_u = card("A6"), card("M6"), card("E6")
            ok1 = isinstance(f_s, (int, float)) and abs(f_s - RET["Sodapop"]["sales"]) < 1
            ok2 = isinstance(f_m, (int, float)) and abs(f_m - RET["Sodapop"]["margin"]) < 0.0006
            report["slicer_test"] = {"sales": f_s, "margin": f_m, "units": f_u,
                                     "expected_sales": RET["Sodapop"]["sales"],
                                     "expected_margin": RET["Sodapop"]["margin"], "pass": bool(ok1 and ok2)}
            wsVal.Range("C25").Value = "MATCH" if (ok1 and ok2) else "MISMATCH"
            step(f"Sodapop filter: sales={f_s!r} margin={f_m!r} -> {'MATCH' if ok1 and ok2 else 'MISMATCH'}")
            # the same filter must reach the summary pivots and, through them, the pivot charts
            pt_d_ret = wsD.PivotTables("PT_D_Retailer")
            pt_d_ret.RefreshTable()
            rng = pt_d_ret.TableRange1
            vis = [rng.Cells(r, 1).Value for r in range(2, rng.Rows.Count + 1)]
            ok4 = bool(vis) and vis[0] == "Sodapop" and "FizzySip" not in vis \
                and isinstance(rng.Cells(2, 2).Value, (int, float)) \
                and abs(rng.Cells(2, 2).Value - RET["Sodapop"]["sales"]) < 1
            wsVal.Range("C27").Value = "MATCH" if ok4 else "MISMATCH"
            report["pivot_slicer_test"] = {"visible": vis, "pass": bool(ok4)}
            step(f"dashboard pivot follows slicer: {vis} -> {'MATCH' if ok4 else 'MISMATCH'}")
            sc_ret.ClearManualFilter()
            app.CalculateFullRebuild()
            time.sleep(0.4)
            g_s = card("A6")
            ok3 = isinstance(g_s, (int, float)) and abs(g_s - T["sales"]) < 1
            report["slicer_clear"] = {"sales": g_s, "pass": bool(ok3)}
            wsVal.Range("C26").Value = "MATCH" if ok3 else "MISMATCH"
            step(f"cleared: sales={g_s!r} -> {'MATCH' if ok3 else 'MISMATCH'}")
        except Exception as exc:  # noqa: BLE001
            report["errors"].append(f"slicer test: {exc}")
            step(f"slicer test failed: {exc}")
            wsVal.Range("C25").Value = "NOT TESTED"
            wsVal.Range("C26").Value = "NOT TESTED"

        try:
            wsVal.Range("C23").Value = wsP.PivotTables().Count + wsD.PivotTables().Count
            wsVal.Range("C24").Value = wb.SlicerCaches.Count
        except Exception as exc:  # noqa: BLE001
            step(f"counts write failed: {exc}")

        # leave every pivot showing the full portfolio, not the test filter
        try:
            for i in range(1, wb.Worksheets.Count + 1):
                ws_i = wb.Worksheets(i)
                for j in range(1, ws_i.PivotTables().Count + 1):
                    ws_i.PivotTables(j).RefreshTable()
            step("all pivot tables refreshed to the unfiltered view")
        except Exception as exc:  # noqa: BLE001
            step(f"final pivot refresh failed: {exc}")

        app.CalculateFullRebuild()
        wb.SaveAs(FINAL, FileFormat=xlOpenXMLWorkbook)
        step(f"saved {FINAL}")
        report["saved"] = FINAL

        try:
            vals = {}
            for name in ("PT_KPI", "PT_Retailer", "PT_Region", "PT_Month"):
                pt = wsP.PivotTables(name)
                vals[name] = str(pt.TableRange1.Value)[:600]
            report["pivot_preview"] = vals
        except Exception as exc:  # noqa: BLE001
            step(f"readback failed: {exc}")

        wb.Close(SaveChanges=False)
    finally:
        try:
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
        json.dump(report, open(REPORT, "w", encoding="utf-8"), indent=1, default=str)


if __name__ == "__main__":
    build()
