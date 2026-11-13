#!/usr/bin/env python3
"""Reproduce the numerical analysis behind Coca_Cola_Sales_Analysis.xlsx.

Standalone: reads the supplied workbook and prints every figure used in the
dashboard, report and Loom script, plus the reconciliation checks.

Requirements
------------
    Python 3.9+ and openpyxl   (pip install openpyxl)

Usage
-----
    python Reproduce_Coca_Cola_Analysis.py --input "Coke-sales-analysis-stubs (1).xlsx"
    python Reproduce_Coca_Cola_Analysis.py --input source.xlsx --json facts_out.json

What it does NOT do
-------------------
It does not create PivotTables, slicers, pivot charts or Excel formulas. Those
are native Excel objects and were built in Excel itself; see
Coca_Cola_Excel_Implementation_Guide.docx for the click-by-click path.

Exit code is 1 if any internal identity check fails, 0 otherwise.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from collections import defaultdict

try:
    import openpyxl
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required: pip install openpyxl")

HEADER_ROW = 5
FIRST_ROW = 6
FIELDS = ["Retailer", "Retailer ID", "Invoice Date", "Region", "State", "City", "Beverage Brand",
          "Price per Unit", "Units Sold", "Total Sales", "Operating Profit", "Operating Margin"]


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if "Data" not in wb.sheetnames:
        sys.exit(f"No 'Data' sheet found. Sheets present: {wb.sheetnames}")
    ws = wb["Data"]
    headers = [ws.cell(HEADER_ROW, c).value for c in range(2, 14)]
    if headers != FIELDS:
        sys.exit(f"Unexpected headers on row {HEADER_ROW}: {headers}")

    rows, blank = [], 0
    for r in range(FIRST_ROW, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(2, 14)]
        if all(v is None for v in vals):
            blank += 1
            continue
        (ret, rid, dt, reg, st, city, brand, price, units, sales, profit, margin) = vals
        rows.append(dict(source_row=r, retailer=ret, retailer_id=str(rid), date=dt, region=reg, state=st,
                         city=city, location=f"{st} | {city}", brand=brand, price=float(price),
                         units=int(units), sales=float(sales), profit=float(profit),
                         margin=float(margin), month=dt.month, year=dt.year,
                         year_month=f"{dt.year}-{dt.month:02d}"))
    return rows, blank


def agg(rows):
    s = sum(r["sales"] for r in rows)
    u = sum(r["units"] for r in rows)
    p = sum(r["profit"] for r in rows)
    listed = sum(r["price"] * r["units"] for r in rows)
    return dict(records=len(rows), sales=s, units=u, profit=p,
                margin=(p / s if s else None),
                sales_per_unit=(s / u if u else None),
                profit_per_unit=(p / u if u else None),
                listed_price=(listed / u if u else None))


def pearson(xs, ys):
    n = len(xs)
    sx, sy = sum(xs), sum(ys)
    sxy = sum(a * b for a, b in zip(xs, ys))
    sx2, sy2 = sum(a * a for a in xs), sum(b * b for b in ys)
    den = math.sqrt((n * sx2 - sx * sx) * (n * sy2 - sy * sy))
    return (n * sxy - sx * sy) / den if den else None


def by(rows, key):
    g = defaultdict(list)
    for r in rows:
        g[r[key]].append(r)
    return g


def money(x):
    return f"${x:,.0f}"


def pct(x):
    return f"{x * 100:.1f}%"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="path to the supplied .xlsx file")
    ap.add_argument("--json", help="optional path to write the figures as JSON")
    args = ap.parse_args()

    rows, blank = load(args.input)
    total = agg(rows)
    checks = []

    # ---- identity and quality checks
    sales_breaks = [r for r in rows if abs(r["price"] * r["units"] - r["sales"]) > 0.01]
    margin_breaks = [r for r in rows if r["sales"] and abs(r["profit"] / r["sales"] - r["margin"]) > 0.0005]
    keys = {(r["retailer"], r["location"], r["brand"], r["year_month"]) for r in rows}
    full = {(r["retailer"], r["location"], r["brand"], r["year_month"]) for r in rows}
    checks.append(("Sales identity (Price x Units = Total Sales)", 0, len(sales_breaks)))
    checks.append(("Margin identity (Profit / Sales = supplied margin)", 0, len(margin_breaks)))
    checks.append(("Negative values", 0, sum(1 for r in rows if r["sales"] < 0 or r["profit"] < 0 or r["units"] < 0)))
    checks.append(("Duplicate retailer-location-brand-month keys", 0, len(rows) - len(keys)))
    checks.append(("Blank rows skipped", 0, blank))
    months = {r["year_month"] for r in rows}
    per_month = {m: len([r for r in rows if r["year_month"] == m]) for m in months}
    checks.append(("Months with partial coverage", 0, sum(1 for v in per_month.values() if v != 324)))
    checks.append(("Zero-sales records (flagged, retained)", 2,
                   sum(1 for r in rows if r["sales"] == 0)))

    print("=" * 78)
    print("COCA-COLA SALES ANALYSIS  |  reproduction of the delivered workbook")
    print("=" * 78)
    print(f"Records                {total['records']:,}")
    print(f"Date range             {min(r['date'] for r in rows):%Y-%m-%d} to {max(r['date'] for r in rows):%Y-%m-%d}")
    print(f"Total sales            {money(total['sales'])}")
    print(f"Units sold             {total['units']:,}")
    print(f"Operating profit       {money(total['profit'])}")
    print(f"Weighted margin        {pct(total['margin'])}")
    print(f"Realised sales / unit  {total['sales_per_unit']:.4f}")
    print(f"Unit-weighted price    {total['listed_price']:.4f}")
    print(f"Profit / unit          {total['profit_per_unit']:.4f}")

    # ---- dimension tables
    for label, key in (("RETAILER", "retailer"), ("REGION", "region"), ("BRAND", "brand"),
                       ("MONTH", "year_month")):
        print(f"\n{label}")
        print(f"  {'Segment':<16}{'Sales':>13}{'Units':>12}{'Profit':>13}{'Margin':>9}{'SalesSh':>9}{'ProfSh':>9}")
        table = sorted(((k, agg(v)) for k, v in by(rows, key).items()), key=lambda kv: -kv[1]["sales"])
        for name, a in table:
            print(f"  {name:<16}{money(a['sales']):>13}{a['units']:>12,}{money(a['profit']):>13}"
                  f"{pct(a['margin']):>9}{a['sales'] / total['sales']:>9.1%}{a['profit'] / total['profit']:>9.1%}")

    print("\nMONTH-OVER-MONTH")
    ordered = [agg(v) for _, v in sorted(by(rows, "year_month").items())]
    names = sorted(by(rows, "year_month"))
    prev = None
    for name, a in zip(names, ordered):
        change = f"{a['sales'] / prev - 1:+.1%}" if prev else "-"
        print(f"  {name}  sales {money(a['sales']):>12}  units {a['units']:>10,}  margin {pct(a['margin']):>6}  "
              f"MoM {change:>7}")
        prev = a["sales"]

    # ---- retailers vs regions confounding
    print("\nRETAILER / REGION CONFOUNDING")
    for name, a in sorted(by(rows, "retailer").items()):
        regs = sorted({r["region"] for r in a})
        locs = len({r["location"] for r in a})
        print(f"  {name:<10} regions {regs}  locations {locs}")

    # ---- price and volume
    r_pooled = pearson([r["price"] for r in rows], [r["units"] for r in rows])
    print(f"\nPRICE-VOLUME  pooled r = {r_pooled:.3f}  (association only, not elasticity)")
    for b, group in sorted(by(rows, "brand").items()):
        print(f"  {b:<14} r = {pearson([r['price'] for r in group], [r['units'] for r in group]):.3f}")

    # ---- concentration
    locs = sorted(((k, agg(v)["sales"]) for k, v in by(rows, "location").items()), key=lambda kv: -kv[1])
    print("\nCONCENTRATION")
    for k in (5, 10, 20):
        print(f"  top {k:>2} of {len(locs)} cities: {sum(v for _, v in locs[:k]) / total['sales']:.1%} of sales")
    material = [n for n, v in locs if v >= 100000]
    print(f"  cities >= $100,000 annual sales: {len(material)} of {len(locs)} "
          f"({sum(v for _, v in locs if v >= 100000) / total['sales']:.1%} of sales)")

    # ---- checks
    print("\nVALIDATION")
    failed = 0
    for name, expected, actual in checks:
        ok = expected == actual
        failed += 0 if ok else 1
        print(f"  [{'PASS' if ok else 'FAIL'}] {name}: expected {expected}, actual {actual}")

    if args.json:
        payload = dict(total=total, records=total["records"],
                       checks=[dict(check=c, expected=e, actual=a) for c, e, a in checks],
                       correlation=r_pooled,
                       retailers={k: agg(v) for k, v in by(rows, "retailer").items()},
                       regions={k: agg(v) for k, v in by(rows, "region").items()},
                       brands={k: agg(v) for k, v in by(rows, "brand").items()})
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, indent=1, default=str)
        print(f"\nJSON written to {args.json}")

    print("\nDone." if not failed else f"\n{failed} CHECK(S) FAILED.")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
