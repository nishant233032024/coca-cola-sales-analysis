"""Single source of truth for every number used in the Coca-Cola project.

Reads work/source.xlsx (the supplied attachment, unchanged) and writes facts.json.
No hard-coded business numbers anywhere else in the build.
"""
from __future__ import annotations

import json
import math
import os
from collections import defaultdict

import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
SOURCE = os.path.join(HERE, "..", "source.xlsx")
DATA_START, DATA_END = 6, 3893
HEADER_ROW = 5


def load_rows():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    ws = wb["Data"]
    headers = [ws.cell(HEADER_ROW, c).value for c in range(2, 14)]
    rows = []
    for r in range(DATA_START, DATA_END + 1):
        vals = [ws.cell(r, c).value for c in range(2, 14)]
        if all(v is None for v in vals):
            continue
        (ret, rid, dt, reg, st, city, brand, price, units, sales, profit, margin) = vals
        rows.append(
            dict(
                source_row=r,
                retailer=ret,
                retailer_id=str(rid),
                date=dt,
                region=reg,
                state=st,
                city=city,
                location=f"{st} | {city}",
                brand=brand,
                price=float(price),
                units=int(units),
                sales=float(sales),
                profit=float(profit),
                source_margin=float(margin),
                year=dt.year,
                month=dt.month,
                month_start=dt.replace(day=1),
                year_month=f"{dt.year}-{dt.month:02d}",
                quarter=(dt.month - 1) // 3 + 1,
            )
        )
    return headers, rows


def agg(rows):
    s = sum(r["sales"] for r in rows)
    u = sum(r["units"] for r in rows)
    p = sum(r["profit"] for r in rows)
    listed = sum(r["price"] * r["units"] for r in rows)
    return dict(
        records=len(rows),
        sales=s,
        units=u,
        profit=p,
        margin=(p / s if s else None),
        sales_per_unit=(s / u if u else None),
        profit_per_unit=(p / u if u else None),
        listed_price=(listed / u if u else None),
    )


def group(rows, key):
    g = defaultdict(list)
    for r in rows:
        g[r[key]].append(r)
    return g


def pearson(xs, ys):
    n = len(xs)
    sx, sy = sum(xs), sum(ys)
    sxy = sum(a * b for a, b in zip(xs, ys))
    sx2, sy2 = sum(a * a for a in xs), sum(b * b for b in ys)
    num = n * sxy - sx * sy
    den = math.sqrt((n * sx2 - sx * sx) * (n * sy2 - sy * sy))
    return num / den if den else None


def build():
    headers, rows = load_rows()
    total = agg(rows)
    total["margin"] = total["profit"] / total["sales"]

    facts = {
        "filename": "Coke-sales-analysis-stubs (1).xlsx",
        "sheet_used": "Data",
        "header_row": HEADER_ROW,
        "headers": headers,
        "data_rows": [DATA_START, DATA_END],
        "records": len(rows),
        "date_min": min(r["date"] for r in rows).strftime("%Y-%m-%d"),
        "date_max": max(r["date"] for r in rows).strftime("%Y-%m-%d"),
        "years": sorted({r["year"] for r in rows}),
        "total": total,
        "counts": {
            "retailers": len({r["retailer"] for r in rows}),
            "regions": len({r["region"] for r in rows}),
            "states": len({r["state"] for r in rows}),
            "cities": len({r["location"] for r in rows}),
            "brands": len({r["brand"] for r in rows}),
            "months": len({r["year_month"] for r in rows}),
        },
        "levels": {},
        "confounding": {},
    }

    # dimension tables
    for name, key, order in (
        ("retailer", "retailer", None),
        ("region", "region", None),
        ("brand", "brand", None),
        ("month", "year_month", None),
        ("state", "state", None),
        ("location", "location", None),
    ):
        g = group(rows, key)
        out = []
        for k, v in g.items():
            a = agg(v)
            a["name"] = k
            a["sales_share"] = a["sales"] / total["sales"]
            a["profit_share"] = a["profit"] / total["profit"]
            a["locations"] = len({r["location"] for r in v})
            a["retailers"] = sorted({r["retailer"] for r in v})
            out.append(a)
        out.sort(key=lambda d: -d["sales"])
        facts["levels"][name] = out

    # retailer -> regions / locations (confounding evidence)
    for ret in sorted({r["retailer"] for r in rows}):
        sel = [r for r in rows if r["retailer"] == ret]
        facts["confounding"][ret] = dict(
            regions=sorted({r["region"] for r in sel}),
            locations=len({r["location"] for r in sel}),
            states=sorted({r["state"] for r in sel}),
        )
    region_retailers = {}
    for reg in sorted({r["region"] for r in rows}):
        region_retailers[reg] = sorted({r["retailer"] for r in rows if r["region"] == reg})
    facts["region_retailers"] = region_retailers

    # monthly with MoM and pp change
    months = sorted(facts["levels"]["month"], key=lambda d: d["name"])
    prev = None
    prevm = None
    for m in months:
        m["mom"] = (m["sales"] / prev - 1) if prev else None
        m["margin_pp"] = (m["margin"] - prevm) * 100 if prevm is not None else None
        prev = m["sales"]
        prevm = m["margin"]
    facts["months"] = months
    hi = max(months, key=lambda d: d["sales"])
    lo = min(months, key=lambda d: d["sales"])
    facts["month_extremes"] = dict(
        high=hi["name"], high_sales=hi["sales"],
        low=lo["name"], low_sales=lo["sales"],
        diff=hi["sales"] - lo["sales"],
        rel=hi["sales"] / lo["sales"] - 1,
        high_units_month=max(months, key=lambda d: d["units"])["name"],
        best_margin_month=max(months, key=lambda d: d["margin"])["name"],
        worst_margin_month=min(months, key=lambda d: d["margin"])["name"],
        margin_range_pp=(max(m["margin"] for m in months) - min(m["margin"] for m in months)) * 100,
    )

    # brand x region matrix
    matrix = {}
    for b in sorted({r["brand"] for r in rows}):
        matrix[b] = {}
        for reg in facts["levels"]["region"]:
            sel = [r for r in rows if r["brand"] == b and r["region"] == reg["name"]]
            if sel:
                matrix[b][reg["name"]] = agg(sel)
    facts["brand_region"] = matrix

    # retailer x region
    rr = {}
    for ret in sorted({r["retailer"] for r in rows}):
        rr[ret] = {}
        for reg in sorted({r["region"] for r in rows}):
            sel = [r for r in rows if r["retailer"] == ret and r["region"] == reg]
            if sel:
                rr[ret][reg] = agg(sel)
    facts["retailer_region"] = rr

    # region x month sales matrix
    rm = {}
    for reg in sorted({r["region"] for r in rows}):
        rm[reg] = {}
        for mo in range(1, 13):
            sel = [r for r in rows if r["region"] == reg and r["month"] == mo]
            rm[reg][f"{mo:02d}"] = agg(sel)["sales"] if sel else None
    facts["region_month"] = rm

    # concentration
    locs = facts["levels"]["location"]
    cum = 0.0
    for i, l in enumerate(locs, 1):
        cum += l["sales"]
        l["cum_share"] = cum / total["sales"]
    facts["concentration"] = {
        "top5": sum(x["sales"] for x in locs[:5]) / total["sales"],
        "top10": sum(x["sales"] for x in locs[:10]) / total["sales"],
        "top20": sum(x["sales"] for x in locs[:20]) / total["sales"],
    }

    # materiality rule: >= $100,000 annual sales (54 locations)
    material = [l for l in locs if l["sales"] >= 100000]
    facts["materiality"] = dict(threshold=100000, count=len(material), total=len(locs),
                                sales_share=sum(l["sales"] for l in material) / total["sales"])

    tm = total["margin"]
    facts["high_sales_low_margin"] = [
        {k: l[k] for k in ("name", "sales", "margin", "profit", "units", "retailers")}
        for l in locs if l["sales"] >= 100000 and l["margin"] < tm
    ]
    facts["low_sales_high_margin"] = [
        {k: l[k] for k in ("name", "sales", "margin", "profit")}
        for l in locs if l["sales"] < 100000 and l["margin"] > tm
    ]

    # identity checks
    sales_breaks = [r for r in rows if abs(r["price"] * r["units"] - r["sales"]) > 0.01]
    margin_breaks = [
        r for r in rows
        if r["sales"] and abs(r["profit"] / r["sales"] - r["source_margin"]) > 0.0005
    ]
    facts["quality"] = dict(
        sales_identity_breaks=len(sales_breaks),
        margin_identity_breaks=len(margin_breaks),
        negatives=sum(1 for r in rows if r["sales"] < 0 or r["profit"] < 0 or r["units"] < 0),
        zero_sales=[dict(location=r["location"], brand=r["brand"], date=r["date"].strftime("%Y-%m-%d"),
                         retailer=r["retailer"], supplied_margin=r["source_margin"])
                    for r in rows if r["sales"] == 0],
        exact_dupes=len(rows) - len({(r["retailer"], r["retailer_id"], r["date"], r["region"], r["state"],
                                      r["city"], r["brand"], round(r["price"], 6), r["units"],
                                      round(r["sales"], 4), round(r["profit"], 4),
                                      round(r["source_margin"], 6)) for r in rows}),
        key_dupes=len(rows) - len({(r["retailer"], r["location"], r["brand"], r["month_start"]) for r in rows}),
        price_min=min(r["price"] for r in rows),
        price_max=max(r["price"] for r in rows),
        margin_min=min(r["source_margin"] for r in rows),
        margin_max=max(r["source_margin"] for r in rows),
        full_months=all(len([r for r in rows if r["month"] == m]) == 324 for m in range(1, 13)),
        records_per_month=sorted({len([r for r in rows if r["month"] == m]) for m in range(1, 13)}),
    )

    # simple average of row margins (the wrong way) for the teaching contrast
    facts["row_margin_simple_avg"] = sum(r["source_margin"] for r in rows) / len(rows)

    # price <-> volume association
    pooled = pearson([r["price"] for r in rows], [r["units"] for r in rows])
    by_brand = {}
    for b in sorted({r["brand"] for r in rows}):
        sel = [r for r in rows if r["brand"] == b]
        by_brand[b] = pearson([r["price"] for r in sel], [r["units"] for r in sel])
    facts["correlation"] = dict(pooled=pooled, by_brand=by_brand,
                                n=len(rows),
                                brand_n=len([r for r in rows if r["brand"] == "Coca-Cola"]))

    # price band table (unit-weighted)
    bands = [("<= $0.30", 0, 0.30), ("$0.31-$0.45", 0.30001, 0.45),
             ("$0.46-$0.60", 0.45001, 0.60), ("$0.61-$0.75", 0.60001, 0.75),
             ("$0.76-$0.90", 0.75001, 0.90), ("> $0.90", 0.90001, 99)]
    pb = []
    for label, lo_, hi_ in bands:
        sel = [r for r in rows if lo_ <= r["price"] <= hi_]
        a = agg(sel)
        a["label"] = label
        pb.append(a)
    facts["price_bands"] = pb

    # brand x month sales
    bm = {}
    for b in sorted({r["brand"] for r in rows}):
        bm[b] = {f"{m:02d}": agg([r for r in rows if r["brand"] == b and r["month"] == m])["sales"]
                 for m in range(1, 13)}
    facts["brand_month"] = bm

    # Coca-Cola brand vs portfolio
    coke = [r for r in rows if r["brand"] == "Coca-Cola"]
    facts["coke_brand"] = agg(coke)
    facts["coke_brand"]["share_of_portfolio_sales"] = facts["coke_brand"]["sales"] / total["sales"]
    facts["coke_brand"]["share_of_portfolio_profit"] = facts["coke_brand"]["profit"] / total["profit"]

    # revenue decomposition Jan -> Dec
    jan = months[0]
    dec = months[-1]
    facts["decomp"] = dict(
        jan=jan["name"], dec=dec["name"],
        sales_growth=dec["sales"] / jan["sales"] - 1,
        units_growth=dec["units"] / jan["units"] - 1,
        price_growth=dec["sales_per_unit"] / jan["sales_per_unit"] - 1,
        jan_rpu=jan["sales_per_unit"], dec_rpu=dec["sales_per_unit"],
        jan_units=jan["units"], dec_units=dec["units"],
    )

    with open(os.path.join(HERE, "..", "facts.json"), "w", encoding="utf-8") as f:
        json.dump(facts, f, indent=1, default=str)
    return facts


if __name__ == "__main__":
    f = build()
    t = f["total"]
    print("records", f["records"], "sales", t["sales"], "units", t["units"], "profit", t["profit"])
    print("margin", round(t["margin"], 6), "spu", round(t["sales_per_unit"], 6))
    print("month extremes", f["month_extremes"])
    print("decomp", f["decomp"])
    print("corr", f["correlation"]["pooled"])
    print("materiality", f["materiality"])
    print("quality", {k: v for k, v in f["quality"].items() if k != "zero_sales"})
