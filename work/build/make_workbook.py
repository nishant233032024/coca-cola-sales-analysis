"""Assemble the base workbook (no pivots yet)."""
import os

import openpyxl

import build_base as B
import build_base2 as B2

OUT = B.OUT


def main():
    F = B.F
    B2.FACTS = F
    B2.GLOBAL_FACTS = F

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    B.s_readme(wb)
    B.s_raw(wb)
    B.s_clean(wb)
    B.s_dictionary(wb)
    B.s_quality(wb)
    B.s_kpi(wb)
    B.s_retailer(wb)
    ws, first_state, states = B.s_region(wb)
    B.s_region_states(ws, first_state, states)
    B.s_brand(wb)
    B.s_monthly(wb)
    B.s_price(wb)

    B2.s_pivot_placeholder(wb)
    B2.s_dashboard(wb)
    B2.s_insights(wb)
    B2.s_implementation(wb)
    B2.s_validation(wb)
    B2.s_script(wb)

    wb.properties.title = "Coca-Cola Sales Analysis 2021"
    wb.properties.creator = "Nishant"
    wb.properties.description = ("Excel sales and operating-profitability analysis of the supplied Coca-Cola "
                                "dataset. Educational dataset, not audited company reporting.")
    wb.active = 0
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print("saved", OUT, os.path.getsize(OUT), "bytes", len(wb.sheetnames), "sheets")


if __name__ == "__main__":
    main()
