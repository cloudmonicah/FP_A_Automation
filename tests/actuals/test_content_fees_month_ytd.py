"""Isolated CBL Content Fees monthly/YTD Tableau capture; original TWB untouched."""

from __future__ import annotations
import argparse
import csv
import re
import time
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from uuid import uuid4


def next_month(d):
    return date(d.year + (d.month == 12), d.month % 12 + 1, 1)


def prepare(source, destination, tab, period, basis):
    tree = ET.parse(source)
    root = tree.getroot()
    sheets = [s for s in root.findall("./worksheets/worksheet") if s.get("name") == tab]
    if len(sheets) != 1:
        raise ValueError(f"Expected one worksheet {tab!r}; found {len(sheets)}")
    sheet = sheets[0]
    fiscal_start = date(period.year - (period.month < 7), 7, 1)
    lower = date(period.year, period.month, 1) if basis == "MONTH" else fiscal_start
    upper = next_month(period)
    start = upper
    key = lambda x: re.sub(r"[^a-z0-9]", "", str(x or "").lower())
    parameter_names = {
        c.get("name")
        for c in sheet.iter("column")
        if key(c.get("caption")) == "forecaststart"
    }
    if len(parameter_names) != 1:
        raise ValueError(
            "Cannot identify unique Forecast Start parameter; refusing to run"
        )
    edits = 0
    for column in root.iter("column"):
        if (
            column.get("name") in parameter_names
            and "param-domain-type" in column.attrib
        ):
            column.set("value", f"#{start:%Y-%m-%d}#")
            calc = column.find("calculation")
            if calc is not None:
                calc.set("formula", f"#{start:%Y-%m-%d}#")
            edits += 1
    if not edits:
        raise ValueError("Forecast Start parameter not updated; refusing to run")
    dependencies = [
        d
        for d in sheet.findall(".//datasource-dependencies")
        if any(c.get("caption") == "Actual YTD" for c in d.findall("column"))
    ]
    if len(dependencies) != 1:
        raise ValueError("Cannot identify unique Actual YTD datasource dependency")
    actual_cols = [
        c for c in dependencies[0].findall("column") if c.get("caption") == "Actual YTD"
    ]
    if len(actual_cols) != 1:
        raise ValueError("Actual YTD column ambiguous")
    actual_name = actual_cols[0].get("name")
    formula = f"IF [GL Posting Date] >= #{lower:%Y-%m-%d}# AND [GL Posting Date] < #{upper:%Y-%m-%d}# THEN [Actual] ELSE 0 END"
    calc_edits = 0
    for col in root.iter("column"):
        if col.get("name") == actual_name:
            calc = col.find("calculation")
            if calc is not None:
                calc.set("formula", formula)
                calc_edits += 1
    if not calc_edits:
        raise ValueError("Actual YTD calculation not updated; refusing to run")
    filters = [
        f
        for f in sheet.findall(".//filter")
        if f.get("column", "").endswith("[my:GL Posting Date:ok]")
    ]
    if len(filters) != 1:
        raise ValueError(f"Expected one month date filter; found {len(filters)}")
    filt = filters[0]
    for child in list(filt):
        filt.remove(child)
    union = ET.SubElement(filt, "groupfilter", {"function": "union"})
    cursor = lower
    while cursor < upper:
        ET.SubElement(
            union,
            "groupfilter",
            {
                "function": "member",
                "level": "[my:GL Posting Date:ok]",
                "member": f"{cursor:%Y%m}",
            },
        )
        cursor = next_month(cursor)
    tree.write(destination, encoding="utf-8", xml_declaration=True)
    print(
        f"{basis}: date window [{lower}, {upper}), parameter edits={edits}, actual formula edits={calc_edits}"
    )
    print(f"{basis}: formula: {formula}")


def capture(source, tab, period, basis, output, wait):
    from automations.actuals.tableau_forecast_service import open_tableau_workbook
    from automations.actuals.tableau_desktop_export_service import (
        get_tableau_crosstab_text,
    )
    from automations.actuals.tableau_desktop_close_service import (
        close_tableau_forecast_workbook,
    )

    temp = source.with_name(f"{source.stem}_Individual_{basis}_{uuid4().hex[:8]}.twb")
    opened = False
    closed = False
    try:
        prepare(source, temp, tab, period, basis)
        open_tableau_workbook(tableau_workbook_path=str(temp))
        opened = True
        time.sleep(wait)
        raw = get_tableau_crosstab_text(
            tableau_workbook_path=str(temp), worksheet_name=tab
        )
        if not raw or not str(raw).strip():
            raise RuntimeError(f"{basis}: empty crosstab")
        path = output / f"CBL_Programming_fees_{period:%Y-%m}_{basis}.tsv"
        path.write_text(str(raw), encoding="utf-8-sig")
        rows = list(csv.DictReader(str(raw).splitlines(), delimiter="\t"))
        print(f"{basis}: raw saved: {path}")
        print(f"{basis}: rows: {len(rows)}; headers: {list(rows[0]) if rows else []}")
        if (
            len(rows) != 1
            or "Actual YTD" not in rows[0]
            or "GL Account Mapping" not in rows[0]
        ):
            print(f"{basis}: REVIEW REQUIRED: unexpected crosstab structure")
            return None
        if rows[0]["GL Account Mapping"].strip().upper() != "CONTENT FEES":
            print(
                f'{basis}: REVIEW REQUIRED: unexpected category {rows[0]["GL Account Mapping"]!r}'
            )
            return None
        amount = float(rows[0]["Actual YTD"].replace(",", "").replace("$", "").strip())
        print(f"{basis}: CONTENT FEES Actual YTD measure = {amount:,.2f}")
        return amount
    finally:
        if opened:
            try:
                close_tableau_forecast_workbook(tableau_workbook_path=str(temp))
                closed = True
            except Exception as exc:
                print(f"{basis}: CLOSE WARNING: {exc}")
        if not opened or closed:
            temp.unlink(missing_ok=True)
        else:
            print(f"{basis}: temporary workbook retained for inspection: {temp}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--month", default="2026-08")
    parser.add_argument("--output", default=r"C:\Temp\DirectCostTest\Individual")
    parser.add_argument("--wait", type=int, default=12)
    args = parser.parse_args()
    from automations.actuals.mapping import INDIVIDUAL_DIRECT_COST_MAPPINGS

    routes = [
        r
        for r in INDIVIDUAL_DIRECT_COST_MAPPINGS
        if r["company"] == "CBL" and r["fs_line"].casefold() == "programming fees"
    ]
    if len(routes) != 1:
        raise ValueError(
            f"Expected one CBL Programming fees route, found {len(routes)}"
        )
    route = routes[0]
    source = Path(route["individual_tableau_workbook"])
    if not source.is_file():
        raise FileNotFoundError(source)
    tab = route["individual_tableau_worksheet"]
    year, month = map(int, args.month.split("-"))
    period = date(year, month, 1)
    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)
    print("SOURCE:", source, "\nWORKSHEET:", tab)
    print(
        "IMPORTANT: Other source filters (including excluded GL Business Code MD) remain unchanged."
    )
    for basis, reference in [("MONTH", 1219837), ("YTD", 2483883)]:
        print("\n" + "=" * 60 + "\n" + basis)
        amount = capture(source, tab, period, basis, output, args.wait)
        if period == date(2026, 8, 1) and amount is not None:
            print(
                f"{basis}: consolidated reference = {reference:,.2f}; difference = {amount-reference:,.2f}"
            )
            print(
                "Comparison is diagnostic only; differing source filters may prevent reconciliation."
            )


if __name__ == "__main__":
    main()
