"""Test individual Direct Cost Tableau crosstab extraction."""

from __future__ import annotations

import argparse
import csv
import re
import time
from datetime import date
from pathlib import Path
from uuid import uuid4
import xml.etree.ElementTree as ET


def main():
    parser = argparse.ArgumentParser()

    parser.add_argument("--company", default="CBL")
    parser.add_argument("--fs-line", default="Programming fees")
    parser.add_argument("--month", default="2026-08")
    parser.add_argument(
        "--output",
        default=r"C:\Temp\DirectCostTest\Individual",
    )
    parser.add_argument("--wait", type=int, default=12)

    args = parser.parse_args()

    from automations.actuals.mapping import (
        INDIVIDUAL_DIRECT_COST_MAPPINGS,
    )

    matches = [
        route
        for route in INDIVIDUAL_DIRECT_COST_MAPPINGS
        if route["company"] == args.company
        and route["fs_line"].casefold() == args.fs_line.casefold()
    ]

    if len(matches) != 1:
        raise ValueError(
            f"Expected one mapping for "
            f"{args.company}/{args.fs_line}, "
            f"found {len(matches)}"
        )

    route = matches[0]

    source = Path(route["individual_tableau_workbook"])
    worksheet = route["individual_tableau_worksheet"]

    if not source.is_file():
        raise FileNotFoundError(source)

    output = Path(args.output)
    output.mkdir(parents=True, exist_ok=True)

    year, month = map(int, args.month.split("-"))

    if not 1 <= month <= 12:
        raise ValueError("Invalid month")

    forecast_start = date(
        year + int(month == 12),
        month % 12 + 1,
        1,
    )

    temp = source.with_name(f"{source.stem}_IndividualTest_{uuid4().hex[:8]}.twb")

    tree = ET.parse(source)
    root = tree.getroot()

    sheets = [
        sheet
        for sheet in root.findall("./worksheets/worksheet")
        if sheet.get("name") == worksheet
    ]

    if len(sheets) != 1:
        raise ValueError(f"Worksheet missing or ambiguous: {worksheet}")

    sheet = sheets[0]

    names = {
        column.get("name")
        for column in sheet.iter("column")
        if re.sub(
            r"[^a-z0-9]",
            "",
            str(column.get("caption", "")).lower(),
        )
        == "forecaststart"
    }

    changed = 0

    if len(names) == 1:
        for column in root.iter("column"):
            if column.get("name") in names and "param-domain-type" in column.attrib:
                column.set(
                    "value",
                    f"#{forecast_start:%Y-%m-%d}#",
                )

                calculation = column.find("calculation")

                if calculation is not None:
                    calculation.set(
                        "formula",
                        f"#{forecast_start:%Y-%m-%d}#",
                    )

                changed += 1

    tree.write(
        temp,
        encoding="utf-8",
        xml_declaration=True,
    )

    print("SOURCE:", source, flush=True)
    print("WORKSHEET:", worksheet, flush=True)
    print("FORECAST START:", forecast_start, flush=True)
    print("PARAMETER EDITS:", changed, flush=True)
    print(
        "NOTE: Existing company, GL and date filters " "are not modified.",
        flush=True,
    )

    opened = False
    closed = False

    try:
        from automations.forecast.tableau_service import (
            open_tableau_workbook,
        )

        from automations.actuals.tableau_desktop_export_service import (
            get_tableau_crosstab_text,
        )

        open_tableau_workbook(tableau_workbook_path=str(temp))

        opened = True

        if args.wait:
            time.sleep(args.wait)

        text = get_tableau_crosstab_text(
            tableau_workbook_path=str(temp),
            worksheet_name=worksheet,
        )

        if not text or not str(text).strip():
            raise RuntimeError("Empty Tableau crosstab")

        filename = (
            f"{args.company}_"
            f"{re.sub(r'[^A-Za-z0-9]+', '_', args.fs_line)}_"
            f"{args.month}_raw.tsv"
        )

        destination = output / filename

        destination.write_text(
            str(text),
            encoding="utf-8-sig",
        )

        print("RAW CROSSTAB:", destination, flush=True)

        rows = list(
            csv.reader(
                str(text).splitlines(),
                delimiter="\t",
            )
        )

        print("ROW COUNT:", len(rows), flush=True)
        print("FIRST 3 ROWS:", flush=True)

        for index, row in enumerate(rows[:3], 1):
            print(
                f"ROW {index}:",
                repr(row[:20]),
                flush=True,
            )

        print(
            "SUCCESS: Raw Tableau capture completed.",
            flush=True,
        )

    finally:
        if opened:
            try:
                from automations.actuals.tableau_desktop_close_service import (
                    close_tableau_forecast_workbook,
                )

                close_tableau_forecast_workbook(tableau_workbook_path=str(temp))

                closed = True

            except Exception as exc:
                print("CLOSE WARNING:", exc, flush=True)

        if not opened or closed:
            temp.unlink(missing_ok=True)
        else:
            print(
                "Temporary workbook retained because " "closing was not confirmed:",
                temp,
                flush=True,
            )


if __name__ == "__main__":
    main()
