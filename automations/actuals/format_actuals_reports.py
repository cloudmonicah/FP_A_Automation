"""Format completed OPEX / Direct Cost actuals validation reports.

Monthly: A:H
YTD:     I:P

Two variances per period:
1. Consolidated Tableau - Financial Statement
2. Individual Tableau - Consolidated Tableau

Original source workbook is never overwritten.
"""

from pathlib import Path
from copy import copy
import argparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TOLERANCE_000 = 1.0  # $1,000 tolerance, since values are in $000

HEADERS = [
    "Financial Statement Line",
    "Financial Statement ($000)",
    "Consolidated Tableau ($000)",
    "Individual Tableau ($000)",
    "Variance 1: FS vs Consolidated ($000)",
    "Variance 2: Consolidated vs Individual ($000)",
    "Monthly Status",
    "Exception / Notes",
    "Financial Statement ($000)",
    "YTD Consolidated Tableau ($000)",
    "YTD Individual Tableau ($000)",
    "Variance 1: FS vs Consolidated ($000)",
    "Variance 2: Consolidated vs Individual ($000)",
    "YTD Status",
    "Exception / Notes",
]

# The approved report has 15 populated columns, A:O.
# Monthly: A:H; YTD: I:O.


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def variance(left, right):
    if not is_number(left) or not is_number(right):
        return None
    return round(float(right) - float(left), 6)


def validation_status(fs, consolidated, individual, old_group, old_individual):
    v1 = variance(fs, consolidated)
    v2 = variance(consolidated, individual)

    if v1 is None:
        return "REVIEW REQUIRED"

    if abs(v1) > TOLERANCE_000:
        return "REVIEW REQUIRED"

    if str(old_group or "").strip().upper() in (
        "FAIL",
        "ERROR",
        "ACTION REQUIRED",
        "REVIEW REQUIRED",
    ):
        return "REVIEW REQUIRED"

    if v2 is None:
        return "PENDING"

    if abs(v2) > TOLERANCE_000:
        return "REVIEW REQUIRED"

    if str(old_individual or "").strip().upper() in (
        "FAIL",
        "ERROR",
        "ACTION REQUIRED",
        "REVIEW REQUIRED",
    ):
        return "REVIEW REQUIRED"

    if str(old_individual or "").strip().upper() != "PASS":
        return "PENDING"

    return "PASS"


def overall_status(statuses):
    if any(value == "REVIEW REQUIRED" for value in statuses):
        return "REVIEW REQUIRED"
    if not statuses or any(value != "PASS" for value in statuses):
        return "PENDING"
    return "PASS"


def style_status(cell):
    status = str(cell.value or "").upper()

    if status == "PASS":
        fill, color = "E2F0D9", "006100"
    elif status == "REVIEW REQUIRED":
        fill, color = "F4CCCC", "9C0006"
    else:
        fill, color = "FFF2CC", "9C6500"

    cell.fill = PatternFill(fill_type="solid", fgColor=fill)
    cell.font = Font(bold=True, color=color)
    cell.alignment = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )


def build_note(original_note, v1, v2, individual_status):
    parts = []

    if original_note:
        note = str(original_note).strip()
        if note and note.lower() != "reconciled.":
            parts.append(note)

    if v1 is None:
        parts.append("FS / consolidated comparison unavailable")
    elif abs(v1) > TOLERANCE_000:
        parts.append(f"FS vs consolidated difference: {v1:,.3f} ($000)")

    if v2 is None:
        parts.append("Individual Tableau not validated")
    elif abs(v2) > TOLERANCE_000:
        parts.append(f"Consolidated vs individual difference: {v2:,.3f} ($000)")

    if (
        v2 is not None
        and str(individual_status or "").upper() in ("FAIL", "ERROR", "REVIEW REQUIRED")
        and abs(v2) <= TOLERANCE_000
    ):
        parts.append("Existing individual validation flagged an exception")

    return " | ".join(parts) if parts else "Reconciled."


def prepare_template_sheet(ws):
    # Unmerge existing template ranges before changing the layout.
    for merged_range in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(merged_range))

    # Clear template sample values only.
    for row in ws.iter_rows(
        min_row=7,
        max_row=max(ws.max_row, 25),
        min_col=1,
        max_col=15,
    ):
        for cell in row:
            cell.value = None

    # Rebuild the approved header area.
    ws.merge_cells("A1:M1")
    ws.merge_cells("A2:B2")
    ws.merge_cells("A3:B3")
    ws.merge_cells("A4:G4")

    ws.merge_cells("B6:G6")
    ws.merge_cells("H6:M6")

    ws["B6"] = "Monthly Actuals"
    ws["H6"] = "YTD Actuals"

    for address in ("B6", "H6"):
        cell = ws[address]
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
        )
        cell.font = Font(
            name="Aptos",
            size=12,
            bold=True,
            color="FFFFFF",
        )
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="244062",
        )

    for col, heading in enumerate(HEADERS, 1):
        cell = ws.cell(7, col)
        cell.value = heading
        cell.font = Font(
            name="Aptos",
            bold=True,
            color="FFFFFF",
        )
        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="244062",
        )
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    ws.row_dimensions[6].height = 25
    ws.row_dimensions[7].height = 60

    widths = {
        "A": 33,
        "B": 19,
        "C": 21,
        "D": 21,
        "E": 21,
        "F": 21,
        "G": 19,
        "H": 19,
        "I": 21,
        "J": 21,
        "K": 21,
        "L": 21,
        "M": 19,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "B8"
    ws.sheet_view.showGridLines = False


def clean_variance_formatting(ws):
    """Remove inherited variance fills and display rounded $000 values."""

    from openpyxl.styles import PatternFill, Alignment

    # Monthly and YTD financial values, including both variances
    financial_columns = (2, 3, 4, 5, 6, 9, 10, 11, 12, 13)

    # Only the four variance columns should have their fills cleared
    variance_columns = (5, 6, 12, 13)

    whole_number_format = "#,##0;[Red](#,##0);0"

    for row in range(8, ws.max_row + 1):

        for col in financial_columns:
            cell = ws.cell(row, col)

            # Display rounded values, without modifying source amounts
            cell.number_format = whole_number_format
            cell.alignment = Alignment(
                horizontal="right",
                vertical="center",
            )

        for col in variance_columns:
            cell = ws.cell(row, col)

            # Remove inherited green or other background fills
            cell.fill = PatternFill(fill_type=None)

            # Remove any existing conditional formatting effects
            # through the worksheet-level cleanup below.

    # The source template may have conditional formatting rules
    # that recolor entire variance columns.
    ws.conditional_formatting._cf_rules.clear()


def format_report(source, template=None, destination=None):
    source = Path(source)
    template = Path(template) if template is not None else None

    if destination is None:
        destination = source.with_name(source.stem + "_Template_Formatted_v2.xlsx")
    else:
        destination = Path(destination)

    if destination.resolve() == source.resolve():
        raise ValueError("Cannot overwrite the source workbook.")

    if not source.is_file():
        raise FileNotFoundError(f"Source not found: {source}")

    if template is not None and not template.is_file():
        raise FileNotFoundError(f"Template not found: {template}")

    old = load_workbook(source, data_only=False)

    layout = load_workbook(template) if template is not None else Workbook()

    try:
        # Only use the original template sheet.
        base = layout.active

        for sheet in list(layout):
            if sheet != base:
                layout.remove(sheet)

        results = {}

        for index, company in enumerate(("ALIV", "CBL")):
            matches = [
                sheet
                for sheet in old
                if sheet.title.startswith(company + " ")
                and sheet.title != "Executive Summary"
            ]

            if len(matches) != 1:
                raise ValueError(
                    f"Expected exactly one {company} source sheet; "
                    f"found {[s.title for s in matches]}"
                )

            original = matches[0]

            if original["A6"].value != "Financial Statement Line":
                raise ValueError(f"{original.title}: unexpected source column layout.")

            ws = base if index == 0 else layout.copy_worksheet(base)

            prepare_template_sheet(ws)
            ws.title = original.title

            ws["A1"] = f"{original.title.upper()} ACTUALS VALIDATION"
            ws["C2"] = original["C2"].value
            ws["C3"] = original["C3"].value

            statuses = []
            output_row = 8

            for source_row in range(7, original.max_row + 1):

                def x(column):
                    return original.cell(
                        row=source_row,
                        column=column,
                    ).value

                line = x(1)

                if not line:
                    continue

                # MONTHLY SOURCE VALUES
                monthly_fs = x(2)
                monthly_group = x(3)
                monthly_individual = x(12)

                monthly_v1 = variance(monthly_fs, monthly_group)
                monthly_v2 = variance(
                    monthly_group,
                    monthly_individual,
                )

                monthly_status = validation_status(
                    monthly_fs,
                    monthly_group,
                    monthly_individual,
                    x(5),
                    x(15),
                )

                monthly_note = build_note(
                    x(6),
                    monthly_v1,
                    monthly_v2,
                    x(15),
                )

                # YTD SOURCE VALUES
                ytd_fs = x(8)
                ytd_group = x(9)
                ytd_individual = x(16)

                ytd_v1 = variance(ytd_fs, ytd_group)
                ytd_v2 = variance(
                    ytd_group,
                    ytd_individual,
                )

                ytd_status = validation_status(
                    ytd_fs,
                    ytd_group,
                    ytd_individual,
                    x(11),
                    x(19),
                )

                ytd_note = build_note(
                    None,
                    ytd_v1,
                    ytd_v2,
                    x(19),
                )

                # A:O final report layout
                values = [
                    line,  # A
                    monthly_fs,  # B
                    monthly_group,  # C
                    monthly_individual,  # D
                    monthly_v1,  # E
                    monthly_v2,  # F
                    monthly_status,  # G
                    monthly_note,  # H
                    ytd_fs,  # I
                    ytd_group,  # J
                    ytd_individual,  # K
                    ytd_v1,  # L
                    ytd_v2,  # M
                    ytd_status,  # N
                    ytd_note,  # O
                ]

                for col, value in enumerate(values, 1):
                    cell = ws.cell(output_row, col)
                    cell.value = value

                    cell.alignment = Alignment(
                        vertical="center",
                        wrap_text=col in (1, 7, 8, 14, 15),
                    )

                    if col in (2, 3, 4, 5, 6, 9, 10, 11, 12, 13):
                        cell.number_format = "#,##0;[Red](#,##0);0"

                style_status(ws.cell(output_row, 7))
                style_status(ws.cell(output_row, 14))

                statuses.extend([monthly_status, ytd_status])

                output_row += 1

            ws["G2"] = overall_status(statuses)
            style_status(ws["G2"])

            ws.auto_filter.ref = f"A7:O{max(8, output_row - 1)}"

            clean_variance_formatting(ws)

            results[company] = {
                "rows": output_row - 8,
                "status": ws["G2"].value,
            }

        layout.active = 0

        destination.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        layout.save(destination)

        print("\nVALIDATION FORMAT RESULTS")
        for company, result in results.items():
            print(f"{company}: {result['rows']} rows | " f"{result['status']}")

        print("Sheets:", ", ".join(layout.sheetnames))

    finally:
        old.close()
        layout.close()

    return destination


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("source")
    parser.add_argument("--template", default=None)
    parser.add_argument("--output", default=None)

    args = parser.parse_args()

    output = format_report(
        args.source,
        args.template,
        args.output,
    )

    print("Created:", output)
