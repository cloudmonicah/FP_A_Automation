"""Shared actuals report layout and period preparation for group captures."""

from pathlib import Path
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from copy import copy
import re
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment

EXTRA_HEADERS = [
    "YTD Statement ($000)",
    "YTD Group Tableau ($000)",
    "YTD Group - Statement ($000)",
    "YTD Group Status",
    "Monthly Individual ($000)",
    "Monthly Individual - Statement ($000)",
    "Monthly Individual - Group ($000)",
    "Monthly Individual Status",
    "YTD Individual ($000)",
    "YTD Individual - Statement ($000)",
    "YTD Individual - Group ($000)",
    "YTD Individual Status",
]


def canonical(value):
    return " ".join(str(value or "").upper().replace("&", " AND ").split())


def rounded(value):
    return int(Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def add_company_columns(sheet):
    from openpyxl.utils import get_column_letter

    for column, label in enumerate(EXTRA_HEADERS, 8):
        cell = sheet.cell(6, column, label)
        cell.fill = PatternFill("solid", fgColor="145A70")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        sheet.column_dimensions[get_column_letter(column)].width = 21
    for cell in sheet[6]:
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
    sheet.row_dimensions[6].height = 55
    sheet.auto_filter.ref = f"A6:S{sheet.max_row}"
    sheet.freeze_panes = "B7"
    for row in range(7, sheet.max_row + 1):
        for column in (11, 15, 19):
            if sheet.cell(row, column).value is None:
                sheet.cell(row, column, "NOT VALIDATED")
        for column in (8, 9, 10, 12, 13, 14, 16, 17, 18):
            sheet.cell(row, column).number_format = "#,##0;[Red](#,##0);–"


def status_cell(cell, status):
    cell.value = status
    cell.font = Font(
        bold=True,
        color=(
            "006100"
            if status == "PASS"
            else "9C0006" if status in ("FAIL", "ACTION REQUIRED") else "666666"
        ),
    )
    cell.fill = PatternFill(
        "solid",
        fgColor=(
            "E2EFDA"
            if status == "PASS"
            else "F4CCCC" if status in ("FAIL", "ACTION REQUIRED") else "EEEEEE"
        ),
    )


def write_area_workbook(
    *, area, output_file, closed_period, forecast_start, aliv_result, cbl_result
):
    # All areas use the same OPEX formatter and its company/detail result schema.
    from automations.actuals.opex_validation_service import (
        write_opex_validation_workbook,
    )

    for result in (aliv_result, cbl_result):
        for item in result["details"]:
            item.setdefault("tableau_category", item["statement_line"])
    path = write_opex_validation_workbook(
        output_file=output_file,
        closed_period=closed_period,
        forecast_start=forecast_start,
        aliv_result=aliv_result,
        cbl_result=cbl_result,
    )
    report = load_workbook(path)
    try:
        for company in ("ALIV", "CBL"):
            sheet = report[f"{company} OPEX"]
            sheet.title = f"{company} {area}"
            add_company_columns(sheet)
        for sheet in report:
            for cells in sheet:
                for cell in cells:
                    if isinstance(cell.value, str) and "OPEX" in cell.value:
                        cell.value = cell.value.replace("OPEX", area)
        refresh_summary(report)
        report.save(path)
    finally:
        report.close()
    return path


def statement_ytd(path, closed_period):
    if not path:
        return {}
    from automations.actuals.adaptive_aliv_statement_reader import read_aliv_period

    return {
        k: rounded(v[1])
        for k, v in read_aliv_period(path, closed_period).items()
        if v is not None
    }


def merge_group_values(
    report_path,
    monthly,
    ytd,
    closed_period,
    area="OPEX",
    aliv_fs_path=None,
    cbl_fs_path=None,
):
    from automations.actuals.validation_config import (
        OPEX_MAPPINGS,
        DIRECT_COST_MAPPINGS,
        AGG_TOTAL,
    )

    mappings = OPEX_MAPPINGS if area == "OPEX" else DIRECT_COST_MAPPINGS
    fs_ytd = statement_ytd(aliv_fs_path, closed_period) if aliv_fs_path else {}
    from automations.actuals.cbl_statement_periods import read_cbl_area

    cbl_ytd = (
        {
            canonical(k): rounded(v)
            for k, v in read_cbl_area(cbl_fs_path, area, closed_period)["YTD"].items()
        }
        if cbl_fs_path
        else {}
    )
    report = load_workbook(report_path)
    try:
        for company in ("ALIV", "CBL"):
            sheet = report[f"{company} {area}"]
            add_company_columns(sheet)
            for row in range(7, sheet.max_row + 1):
                label = canonical(sheet.cell(row, 1).value)
                routes = [
                    m
                    for m in mappings
                    if m["company"] == company and canonical(m["fs_line"]) == label
                ]
                if len(routes) != 1:
                    status_cell(sheet.cell(row, 11), "ACTION REQUIRED")
                    continue
                mapping = routes[0]
                if area == "OPEX":
                    keys = [canonical(mapping["tableau_category"])]
                elif mapping["aggregation"] == AGG_TOTAL:
                    keys = [f"{company} TOTAL"]
                else:
                    keys = [canonical(k) for k in mapping.get("tableau_categories", [])]
                normalized = {canonical(k): v for k, v in ytd[company].items()}
                if not keys or any(k not in normalized for k in keys):
                    status_cell(sheet.cell(row, 11), "ACTION REQUIRED")
                    continue
                sheet.cell(
                    row,
                    9,
                    rounded(sum(Decimal(str(normalized[k])) for k in keys) / 1000),
                )
                statement_values = fs_ytd if company == "ALIV" else cbl_ytd
                if statement_values.get(label) is not None:
                    sheet.cell(row, 8, statement_values[label])
                    difference = sheet.cell(row, 9).value - statement_values[label]
                    sheet.cell(row, 10, difference)
                    status_cell(
                        sheet.cell(row, 11), "PASS" if difference == 0 else "FAIL"
                    )
                else:
                    status_cell(sheet.cell(row, 11), "NOT VALIDATED")
                    sheet.cell(row, 11).comment = Comment(
                        "Financial statement YTD layout is not confirmed for this company.",
                        "Validation",
                    )
        for name in ("Group Actuals", "Validation Summary"):
            if name in report.sheetnames:
                del report[name]
        report.save(report_path)
    finally:
        report.close()


def merge_individual_values(report, groups):
    """Write captures to either the intermediate A:S or approved A:O report."""
    for mapping, pair in groups:
        sheet = report[mapping["output_company_sheet"]]
        compact = str(sheet["A7"].value or "").strip() == "Financial Statement Line"
        legacy = str(sheet["A6"].value or "").strip() == "Financial Statement Line"
        if not (compact or legacy):
            raise ValueError(f"{sheet.title}: unrecognized validation layout")
        if legacy:
            add_company_columns(sheet)
        first_row = 8 if compact else 7
        matches = [
            r
            for r in range(first_row, sheet.max_row + 1)
            if canonical(sheet.cell(r, 1).value) == canonical(mapping["fs_line"])
        ]
        if len(matches) != 1:
            raise ValueError(
                f"Expected one row for {mapping['company']} / {mapping['fs_line']}; found {len(matches)}"
            )
        row = matches[0]
        month, ytd = pair
        if compact:
            # D/K are the approved individual Tableau amounts, in $000.
            sheet.cell(row, 4).value = month[4]
            sheet.cell(row, 11).value = ytd[4]
            for (
                consolidated_col,
                individual_col,
                variance_col,
                status_col,
                note_col,
                result,
            ) in ((3, 4, 6, 7, 8, month), (10, 11, 13, 14, 15, ytd)):
                consolidated = sheet.cell(row, consolidated_col).value
                individual = sheet.cell(row, individual_col).value
                sheet.cell(row, variance_col).value = variance(consolidated, individual)
                fs = sheet.cell(row, 2 if result is month else 9).value
                sheet.cell(row, status_col).value = validation_status(
                    fs, consolidated, individual, None, result[7]
                )
                style_status(sheet.cell(row, status_col))
                note = str(result[13] or "").strip()
                if note:
                    existing = str(sheet.cell(row, note_col).value or "").strip()
                    sheet.cell(row, note_col).value = (existing + " | " + note).strip(
                        " |"
                    )
        else:
            values = {
                8: ytd[1],
                9: ytd[2],
                12: month[4],
                13: month[5],
                14: month[6],
                16: ytd[4],
                17: ytd[5],
                18: ytd[6],
            }
            for column, value in values.items():
                sheet.cell(row, column).value = value
            status_cell(sheet.cell(row, 15), month[7])
            status_cell(sheet.cell(row, 19), ytd[7])
            for col, result in ((15, month), (19, ytd)):
                if result[13]:
                    sheet.cell(row, col).comment = Comment(
                        str(result[13]), "Validation"
                    )
    if any(
        str(s["A6"].value or "").strip() == "Financial Statement Line"
        for s in report
        if s.title.startswith(("ALIV ", "CBL "))
    ):
        for name in ("Individual Actuals", "Validation Summary", "Group Actuals"):
            if name in report.sheetnames:
                del report[name]
        refresh_summary(report)


def refresh_summary(report):
    sheet = report["Executive Summary"]
    anchor = next(
        (
            r
            for r in range(1, sheet.max_row + 1)
            if sheet.cell(r, 1).value == "ADDITIONAL VALIDATION CHECKS"
        ),
        sheet.max_row + 3,
    )
    if anchor <= sheet.max_row:
        sheet.delete_rows(anchor, sheet.max_row - anchor + 1)
    sheet.cell(anchor, 1, "ADDITIONAL VALIDATION CHECKS")
    for col, label in enumerate(
        ["Company", "Check", "Passed", "Failed", "Action Required", "Not Validated"], 1
    ):
        sheet.cell(anchor + 1, col, label)
    position = anchor + 2
    for company in ("ALIV", "CBL"):
        company_sheets = [
            s
            for s in report
            if s.title.startswith(company + " ") and s.title != "Executive Summary"
        ]
        for label, column in [
            ("Group YTD", 11),
            ("Individual Monthly", 15),
            ("Individual YTD", 19),
        ]:
            statuses = [
                s.cell(r, column).value
                for s in company_sheets
                for r in range(7, s.max_row + 1)
            ]
            for col, val in enumerate(
                [
                    company,
                    label,
                    statuses.count("PASS"),
                    statuses.count("FAIL"),
                    sum(v in ("ACTION REQUIRED", "ERROR") for v in statuses),
                    sum(
                        v not in ("PASS", "FAIL", "ACTION REQUIRED", "ERROR")
                        for v in statuses
                    ),
                ],
                1,
            ):
                sheet.cell(position, col, val)
            position += 1
    for row in (anchor, anchor + 1):
        for cell in sheet[row][:6]:
            cell.fill = PatternFill("solid", fgColor="145A70")
            cell.font = Font(bold=True, color="FFFFFF")
    exceptions = []
    pending = False
    for detail in report:
        if not detail.title.startswith(("ALIV ", "CBL ")):
            continue
        for row in range(7, detail.max_row + 1):
            for label, column in [
                ("Group YTD", 11),
                ("Individual Monthly", 15),
                ("Individual YTD", 19),
            ]:
                cell = detail.cell(row, column)
                if cell.value in ("FAIL", "ACTION REQUIRED", "ERROR"):
                    exceptions.append(
                        [
                            detail.title,
                            detail.cell(row, 1).value,
                            label,
                            cell.value,
                            (
                                cell.comment.text
                                if cell.comment
                                else "Review the variances on the company sheet."
                            ),
                        ]
                    )
                elif cell.value != "PASS":
                    pending = True
    sheet.cell(position + 1, 1, "ADDITIONAL EXCEPTIONS REQUIRING REVIEW")
    for col, label in enumerate(
        ["Company Sheet", "Statement Line", "Check", "Status", "Action / Note"], 1
    ):
        sheet.cell(position + 2, col, label)
    for offset, values in enumerate(
        exceptions
        or [
            [
                "",
                "",
                "",
                "",
                "No additional exceptions identified; unvalidated checks remain pending.",
            ]
        ]
    ):
        for col, value in enumerate(values, 1):
            cell = sheet.cell(position + 3 + offset, col, value)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for cells in sheet:
        for cell in cells:
            if isinstance(cell.value, str) and cell.value.startswith(
                "OVERALL VALIDATION STATUS:"
            ):
                if exceptions:
                    cell.value = "OVERALL VALIDATION STATUS: REVIEW REQUIRED"
                elif pending and cell.value.endswith("PASS"):
                    cell.value = "OVERALL VALIDATION STATUS: PENDING"
                if not cell.value.endswith(": PASS"):
                    cell.fill = PatternFill("solid", fgColor="FFF2CC")


def report_status(path):
    report = load_workbook(path, read_only=True, data_only=True)
    try:
        for row in report["Executive Summary"]:
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith(
                    "OVERALL VALIDATION STATUS:"
                ):
                    return cell.value.split(":", 1)[1].strip()
        return "PENDING"
    finally:
        report.close()


def prepare_group_copy(source, destination, tab, closed_period):
    root = ET.parse(source).getroot()
    sheet = root.find(f"./worksheets/worksheet[@name='{tab}']")
    if sheet is None:
        raise ValueError(f"Group worksheet not found: {tab}")
    start = date(
        closed_period.year + (closed_period.month == 12),
        closed_period.month % 12 + 1,
        1,
    )
    fiscal_start = date(closed_period.year - (closed_period.month < 7), 7, 1)
    fiscal_end = fiscal_start.year + 1
    names = {
        c.get("name")
        for c in sheet.iter("column")
        if c.get("caption") == "Forecast Start"
    }
    if len(names) != 1:
        raise ValueError("Expected one group Forecast Start dependency.")
    for c in root.iter("column"):
        if c.get("name") in names and "param-domain-type" in c.attrib:
            c.set("value", f"#{start:%Y-%m-%d}#")
            if c.find("calculation") is not None:
                c.find("calculation").set("formula", f"#{start:%Y-%m-%d}#")
    months = []
    current = fiscal_start
    while current < start:
        months.append(current)
        current = date(current.year + (current.month == 12), current.month % 12 + 1, 1)
    for f in root.iter("filter"):
        column = f.get("column", "")
        if column.endswith("[my:GL Posting Date:ok]"):
            for child in list(f):
                f.remove(child)
            union = ET.SubElement(f, "groupfilter", function="union")
            for month in months:
                ET.SubElement(
                    union,
                    "groupfilter",
                    function="member",
                    level="[my:GL Posting Date:ok]",
                    member=f"{month:%Y%m}",
                )
        if column.endswith("[none:Fiscal Year:ok]"):
            for child in list(f):
                f.remove(child)
            ET.SubElement(
                f,
                "groupfilter",
                function="member",
                level="[none:Fiscal Year:ok]",
                member=str(fiscal_end),
            )
        if column.endswith("[none:Forecast Type:nk]"):
            for member in f.findall('.//groupfilter[@function="member"]'):
                if re.fullmatch(
                    r"FY\d{2} Forecast_\d{6}", member.get("member", "").strip('"'), re.I
                ):
                    forecast_end = start.year + (start.month >= 7)
                    member.set(
                        "member", f'"FY{forecast_end%100:02} Forecast_{start:%m%Y}"'
                    )
    # Preserve all department, account, and Excl Media filters from the supplied TWB.
    ET.ElementTree(root).write(destination, encoding="utf-8", xml_declaration=True)


# Approved compact A:O report implementation, housed in actuals_report_layout.py
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from os import replace as _atomic_replace

TOLERANCE_000 = 1.0  # $1,000 tolerance, since values are in $000

HEADERS = [
    "Financial Statement Line",
    "Financial Statement ($000)",
    "Consolidated Tableau ($000)",
    "Individual Tableau ($000)",
    "Variance 1: FS vs Consolidated ($000)",
    "Variance 2: Consolidated vs Individual ($000)",
    "Monthly Status",
    "Financial Statement ($000)",
    "YTD Consolidated Tableau ($000)",
    "YTD Individual Tableau ($000)",
    "Variance 1: FS vs Consolidated ($000)",
    "Variance 2: Consolidated vs Individual ($000)",
    "YTD Status",
]

# The approved report has 15 populated columns, A:O.
# Monthly: A:H; YTD: I:O.


def is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def variance(left, right):
    if not is_number(left) or not is_number(right):
        return None
    return round(float(right) - float(left), 6)


def validation_status(
    fs, consolidated, individual, old_group=None, old_individual=None
):
    if not all(is_number(v) for v in (fs, consolidated, individual)):
        return "PENDING"

    v1 = variance(fs, consolidated)
    v2 = variance(consolidated, individual)

    if abs(v1) <= TOLERANCE_000 and abs(v2) <= TOLERANCE_000:
        return "PASS"

    return "REVIEW REQUIRED"


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

    for address in ("B6", "I6"):
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
            fgColor="703070" if address == "B6" else "145A70",
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
            fgColor="703070" if col <= 8 else "145A70",
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
        "H": 43,
        "I": 19,
        "J": 21,
        "K": 21,
        "L": 21,
        "M": 21,
        "N": 19,
        "O": 43,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    ws.freeze_panes = "B8"
    ws.sheet_view.showGridLines = False


def clean_variance_formatting(ws):
    """Remove inherited variance fills and display rounded $000 values."""

    from openpyxl.styles import PatternFill, Alignment

    # Monthly and YTD financial values, including both variances
    financial_columns = (2, 3, 4, 5, 6, 8, 9, 10, 11, 12)

    # Only the four variance columns should have their fills cleared
    variance_columns = (5, 6, 11, 12)

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


def _build_compact_report(
    source,
    template=None,
    destination=None,
    forecast_type=None,
    forecast_start=None,
    run_at=None,
):
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
    # Build the approved A:O layout internally when no template is supplied.
    layout = load_workbook(template) if template is not None else Workbook()

    try:
        # Only use the original template sheet.
        base = layout.active

        for sheet in list(layout):
            if sheet != base:
                layout.remove(sheet)

        from datetime import datetime

        run_at = run_at or datetime.now()
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
            ws["A2"] = "Closed period"
            ws["C2"] = original["C2"].value
            ws["A3"] = "Forecast start / type"
            ws["C3"] = (
                f"{forecast_start or 'Not supplied'} | {forecast_type or 'Not supplied'}"
            )
            ws["A4"] = "Run date"
            ws["I4"] = run_at.strftime("%Y-%m-%d %H:%M:%S")

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
                    ytd_fs,  # I
                    ytd_group,  # J
                    ytd_individual,  # K
                    ytd_v1,  # L
                    ytd_v2,  # M
                    ytd_status,  # N
                ]

                for col, value in enumerate(values, 1):
                    cell = ws.cell(output_row, col)
                    cell.value = value

                    cell.alignment = Alignment(
                        vertical="center",
                        wrap_text=col in (1, 7, 13),
                    )

                    if col in (2, 3, 4, 5, 6, 8, 9, 10, 11, 12):
                        cell.number_format = "#,##0;[Red](#,##0);0"

                style_status(ws.cell(output_row, 7))
                style_status(ws.cell(output_row, 13))

                statuses.extend([monthly_status, ytd_status])

                output_row += 1

            ws["G2"] = overall_status(statuses)
            style_status(ws["G2"])

            ws.auto_filter.ref = f"A7:M{max(8, output_row - 1)}"

            clean_variance_formatting(ws)

            results[company] = {
                "rows": output_row - 8,
                "status": ws["G2"].value,
            }

        summary = layout.create_sheet("Executive Summary", 0)
        summary.merge_cells("A1:H2")
        summary["A1"] = "ACTUALS VALIDATION | EXECUTIVE SUMMARY"
        summary["A1"].font = Font(name="Aptos", size=17, bold=True, color="FFFFFF")
        summary["A1"].fill = PatternFill("solid", fgColor="243F5C")
        summary["A4"] = "Closed period"
        summary["B4"] = str(base["C2"].value or "See company worksheets")
        summary["A5"] = "Forecast start date"
        summary["B5"] = str(forecast_start or "Not supplied")
        summary["A6"] = "Forecast type"
        summary["B6"] = str(forecast_type or "Not supplied")
        summary["A7"] = "Date run"
        summary["B7"] = run_at.strftime("%Y-%m-%d %H:%M:%S")
        for col, heading in enumerate(
            (
                "Company",
                "Rows",
                "Status",
                "Monthly PASS",
                "Monthly Review",
                "Monthly Pending",
                "YTD PASS",
                "YTD Review",
                "YTD Pending",
            ),
            1,
        ):
            cell = summary.cell(10, col, heading)
            cell.fill = PatternFill("solid", fgColor="145A70")
            cell.font = Font(bold=True, color="FFFFFF")
        for row_idx, company in enumerate(("ALIV", "CBL"), 11):
            ws = layout[next(name for name in layout.sheetnames if name.startswith(company + " ") and name != "Executive Summary")]
            monthly = [str(ws.cell(r, 7).value or "") for r in range(8, ws.max_row + 1)]
            ytd = [str(ws.cell(r, 13).value or "") for r in range(8, ws.max_row + 1)]
            data = (
                company,
                results[company]["rows"],
                results[company]["status"],
                monthly.count("PASS"),
                monthly.count("REVIEW REQUIRED"),
                monthly.count("PENDING"),
                ytd.count("PASS"),
                ytd.count("REVIEW REQUIRED"),
                ytd.count("PENDING"),
            )
            for c, value in enumerate(data, 1):
                summary.cell(row_idx, c, value)
        summary["A14"] = "Overall status"
        summary["B14"] = overall_status([results[k]["status"] for k in ("ALIV", "CBL")])
        for col in "ABCDEFGHI":
            summary.column_dimensions[col].width = 22
        summary.column_dimensions["A"].width = 30
        summary.freeze_panes = "B11"
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


def finalize_actuals_report_in_place(
    report_path, forecast_type=None, forecast_start=None
):
    """Convert legacy A:S reports to approved A:O; safe for repeated runs."""

    path = Path(report_path)
    staged = path.with_name(path.stem + ".__compact_tmp__.xlsx")

    try:
        existing = load_workbook(path, read_only=True)

        try:
            # Accept both OPEX and Direct Cost, and either compact header variant.
            company_sheets = [
                sheet for sheet in existing
                if sheet.title != "Executive Summary"
                and sheet.title.startswith(("ALIV ", "CBL "))
            ]
            already_compact = (
                len(company_sheets) == 2
                and {sheet.title.split(" ", 1)[0] for sheet in company_sheets}
                    == {"ALIV", "CBL"}
                and all(
                    sheet.cell(7, 1).value == "Financial Statement Line"
                    and sheet.cell(7, 3).value == "Consolidated Tableau ($000)"
                    and sheet.cell(7, 4).value == "Individual Tableau ($000)"
                    and sheet.cell(7, 7).value == "Monthly Status"
                    and sheet.cell(7, 8).value == "Financial Statement ($000)"
                    and sheet.cell(7, 13).value == "YTD Status"
                    for sheet in company_sheets
                )
            )

            if already_compact:
                return str(path)

        finally:
            existing.close()

        _build_compact_report(
            path,
            destination=staged,
            forecast_type=forecast_type,
            forecast_start=forecast_start,
        )

        _atomic_replace(str(staged), str(path))

    finally:
        staged.unlink(missing_ok=True)

    return str(path)
