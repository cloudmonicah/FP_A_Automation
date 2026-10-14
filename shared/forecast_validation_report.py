from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _counts(results):
    out = {"PASS": 0, "FAIL": 0, "ACTION REQUIRED": 0, "NOT VALIDATED": 0}
    for x in results:
        s = str(x.get("status") or "").upper()
        out[s] = out.get(s, 0) + 1
    return out


def write_forecast_validation_report(
    validation_result,
    output_folder,
    fiscal_year,
    forecast_month,
    forecast_year,
    forecast_type,
    tolerance=5.00,
):
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    period = datetime(forecast_year, forecast_month, 1)
    path = output_folder / f"Forecast_Validation_{period:%b}_{forecast_year}.xlsx"
    results = list(validation_result.get("results") or [])
    counts = _counts(results)
    overall = (
        "RECONCILED"
        if not (
            counts.get("FAIL")
            or counts.get("ACTION REQUIRED")
            or counts.get("NOT VALIDATED")
        )
        else "REVIEW REQUIRED"
    )

    wb = Workbook()
    ws = wb.active
    ws.title = "Validation Summary"
    detail = wb.create_sheet("Validation Detail")

    navy = PatternFill("solid", fgColor="1F4E78")
    blue = PatternFill("solid", fgColor="D9EAF7")
    green = PatternFill("solid", fgColor="E2F0D9")
    red = PatternFill("solid", fgColor="FCE4D6")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    grey = PatternFill("solid", fgColor="EDEDED")
    edge = Border(bottom=Side(style="thin", color="B7B7B7"))

    ws.merge_cells("A1:F1")
    c = ws["A1"]
    c.value = "FORECAST VALIDATION REPORT"
    c.font = Font(size=18, bold=True, color="FFFFFF")
    c.fill = navy
    ws.row_dimensions[1].height = 28

    rows = [
        ("Forecast Period", period.strftime("%B %Y")),
        ("Fiscal Year", fiscal_year),
        ("Forecast Start", period),
        ("Forecast Type", forecast_type),
        ("Tolerance", tolerance),
        ("Routes Reported", len(results)),
        ("Passed", counts.get("PASS", 0)),
        ("Failed", counts.get("FAIL", 0)),
        ("Action Required", counts.get("ACTION REQUIRED", 0)),
        ("Not Validated", counts.get("NOT VALIDATED", 0)),
        ("Overall Status", overall),
    ]
    value_cells = {}
    for r, (label, value) in enumerate(rows, start=3):
        ws.cell(r, 1, label).font = Font(bold=True)
        value_cells[label] = ws.cell(r, 2, value)
        ws.cell(r, 1).border = edge
        ws.cell(r, 2).border = edge
    value_cells["Forecast Start"].number_format = "mm/dd/yyyy"
    value_cells["Tolerance"].number_format = "$#,##0.00;[Red]-$#,##0.00"
    value_cells["Overall Status"].fill = green if overall == "RECONCILED" else yellow
    value_cells["Overall Status"].font = Font(bold=True)
    guide_row = 3 + len(rows) + 1
    ws.cell(guide_row, 1, "Status Guide").font = Font(bold=True)
    guides = [
        (
            "PASS",
            "Excel and individual Tableau Grand Totals reconcile within tolerance.",
        ),
        ("FAIL", "Validation completed but the variance exceeds tolerance."),
        (
            "ACTION REQUIRED",
            "The automation could not finish this route (Tableau or a missing "
            "file). See Note; the other routes still ran.",
        ),
        (
            "NOT VALIDATED",
            "Not compared: the forecast line is suspect (for example a different "
            "forecast month) or the route is not set up for automated validation. "
            "See Note.",
        ),
    ]
    for r, (s, m) in enumerate(guides, start=guide_row + 1):
        ws.cell(r, 1, s).font = Font(bold=True)
        ws.cell(r, 2, m)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 75

    headers = [
        "Forecast Workbook",
        "Worksheet",
        "Individual Tableau Workbook",
        "Individual Tableau Worksheet",
        "Excel Grand Total",
        "Tableau Grand Total",
        "Variance",
        "Status",
        "Period Mode",
        "Forecast Start",
        "Forecast Type",
        "Note",
    ]
    for col, h in enumerate(headers, 1):
        c = detail.cell(1, col, h)
        c.font = Font(bold=True)
        c.fill = blue
        c.border = edge
        c.alignment = Alignment(horizontal="center", vertical="center")

    for r, item in enumerate(results, 2):
        vals = [
            item.get("forecast_workbook"),
            item.get("worksheet"),
            item.get("individual_tableau_workbook"),
            item.get("individual_tableau_worksheet"),
            item.get("excel_grand_total"),
            item.get("tableau_grand_total"),
            item.get("variance"),
            item.get("status"),
            item.get("period_mode"),
            item.get("forecast_start"),
            item.get("forecast_type"),
            item.get("note"),
        ]
        for col, v in enumerate(vals, 1):
            detail.cell(r, col, v)
            detail.cell(r, col).border = edge
        for col in (5, 6, 7):
            detail.cell(r, col).number_format = "$#,##0.00;[Red]-$#,##0.00"
        s = str(item.get("status") or "").upper()
        sc = detail.cell(r, 8)
        sc.font = Font(bold=True)
        sc.fill = (
            green
            if s == "PASS"
            else red if s == "FAIL" else grey if s == "NOT VALIDATED" else yellow
        )

    detail.freeze_panes = "A2"
    detail.auto_filter.ref = detail.dimensions
    widths = {
        1: 46,
        2: 28,
        3: 42,
        4: 38,
        5: 18,
        6: 20,
        7: 16,
        8: 18,
        9: 18,
        10: 18,
        11: 26,
        12: 70,
    }
    for col, w in widths.items():
        detail.column_dimensions[get_column_letter(col)].width = w
    for row in detail.iter_rows():
        for c in row:
            c.alignment = Alignment(vertical="top", wrap_text=True)

    try:
        return _save(wb, path)
    finally:
        wb.close()


def _save(wb, path: Path) -> Path:
    """Save without losing the run when the old report is open in Excel.

    The workbook is written to a staging file first. If the existing report
    is locked (open in Excel), the new one is kept under a timestamped name
    next to it instead of failing after a long Tableau run.
    """
    staged = path.with_name(f"~staging_{path.name}")
    wb.save(staged)
    try:
        os.replace(staged, path)
        return path
    except PermissionError:
        fallback = path.with_name(f"{path.stem} ({datetime.now():%Y-%m-%d %H%M%S}){path.suffix}")
        os.replace(staged, fallback)
        return fallback
    finally:
        staged.unlink(missing_ok=True)
