"""Monthly forecast validation report.

One report per forecast month. Categories (OPEX, Revenue, ...) are validated
separately, so the latest results of every category are kept in a small JSON
file next to the report and the whole report is rebuilt from it on every run:
running Revenue replaces Revenue's rows and leaves the other categories as
they were.
"""

from __future__ import annotations

import json
import os
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

STATUSES = ("PASS", "FAIL", "ACTION REQUIRED", "NOT VALIDATED")

NAVY = PatternFill("solid", fgColor="1F4E78")
BLUE = PatternFill("solid", fgColor="D9EAF7")
GREEN = PatternFill("solid", fgColor="E2F0D9")
RED = PatternFill("solid", fgColor="FCE4D6")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="EDEDED")
EDGE = Border(bottom=Side(style="thin", color="B7B7B7"))
MONEY = "$#,##0.00;[Red]-$#,##0.00"


def report_path_for(output_folder, forecast_month, forecast_year) -> Path:
    period = date(forecast_year, forecast_month, 1)
    return Path(output_folder) / f"Forecast_Validation_{period:%b}_{forecast_year}.xlsx"


def _store_path(report_path: Path) -> Path:
    return report_path.with_suffix(".results.json")


def _load_store(path: Path, forecast_type: str) -> dict:
    """Saved results by category; empty if missing, unreadable or another month."""
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    if not isinstance(data, dict) or data.get("forecast_type") != forecast_type:
        return {}
    categories = data.get("categories")
    return categories if isinstance(categories, dict) else {}


def _plain(value):
    """JSON-safe copy of a result value."""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    return value


def load_saved_results(output_folder, forecast_month, forecast_year, forecast_type) -> dict:
    """The month's saved results by category ({} if none yet)."""
    report = report_path_for(output_folder, forecast_month, forecast_year)
    return _load_store(_store_path(report), forecast_type)


def counts(results) -> dict:
    out = {status: 0 for status in STATUSES}
    out["NOT SET UP"] = 0
    for item in results:
        status = str(item.get("status") or "").upper()
        out[status] = out.get(status, 0) + 1
        if item.get("not_configured"):
            out["NOT SET UP"] += 1
    return out


def overall_status(results, categories_missing) -> str:
    c = counts(results)
    suspect_not_validated = c["NOT VALIDATED"] - c["NOT SET UP"]
    if c["FAIL"] or c["ACTION REQUIRED"] or suspect_not_validated:
        return "REVIEW REQUIRED"
    if categories_missing:
        return "INCOMPLETE"
    if not results:
        return "NO RESULTS"
    return "RECONCILED"


def write_forecast_validation_report(
    validation_result,
    output_folder,
    fiscal_year,
    forecast_month,
    forecast_year,
    forecast_type,
    tolerance=5.00,
    categories_run=None,
    all_categories=None,
):
    """Merge this run into the month's results and rebuild the report.

    ``categories_run``: categories this run validated (their old rows are
    replaced). None means the results are the whole month (nothing merged).
    Returns the path written; if the report is open in Excel, a timestamped
    copy is written next to it instead.
    """
    if all_categories is None:
        from automations.forecast.mapping import FORECAST_CATEGORIES

        all_categories = FORECAST_CATEGORIES

    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)
    path = report_path_for(output_folder, forecast_month, forecast_year)
    store_path = _store_path(path)

    run_results = [
        {key: _plain(value) for key, value in item.items()}
        for item in (validation_result.get("results") or [])
    ]
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    store = _load_store(store_path, forecast_type) if categories_run else {}
    run_categories = list(categories_run or all_categories)
    for category in run_categories:
        store[category] = {
            "validated_at": now,
            "results": [r for r in run_results if r.get("category", "") == category],
        }
    # Results without a category (older callers) are kept under "Other".
    uncategorised = [r for r in run_results if r.get("category", "") not in run_categories]
    if uncategorised:
        store["Other"] = {"validated_at": now, "results": uncategorised}

    _write_store(store_path, forecast_type, store)

    ordered = [c for c in all_categories if c in store] + [c for c in store if c not in all_categories]
    results = [r for c in ordered for r in store[c]["results"]]
    missing = [c for c in all_categories if c not in store]

    wb = _build_workbook(
        results=results,
        store=store,
        ordered=ordered,
        missing=missing,
        all_categories=all_categories,
        fiscal_year=fiscal_year,
        period=datetime(forecast_year, forecast_month, 1),
        forecast_type=forecast_type,
        tolerance=tolerance,
    )
    try:
        return _save(wb, path)
    finally:
        wb.close()


def _write_store(path: Path, forecast_type: str, store: dict) -> None:
    payload = {"forecast_type": forecast_type, "categories": store}
    staged = path.with_name(f"~staging_{path.name}")
    staged.write_text(json.dumps(payload, indent=1, default=str), encoding="utf-8")
    os.replace(staged, path)


def _build_workbook(*, results, store, ordered, missing, all_categories, fiscal_year,
                    period, forecast_type, tolerance) -> Workbook:
    c = counts(results)
    overall = overall_status(results, missing)

    wb = Workbook()
    ws = wb.active
    ws.title = "Validation Summary"
    detail = wb.create_sheet("Validation Detail")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
    title = ws.cell(1, 1, "FORECAST VALIDATION REPORT")
    title.font = Font(size=18, bold=True, color="FFFFFF")
    title.fill = NAVY
    ws.row_dimensions[1].height = 28

    rows = [
        ("Forecast Period", period.strftime("%B %Y")),
        ("Fiscal Year", fiscal_year),
        ("Forecast Start", period),
        ("Forecast Type", forecast_type),
        ("Tolerance", tolerance),
        ("Routes Reported", len(results)),
        ("Passed", c["PASS"]),
        ("Failed", c["FAIL"]),
        ("Action Required", c["ACTION REQUIRED"]),
        ("Not Validated", c["NOT VALIDATED"] - c["NOT SET UP"]),
        ("Not Set Up for Validation", c["NOT SET UP"]),
        ("Categories Not Run Yet", ", ".join(missing) or "None"),
        ("Overall Status", overall),
    ]
    cells = {}
    for r, (label, value) in enumerate(rows, start=3):
        ws.cell(r, 1, label).font = Font(bold=True)
        cells[label] = ws.cell(r, 2, value)
        ws.cell(r, 1).border = EDGE
        ws.cell(r, 2).border = EDGE
    cells["Forecast Start"].number_format = "mm/dd/yyyy"
    cells["Tolerance"].number_format = MONEY
    cells["Overall Status"].fill = GREEN if overall == "RECONCILED" else YELLOW
    cells["Overall Status"].font = Font(bold=True)

    # Per-category table.
    row = 3 + len(rows) + 1
    ws.cell(row, 1, "By Category").font = Font(bold=True, size=12)
    row += 1
    headers = ["Category", "Last Validated", "Routes", "Passed", "Failed",
               "Action Required", "Not Validated", "Not Set Up", "Status"]
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row, col, header)
        cell.font = Font(bold=True)
        cell.fill = BLUE
        cell.border = EDGE
    for category in [*ordered, *missing]:
        row += 1
        entry = store.get(category)
        if entry is None:
            values = [category, "Not run yet", None, None, None, None, None, None, "NOT RUN"]
        else:
            cc = counts(entry["results"])
            values = [
                category,
                entry.get("validated_at"),
                len(entry["results"]),
                cc["PASS"],
                cc["FAIL"],
                cc["ACTION REQUIRED"],
                cc["NOT VALIDATED"] - cc["NOT SET UP"],
                cc["NOT SET UP"],
                overall_status(entry["results"], []),
            ]
        for col, value in enumerate(values, 1):
            ws.cell(row, col, value).border = EDGE
        status_cell = ws.cell(row, len(headers))
        status_cell.font = Font(bold=True)
        status_cell.fill = GREEN if values[-1] == "RECONCILED" else GREY if values[-1] == "NOT RUN" else YELLOW

    row += 2
    ws.cell(row, 1, "Status Guide").font = Font(bold=True)
    guides = [
        ("PASS", "Excel and individual Tableau Grand Totals reconcile within tolerance."),
        ("FAIL", "Validation completed but the variance exceeds tolerance."),
        ("ACTION REQUIRED", "The automation could not finish this route (Tableau or a "
                            "missing file). See Note; the other routes still ran."),
        ("NOT VALIDATED", "Not compared: the forecast line is suspect (for example a "
                          "different forecast month), or the route is not set up for "
                          "automated validation yet. See Note."),
        ("INCOMPLETE", "Every route that ran is fine, but some categories have not "
                       "been validated yet this month."),
    ]
    for status, meaning in guides:
        row += 1
        ws.cell(row, 1, status).font = Font(bold=True)
        ws.cell(row, 2, meaning)
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 34
    for col in range(3, len(headers) + 1):
        ws.column_dimensions[get_column_letter(col)].width = 15

    detail_headers = [
        "Category",
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
    for col, header in enumerate(detail_headers, 1):
        cell = detail.cell(1, col, header)
        cell.font = Font(bold=True)
        cell.fill = BLUE
        cell.border = EDGE
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    status_col = detail_headers.index("Status") + 1
    for r, item in enumerate(results, 2):
        values = [
            item.get("category"),
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
        for col, value in enumerate(values, 1):
            cell = detail.cell(r, col, value)
            cell.border = EDGE
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col in (6, 7, 8):
            detail.cell(r, col).number_format = MONEY
        status = str(item.get("status") or "").upper()
        cell = detail.cell(r, status_col)
        cell.font = Font(bold=True)
        cell.fill = (
            GREEN if status == "PASS"
            else RED if status == "FAIL"
            else GREY if status == "NOT VALIDATED"
            else YELLOW
        )

    detail.freeze_panes = "A2"
    detail.auto_filter.ref = f"A1:{get_column_letter(len(detail_headers))}{max(1, len(results) + 1)}"
    widths = [16, 46, 28, 42, 38, 18, 20, 16, 18, 18, 18, 26, 70]
    for col, width in enumerate(widths, 1):
        detail.column_dimensions[get_column_letter(col)].width = width
    return wb


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
        fallback = path.with_name(
            f"{path.stem} ({datetime.now():%Y-%m-%d %H%M%S}){path.suffix}"
        )
        os.replace(staged, fallback)
        return fallback
    finally:
        staged.unlink(missing_ok=True)
