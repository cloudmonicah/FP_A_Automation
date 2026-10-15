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
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

from shared import excel_style as style

STATUSES = ("PASS", "FAIL", "ACTION REQUIRED", "NOT VALIDATED")

# Order rows are listed in within a category: problems first.
_SEVERITY = {"FAIL": 0, "ACTION REQUIRED": 1, "NOT VALIDATED": 2, "PASS": 3}

DETAIL_COLUMNS = (
    # (header, result key, width, kind)
    ("Category", "category", 15, "text"),
    ("Status", "status", 20, "status"),
    ("Forecast Workbook", "forecast_workbook", 44, "text"),
    ("Worksheet", "worksheet", 24, "text"),
    ("Excel Grand Total", "excel_grand_total", 17, "money"),
    ("Tableau Grand Total", "tableau_grand_total", 17, "money"),
    ("Variance", "variance", 14, "variance"),
    ("Note", "note", 60, "note"),
    ("Individual Tableau Workbook", "individual_tableau_workbook", 34, "text"),
    ("Individual Tableau Tab", "individual_tableau_worksheet", 30, "text"),
    ("Period Mode", "period_mode", 15, "text"),
    ("Forecast Type", "forecast_type", 22, "text"),
)


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


def _display_status(item) -> str:
    if item.get("not_configured"):
        return "NOT SET UP"
    return str(item.get("status") or "").upper()


def _sorted(results):
    return sorted(
        results,
        key=lambda item: (
            item.get("_category_order", 0),
            item.get("not_configured", False),
            _SEVERITY.get(str(item.get("status") or "").upper(), 9),
        ),
    )


def _build_workbook(*, results, store, ordered, missing, all_categories, fiscal_year,
                    period, forecast_type, tolerance) -> Workbook:
    order = {category: index for index, category in enumerate(ordered)}
    results = _sorted(
        [{**item, "_category_order": order.get(item.get("category", ""), 99)} for item in results]
    )
    c = counts(results)
    overall = overall_status(results, missing)

    wb = Workbook()
    summary = wb.active
    summary.title = "Validation Summary"
    _summary_sheet(summary, results, store, ordered, missing, c, overall,
                   fiscal_year, period, forecast_type, tolerance)

    review = wb.create_sheet("Needs Review")
    problems = [
        item for item in results
        if _display_status(item) in ("FAIL", "ACTION REQUIRED", "NOT VALIDATED")
    ]
    _detail_sheet(review, problems, tolerance,
                  empty_text="Nothing needs review: every route that ran passed "
                             "or is not set up for validation yet.")

    detail = wb.create_sheet("Validation Detail")
    _detail_sheet(detail, results, tolerance, empty_text="No routes have been validated yet.")
    return wb


def _summary_sheet(ws, results, store, ordered, missing, c, overall,
                   fiscal_year, period, forecast_type, tolerance) -> None:
    last_col = 10
    for col in range(1, last_col + 1):
        ws.column_dimensions[get_column_letter(col)].width = 13
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 17

    # Title banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    title = ws.cell(1, 1, "Forecast Validation Report")
    title.font = style.font(size=style.TITLE_SIZE, bold=True, color=style.WHITE)
    title.fill = style.fill(style.NAVY)
    title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    for col in range(2, last_col + 1):
        ws.cell(1, col).fill = style.fill(style.NAVY)

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    sub = ws.cell(
        2, 1,
        f"{period:%B %Y}  ·  {fiscal_year}  ·  {forecast_type}  ·  "
        f"Generated {datetime.now():%b %d, %Y %I:%M %p}",
    )
    sub.font = style.font(color=style.SLATE, italic=True)
    sub.alignment = Alignment(indent=1)

    # Overall status banner
    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=last_col)
    note = ""
    if missing:
        note = f"  —  not run yet: {', '.join(missing)}"
    elif overall == "RECONCILED" and c["NOT SET UP"]:
        note = f"  —  {c['NOT SET UP']} routes not set up for validation yet"
    banner = ws.cell(4, 1, f"Overall status: {overall}{note}")
    style.status_style(banner, overall)
    banner.font = style.font(size=12, bold=True, color=style.STATUS_COLORS.get(overall, ("", "000000"))[1])
    banner.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    for col in range(2, last_col + 1):
        ws.cell(4, col).fill = style.fill(style.STATUS_COLORS.get(overall, ("FFFFFF",))[0])
    ws.row_dimensions[4].height = 26

    # KPI tiles: big number above its label
    tiles = [
        ("Passed", c["PASS"], "PASS"),
        ("Failed", c["FAIL"], "FAIL"),
        ("Action Required", c["ACTION REQUIRED"], "ACTION REQUIRED"),
        ("Not Validated", c["NOT VALIDATED"] - c["NOT SET UP"], "NOT VALIDATED"),
        ("Not Set Up", c["NOT SET UP"], "NOT SET UP"),
    ]
    for index, (label, value, status) in enumerate(tiles):
        col = 1 + index * 2
        background, text = style.STATUS_COLORS[status]
        for row in (6, 7):
            ws.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + 1)
            for offset in (0, 1):
                ws.cell(row, col + offset).fill = style.fill(background if value else "F2F2F2")
        number = ws.cell(6, col, value)
        number.font = style.font(size=22, bold=True, color=text if value else style.MUTED)
        number.alignment = Alignment(horizontal="center", vertical="center")
        caption = ws.cell(7, col, label)
        caption.font = style.font(bold=True, color=text if value else style.MUTED)
        caption.alignment = Alignment(horizontal="center", vertical="top")
    ws.row_dimensions[6].height = 36
    ws.row_dimensions[7].height = 18

    # Run details
    row = 9
    ws.cell(row, 1, "Run details").font = style.font(size=12, bold=True, color=style.NAVY)
    details = [
        ("Forecast Period", period.strftime("%B %Y")),
        ("Fiscal Year", fiscal_year),
        ("Forecast Start", period),
        ("Forecast Type", forecast_type),
        ("Tolerance", tolerance),
        ("Routes Reported", len(results)),
        ("Categories Not Run Yet", ", ".join(missing) or "None"),
        ("Overall Status", overall),
    ]
    for label, value in details:
        row += 1
        key = ws.cell(row, 1, label)
        key.font = style.font(bold=True, color=style.SLATE)
        key.border = style.UNDERLINE
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=5)
        cell = ws.cell(row, 2, value)
        cell.font = style.font()
        cell.alignment = Alignment(horizontal="left")
        for col in range(2, 6):
            ws.cell(row, col).border = style.UNDERLINE
        if label == "Forecast Start":
            cell.number_format = "mmmm d, yyyy"
        elif label == "Tolerance":
            cell.number_format = '"± "$#,##0.00'
        elif label == "Overall Status":
            style.status_style(cell, value)
            cell.alignment = Alignment(horizontal="left", indent=1)

    # By category
    row += 2
    ws.cell(row, 1, "By category").font = style.font(size=12, bold=True, color=style.NAVY)
    row += 1
    headers = ["Category", "Last Validated", "Routes", "Passed", "Failed",
               "Action Req.", "Not Validated", "Not Set Up", "Status", ""]
    for col, header in enumerate(headers, 1):
        ws.cell(row, col, header)
    ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
    style.header_row(ws, row, 1, last_col)
    ws.row_dimensions[row].height = 22
    for index, category in enumerate([*ordered, *missing]):
        row += 1
        entry = store.get(category)
        if entry is None:
            values = [category, "Not run yet", None, None, None, None, None, None]
            status = "NOT RUN"
        else:
            cc = counts(entry["results"])
            values = [category, entry.get("validated_at"), len(entry["results"]), cc["PASS"],
                      cc["FAIL"], cc["ACTION REQUIRED"], cc["NOT VALIDATED"] - cc["NOT SET UP"],
                      cc["NOT SET UP"]]
            status = overall_status(entry["results"], [])
        band = style.fill(style.BAND) if index % 2 else None
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.font = style.font(bold=(col == 1), color=style.MUTED if entry is None and col > 1 else None)
            cell.border = style.GRID
            cell.alignment = Alignment(horizontal="left" if col <= 2 else "center", vertical="center")
            if col >= 3:
                cell.number_format = style.COUNT
            if band:
                cell.fill = band
        ws.merge_cells(start_row=row, start_column=9, end_row=row, end_column=10)
        status_cell = ws.cell(row, 9, status)
        style.status_style(status_cell, status)
        for col in (9, 10):
            ws.cell(row, col).border = style.GRID
        ws.row_dimensions[row].height = 20

    # Status guide
    row += 2
    ws.cell(row, 1, "What the statuses mean").font = style.font(size=12, bold=True, color=style.NAVY)
    guides = [
        ("PASS", "Excel and the individual Tableau Grand Total agree within the tolerance."),
        ("FAIL", "Compared, but the difference is larger than the tolerance."),
        ("ACTION REQUIRED", "The automation could not finish this route (Tableau or a missing "
                            "file). See Note; the other routes still ran."),
        ("NOT VALIDATED", "Not compared because the forecast line looks wrong, for example it "
                          "is for a different month. See Note."),
        ("NOT SET UP", "This route still needs a department filter before it can be "
                       "validated automatically. It does not stop the month reconciling."),
        ("INCOMPLETE", "Everything that ran is fine, but some categories have not been "
                       "validated this month."),
    ]
    for status, meaning in guides:
        row += 1
        key = ws.cell(row, 1, status)
        style.status_style(key, status)
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=last_col)
        text = ws.cell(row, 2, meaning)
        text.font = style.font(color=style.SLATE)
        text.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        ws.row_dimensions[row].height = 30

    style.finish_sheet(ws, landscape=False)


def _detail_sheet(ws, results, tolerance, empty_text: str) -> None:
    for col, (header, _key, width, _kind) in enumerate(DETAIL_COLUMNS, 1):
        ws.cell(1, col, header)
        ws.column_dimensions[get_column_letter(col)].width = width
    style.header_row(ws, 1, 1, len(DETAIL_COLUMNS))
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "C2"

    if not results:
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(DETAIL_COLUMNS))
        cell = ws.cell(2, 1, empty_text)
        cell.font = style.font(italic=True, color=style.MUTED)
        cell.alignment = Alignment(indent=1, vertical="center")
        ws.row_dimensions[2].height = 24
        style.finish_sheet(ws)
        return

    for r, item in enumerate(results, 2):
        band = style.fill(style.BAND) if r % 2 else None
        for col, (_header, key, _width, kind) in enumerate(DETAIL_COLUMNS, 1):
            value = _display_status(item) if kind == "status" else item.get(key)
            cell = ws.cell(r, col, value)
            cell.font = style.font()
            cell.border = style.GRID
            cell.alignment = Alignment(vertical="top", wrap_text=kind in ("note", "text"))
            if band:
                cell.fill = band
            if kind in ("money", "variance"):
                cell.number_format = style.MONEY
                cell.alignment = Alignment(vertical="top", horizontal="right")
            if kind == "variance" and isinstance(value, (int, float)) and abs(value) > tolerance:
                cell.font = style.font(bold=True, color="9C0006")
                cell.fill = style.fill("FBE3E3")
            if kind == "status":
                style.status_style(cell, value)
            if kind == "note":
                cell.font = style.font(color=style.SLATE)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(DETAIL_COLUMNS))}{len(results) + 1}"
    style.finish_sheet(ws)


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
