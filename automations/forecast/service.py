from __future__ import annotations
from shared.friendly_errors import explain

from datetime import date
from pathlib import Path
import re
import xml.etree.ElementTree as ET

from openpyxl import Workbook

from .excel_service import (
    create_forecast_refresh,
    populate_forecast_sheet_from_audit,
)
from .mapping import FORECAST_ROUTES
from .tableau_service import set_forecast_start_parameter
from shared.forecast_output_paths import (
    get_forecast_validation_output_folder,
)

from shared.forecast_validation_report import (
    write_forecast_validation_report,
)
from .validation_service import run_grouped_individual_validation

MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


def fiscal_year_for(month: int, year: int) -> str:
    """July-June fiscal year: July 2026 .. June 2027 -> FY27."""
    return f"FY{(year + 1 if month >= 7 else year) % 100:02d}"


def build_forecast_period(
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """Build the single monthly period used by Excel and every Tableau source."""
    if forecast_month not in range(1, 13):
        raise ValueError("Forecast month must be between 1 and 12.")

    fiscal_year = str(fiscal_year or "").strip().upper()
    if not re.fullmatch(r"FY\d{2}", fiscal_year):
        raise ValueError("Fiscal year must look like FY27.")

    start = date(forecast_year, forecast_month, 1)
    expected = fiscal_year_for(forecast_month, forecast_year)
    if fiscal_year != expected:
        raise ValueError(
            f"{start:%B %Y} is in {expected} (July-June), not {fiscal_year}."
        )
    forecast_type = f"{fiscal_year} Forecast_{forecast_month:02d}{forecast_year}"

    return {
        "fiscal_year": fiscal_year,
        "forecast_month": forecast_month,
        "forecast_year": forecast_year,
        "forecast_start": start,
        "forecast_start_tableau": f"#{start:%Y-%m-%d}#",
        "forecast_type": forecast_type,
        "target_month_abbr": MONTH_ABBR[forecast_month],
        "display": f"{start:%B %Y}",
    }


def write_consolidated_forecast_copy(
    tableau_workbook_path: str | Path,
    destination: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """
    Write a copy of the consolidated FY Forecast workbook set to the GUI
    period (Forecast Start + saved Forecast Type selections).

    The source workbook is only read; ``destination`` should be a temp copy
    from shared.tableau_temp_copies.temp_twb_path.
    """
    destination = Path(destination)
    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    set_forecast_start_parameter(
        tableau_workbook_path=tableau_workbook_path,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
        output_path=destination,
    )

    tree = ET.parse(destination)
    root = tree.getroot()

    target = period["forecast_type"]
    changed = 0

    # Update saved Forecast Type filter members.
    for groupfilter in root.iter("groupfilter"):
        member = str(groupfilter.get("member") or "").strip('"')
        if re.fullmatch(r"FY\d{2} Forecast_\d{6}", member, re.I):
            groupfilter.set("member", f'"{target}"')
            changed += 1

    # Also update parameter-style declarations if a workbook uses them.
    for column in root.iter("column"):
        caption = re.sub(
            r"[^a-z0-9]+",
            "",
            str(column.get("caption") or "").casefold(),
        )
        if caption != "forecasttype":
            continue

        value = str(column.get("value") or "").strip('"')
        if re.fullmatch(r"FY\d{2} Forecast_\d{6}", value, re.I):
            column.set("value", f'"{target}"')
            calculation = column.find("calculation")
            if calculation is not None:
                formula = str(calculation.get("formula") or "")
                if re.search(r"FY\d{2} Forecast_\d{6}", formula, re.I):
                    calculation.set("formula", f'"{target}"')
            changed += 1

    tree.write(destination, encoding="utf-8", xml_declaration=True)

    # Verify Forecast Type was actually written where applicable.
    saved_members = [
        str(item.get("member") or "").strip('"')
        for item in ET.parse(destination).getroot().iter("groupfilter")
        if re.fullmatch(
            r"FY\d{2} Forecast_\d{6}",
            str(item.get("member") or "").strip('"'),
            re.I,
        )
    ]
    if changed and any(member != target for member in saved_members):
        raise RuntimeError(
            "Forecast Type was updated but could not be verified.\n"
            f"Expected: {target}"
        )

    period["forecast_type_updates"] = changed
    return period


def _temp_copy(source: str | Path, tag: str) -> Path:
    from shared.tableau_temp_copies import temp_twb_path

    return temp_twb_path(source, tag)


def check_consolidated_workbook(
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """Prove the consolidated workbook can be set to the period (temp copy, then deleted)."""
    temporary = _temp_copy(tableau_workbook_path, "Check")
    try:
        return write_consolidated_forecast_copy(
            tableau_workbook_path,
            temporary,
            fiscal_year,
            forecast_month,
            forecast_year,
        )
    finally:
        temporary.unlink(missing_ok=True)


def prepare_monthly_forecast(
    source_folder: str | Path,
    output_folder: str | Path,
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    overwrite_mode: str = "skip",
) -> dict:
    """
    Step 1: create the monthly Excel workbooks.

    The consolidated Tableau workbook is checked (a temp copy is set to the
    period and deleted) so a missing parameter is reported now, not after
    step 2 has opened Tableau. Tableau itself is opened by step 2.
    """
    period = check_consolidated_workbook(
        tableau_workbook_path=tableau_workbook_path,
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    excel_result = create_forecast_refresh(
        source_folder=source_folder,
        output_folder=output_folder,
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
        overwrite_mode=overwrite_mode,
    )

    return {
        "period": period,
        "excel": excel_result,
        "output_folder": str(Path(output_folder)),
    }


def _safe_sheet_title(name: str, used: set[str]) -> str:
    title = re.sub(r"[\\/*?:\[\]]", "_", str(name))[:31] or "Tableau"
    base = title
    suffix = 1
    while title.casefold() in used:
        suffix += 1
        tag = f"_{suffix}"
        title = base[: 31 - len(tag)] + tag
    used.add(title.casefold())
    return title


def write_tableau_audit_workbook(
    extracts: dict[str, list[dict]],
    audit_path: str | Path,
) -> Path:
    """
    Save the exact consolidated Tableau crosstabs into one audit workbook.
    Stage 4 already knows how to consume this workbook.
    """
    audit_path = Path(audit_path)
    audit_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()

    for tableau_name, rows in extracts.items():
        ws = wb.create_sheet(_safe_sheet_title(tableau_name, used))

        headers = []
        for row in rows:
            for key in row:
                if key not in headers:
                    headers.append(key)

        if headers:
            ws.append(headers)
            for row in rows:
                ws.append([row.get(header) for header in headers])

    if not wb.sheetnames:
        wb.create_sheet("No Data")

    wb.save(audit_path)
    wb.close()
    return audit_path


def _resolve_target_filename(
    baseline_name: str,
    target_month_abbr: str,
) -> str:
    resolved, count = re.subn(
        r"_[A-Za-z]{3}\s+Forecast\.xlsx$",
        f"_{target_month_abbr} Forecast.xlsx",
        str(baseline_name),
        count=1,
        flags=re.I,
    )
    if count != 1:
        raise ValueError(f"Could not resolve mapped workbook name: {baseline_name}")
    return resolved


def _close_consolidated(path):
    from automations.actuals.tableau_desktop_close_service import (
        close_tableau_forecast_workbook,
    )

    return close_tableau_forecast_workbook(path)


def populate_monthly_forecast_from_tableau(
    forecast_folder: str | Path,
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    routes: list[dict] | None = None,
    opener=None,
    waiter=None,
    extractor=None,
    closer=None,
    progress=None,
) -> dict:
    """
    Step 2: open a temp copy of the consolidated workbook set to the period,
    capture each mapped crosstab once, close it, then write every route's
    Excel sheet. One failing route never stops the others.

    The source workbook is never modified. ``opener``/``waiter``/
    ``extractor``/``closer`` default to the live Tableau functions.
    """
    from .tableau_service import (
        extract_forecast_worksheets,
        open_tableau_workbook,
        wait_for_tableau_workbook,
    )

    routes = FORECAST_ROUTES if routes is None else routes
    opener = opener or open_tableau_workbook
    waiter = waiter or wait_for_tableau_workbook
    extractor = extractor or extract_forecast_worksheets
    closer = closer or _close_consolidated

    folder = Path(forecast_folder)
    if not folder.is_dir():
        raise FileNotFoundError(
            f"Forecast Output Folder not found:\n{folder}\n\nRun step 1 first."
        )

    temporary = _temp_copy(tableau_workbook_path, "Forecast")
    opened = False
    close_warning = None
    try:
        period = write_consolidated_forecast_copy(
            tableau_workbook_path=tableau_workbook_path,
            destination=temporary,
            fiscal_year=fiscal_year,
            forecast_month=forecast_month,
            forecast_year=forecast_year,
        )
        if progress:
            progress(0, 1, "Opening the consolidated Tableau workbook")
        opener(tableau_workbook_path=temporary)
        opened = True
        waiter(tableau_workbook_path=temporary)

        worksheet_names = [route["tableau_worksheet"] for route in routes]
        extracts, extract_errors = extractor(
            tableau_workbook_path=temporary,
            worksheet_names=worksheet_names,
            progress=progress,
        )
    finally:
        if opened:
            try:
                close_warning = (closer(temporary) or {}).get("warning")
            except Exception as exc:
                close_warning = str(exc)
        if not close_warning:
            # A copy Tableau still holds is removed later by stale cleanup.
            temporary.unlink(missing_ok=True)

    audit_path = folder / f"Tableau_Forecast_Extract_{forecast_year}-{forecast_month:02d}.xlsx"
    write_tableau_audit_workbook(extracts, audit_path)

    passed = []
    failed = []

    for route in routes:
        target_name = _resolve_target_filename(
            route["workbook"],
            period["target_month_abbr"],
        )
        target = folder / target_name

        try:
            # Do not touch the Excel sheet if its Tableau source failed.
            if route["tableau_worksheet"] in extract_errors:
                raise RuntimeError(
                    "Tableau extraction failed: "
                    + extract_errors[route["tableau_worksheet"]]
                )
            if not target.exists():
                raise FileNotFoundError(
                    f"{target_name} is not in the Forecast Output Folder. "
                    "Run step 1 for this month first."
                )
            result = populate_forecast_sheet_from_audit(
                workbook_path=target,
                excel_sheet_name=route["excel_sheet"],
                audit_workbook_path=audit_path,
                tableau_worksheet_name=route["tableau_worksheet"],
                fiscal_year=period["fiscal_year"],
                expected_forecast_type=period["forecast_type"],
            )
            passed.append(
                {
                    "workbook": target_name,
                    "worksheet": route["excel_sheet"],
                    "tableau_worksheet": route["tableau_worksheet"],
                    "rows": result.get("tableau_detail_rows", 0),
                }
            )
        except Exception as exc:
            failed.append(
                {
                    "workbook": target_name,
                    "worksheet": route["excel_sheet"],
                    "tableau_worksheet": route["tableau_worksheet"],
                    "error": explain(exc),
                }
            )

    return {
        "period": period,
        "audit_workbook": str(audit_path),
        "passed": passed,
        "failed": failed,
        "close_warning": close_warning,
        "routes_processed": len(passed) + len(failed),
        "routes_passed": len(passed),
        "routes_failed": len(failed),
        "workbooks_updated": len({item["workbook"] for item in passed}),
        "worksheets_updated": len(passed),
    }


def validate_monthly_forecast(
    forecast_folder: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    tolerance: float = 5.00,
    progress=None,
) -> dict:
    """Step 3: compare each forecast sheet to its individual Tableau workbook.

    Every mapped route appears in the report (PASS, FAIL, NOT VALIDATED or
    ACTION REQUIRED); one bad route never stops the run.
    """
    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    # Fail before any Tableau work if the report folder is unavailable.
    validation_output_folder = get_forecast_validation_output_folder(
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        create=True,
    )

    result = run_grouped_individual_validation(
        routes=FORECAST_ROUTES,
        forecast_folder=forecast_folder,
        target_month_abbr=period["target_month_abbr"],
        forecast_month=forecast_month,
        forecast_year=forecast_year,
        forecast_type=period["forecast_type"],
        tolerance=tolerance,
        progress=progress,
    )

    report_path = write_forecast_validation_report(
        validation_result=result,
        output_folder=validation_output_folder,
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        forecast_type=period["forecast_type"],
        tolerance=tolerance,
    )

    result["period"] = period
    result["validation_output_folder"] = str(validation_output_folder)
    result["validation_report"] = str(report_path)
    return result
