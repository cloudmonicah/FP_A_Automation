from __future__ import annotations

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
from .tableau_service import (
    extract_forecast_worksheets,
    open_tableau_workbook,
    set_forecast_start_parameter,
)
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


def set_consolidated_forecast_period(
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """
    Set the consolidated FY Forecast workbook to the GUI-selected period.

    Forecast Start uses the already-tested tableau_service implementation.
    Forecast Type updates saved categorical member selections that contain
    the monthly FYxx Forecast_MMYYYY value.
    """
    path = Path(tableau_workbook_path)
    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    # Proven Forecast Start updater + disk verification.
    set_forecast_start_parameter(
        tableau_workbook_path=path,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    tree = ET.parse(path)
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

    tree.write(path, encoding="utf-8", xml_declaration=True)

    # Verify Forecast Type was actually written where applicable.
    verify = ET.parse(path).getroot()
    saved_members = [
        str(item.get("member") or "").strip('"')
        for item in verify.iter("groupfilter")
        if re.fullmatch(
            r"FY\d{2} Forecast_\d{6}",
            str(item.get("member") or "").strip('"'),
            re.I,
        )
    ]

    if changed and target not in saved_members:
        raise RuntimeError(
            "Forecast Type was updated but could not be verified.\n"
            f"Expected: {target}"
        )

    period["forecast_type_updates"] = changed
    return period


def prepare_monthly_forecast(
    source_folder: str | Path,
    output_folder: str | Path,
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    overwrite_mode: str = "skip",
    open_tableau: bool = True,
) -> dict:
    """Stage 1 + set the consolidated Tableau period + optionally open Tableau."""
    period = build_forecast_period(
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

    period = set_consolidated_forecast_period(
        tableau_workbook_path=tableau_workbook_path,
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    if open_tableau:
        open_tableau_workbook(tableau_workbook_path=tableau_workbook_path)

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


def populate_monthly_forecast_from_tableau(
    forecast_folder: str | Path,
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """
    Capture the consolidated Tableau crosstabs once and run the locked Stage 4
    population/finalization logic against every production mapping route.
    """
    folder = Path(forecast_folder)
    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    # Re-apply the selected period immediately before extraction.
    period = set_consolidated_forecast_period(
        tableau_workbook_path=tableau_workbook_path,
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    worksheet_names = [route["tableau_worksheet"] for route in FORECAST_ROUTES]

    extracts, extract_errors = extract_forecast_worksheets(
        tableau_workbook_path=tableau_workbook_path,
        worksheet_names=worksheet_names,
    )

    audit_path = folder / (
        f"Tableau_Forecast_Extract_" f"{forecast_year}-{forecast_month:02d}.xlsx"
    )
    write_tableau_audit_workbook(extracts, audit_path)

    passed = []
    failed = []

    for route in FORECAST_ROUTES:
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
            result = populate_forecast_sheet_from_audit(
                workbook_path=target,
                excel_sheet_name=route["excel_sheet"],
                audit_workbook_path=audit_path,
                tableau_worksheet_name=route["tableau_worksheet"],
                fiscal_year=fiscal_year,
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
                    "error": f"{type(exc).__name__}: {exc}",
                }
            )

    return {
        "period": period,
        "audit_workbook": str(audit_path),
        "passed": passed,
        "failed": failed,
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
    tolerance: float = 1000.00,
) -> dict:

    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    # ---------------------------------------------------------
    # STACY VALIDATION OUTPUT LOCATION
    # ---------------------------------------------------------

    validation_output_folder = get_forecast_validation_output_folder(
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        create=True,
    )

    # ---------------------------------------------------------
    # RUN EXISTING STAGE 5 VALIDATION
    # ---------------------------------------------------------

    result = run_grouped_individual_validation(
        routes=FORECAST_ROUTES,
        forecast_folder=forecast_folder,
        target_month_abbr=period["target_month_abbr"],
        forecast_month=forecast_month,
        forecast_year=forecast_year,
        tolerance=tolerance,
    )

    # ---------------------------------------------------------
    # CREATE STACY VALIDATION WORKBOOK
    # ---------------------------------------------------------

    report_path = write_forecast_validation_report(
        validation_result=result,
        output_folder=validation_output_folder,
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        forecast_type=period["forecast_type"],
        tolerance=tolerance,
    )

    # ---------------------------------------------------------
    # RETURN EVERYTHING TO GUI
    # ---------------------------------------------------------

    result["period"] = period

    result["validation_output_folder"] = str(validation_output_folder)

    result["validation_report"] = str(report_path)

    return result
