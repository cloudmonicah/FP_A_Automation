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
from .mapping import FORECAST_CATEGORIES, routes_for
from .tableau_service import (
    extract_forecast_worksheets,
    open_tableau_workbook,
    set_forecast_start_parameter,
)
from shared.forecast_output_paths import (
    get_forecast_validation_output_folder,
)

from shared.forecast_validation_report import (
    merge_saved_results,
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

    # Forecast Start is the 1st of the forecast month (October 2026 -> 10/1/2026).
    # Live run: 10/1 brought all 12 months from Tableau; 9/1 brought only Jul-Aug.
    month_start = start = date(forecast_year, forecast_month, 1)
    forecast_type = f"{fiscal_year} Forecast_{forecast_month:02d}{forecast_year}"

    return {
        "fiscal_year": fiscal_year,
        "forecast_month": forecast_month,
        "forecast_year": forecast_year,
        "forecast_start": start,
        "forecast_start_tableau": f"#{start:%Y-%m-%d}#",
        "forecast_type": forecast_type,
        "target_month_abbr": MONTH_ABBR[forecast_month],
        "display": f"{month_start:%B %Y}",
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
        forecast_month=period["forecast_start"].month,
        forecast_year=period["forecast_start"].year,
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


def _make_sure_tableau_is_open(tableau_workbook_path) -> None:
    """Open the consolidated workbook if step 1 didn't (or it was closed)."""
    from .tableau_service import _get_tableau_window, wait_for_tableau_workbook

    try:
        _get_tableau_window(
            max_attempts=1,
            retry_wait_seconds=0,
            tableau_workbook_path=tableau_workbook_path,
        )
        return
    except Exception:
        pass
    print("The consolidated Tableau workbook is not open; opening it now...")
    open_tableau_workbook(tableau_workbook_path=tableau_workbook_path)
    wait_for_tableau_workbook(tableau_workbook_path)


def populate_monthly_forecast_from_tableau(
    forecast_folder: str | Path,
    tableau_workbook_path: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    categories: list[str] | None = None,
) -> dict:
    """
    Step 2 for one category (OPEX, Direct Cost, Revenue, Other) or all.

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

    routes = routes_for(categories)
    worksheet_names = [route["tableau_worksheet"] for route in routes]

    _make_sure_tableau_is_open(tableau_workbook_path)

    extracts, extract_errors = extract_forecast_worksheets(
        tableau_workbook_path=tableau_workbook_path,
        worksheet_names=worksheet_names,
    )

    # One extract file per category, so running Revenue never overwrites the
    # OPEX extract.
    tag = "All" if not categories else "_".join(c.replace(" ", "") for c in categories)
    audit_path = folder / (
        f"Tableau_Forecast_Extract_{forecast_year}-{forecast_month:02d}_{tag}.xlsx"
    )
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
                    "months_pasted": result.get("months_pasted", []),
                    "months_blank_in_tableau": result.get("months_blank_in_tableau", []),
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

    pasted = []
    for item in passed:
        for month in item.get("months_pasted", []):
            if month not in pasted:
                pasted.append(month)

    return {
        "period": period,
        "categories": list(categories or FORECAST_CATEGORIES),
        "months_pasted": ", ".join(pasted),
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
    tolerance: float = 5.00,
    categories: list[str] | None = None,
) -> dict:
    """Step 3 for one category (OPEX, Direct Cost, Revenue, Other) or all.

    The month's report keeps the latest results of every category: running
    Revenue replaces Revenue's rows and leaves the other categories alone.
    """
    period = build_forecast_period(
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
        forecast_year=forecast_year,
    )

    validation_output_folder = get_forecast_validation_output_folder(
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        create=True,
    )

    result = run_grouped_individual_validation(
        routes=routes_for(categories),
        forecast_folder=forecast_folder,
        target_month_abbr=period["target_month_abbr"],
        forecast_month=forecast_month,
        forecast_year=forecast_year,
        tolerance=tolerance,
        forecast_start=period["forecast_start"],
        forecast_type=period["forecast_type"],
    )

    all_results = merge_saved_results(
        output_folder=validation_output_folder,
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        forecast_type=period["forecast_type"],
        categories_run=list(categories or FORECAST_CATEGORIES),
        results=result["results"],
    )

    report_path = write_forecast_validation_report(
        validation_result={**result, "results": all_results},
        output_folder=validation_output_folder,
        fiscal_year=period["fiscal_year"],
        forecast_month=period["forecast_month"],
        forecast_year=period["forecast_year"],
        forecast_type=period["forecast_type"],
        tolerance=tolerance,
        forecast_start=period["forecast_start"],
    )

    result["categories"] = list(categories or FORECAST_CATEGORIES)
    result["period"] = period
    result["validation_output_folder"] = str(validation_output_folder)
    result["validation_report"] = str(report_path)
    return result
