from __future__ import annotations
from shared.friendly_errors import explain

from datetime import date, datetime

from pathlib import Path
import re

from openpyxl import load_workbook

# ============================================================
# GENERAL HELPERS
# ============================================================


def _norm(value) -> str:
    return " ".join(str(value or "").strip().lower().split())


def _number(value):
    """
    Convert Excel/Tableau numeric values such as:

        480,969
        $480,969
        (480,969)
        -480969
        480969.25

    to float.
    """

    if value is None:
        return 0.0

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()

    if not text:
        return 0.0

    negative_parentheses = text.startswith("(") and text.endswith(")")

    text = (
        text.replace(",", "").replace("$", "").replace("(", "").replace(")", "").strip()
    )

    if not text:
        return 0.0

    try:
        number = float(text)
    except ValueError as exc:
        raise ValueError(f"Could not convert value to number: {value!r}") from exc

    if negative_parentheses:
        number *= -1

    return number


# ============================================================
# RESOLVE MONTHLY FORECAST WORKBOOK NAME
# ============================================================


def resolve_forecast_workbook_name(
    baseline_name: str,
    target_month_abbr: str,
) -> str:
    """
    Convert the baseline mapping filename:

        FY27 OPEX Legal_Aug Forecast.xlsx

    to:

        FY27 OPEX Legal_Oct Forecast.xlsx
    """

    pattern = re.compile(
        r"_(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)" r"\s+Forecast\.xlsx$",
        re.IGNORECASE,
    )

    replacement = f"_{target_month_abbr} Forecast.xlsx"

    resolved = pattern.sub(
        replacement,
        str(baseline_name),
    )

    if resolved == baseline_name:
        raise ValueError(
            "Could not resolve monthly forecast filename from: " f"{baseline_name}"
        )

    return resolved


# ============================================================
# FIND EXCEL WORKSHEET
# ============================================================


def _find_excel_sheet(
    workbook,
    requested_name: str,
):
    """
    Resolve Excel sheet names using exact match first,
    then whitespace/case-normalized match.
    """

    if requested_name in workbook.sheetnames:
        return workbook[requested_name]

    requested_normalized = _norm(requested_name)

    for sheet_name in workbook.sheetnames:
        if _norm(sheet_name) == requested_normalized:
            return workbook[sheet_name]

    raise KeyError(
        f'Excel worksheet "{requested_name}" was not found. '
        f"Available sheets: {workbook.sheetnames}"
    )


# ============================================================
# EXCEL GRAND TOTAL
# ============================================================


def get_excel_forecast_grand_total(
    workbook_path: str | Path,
    worksheet_name: str,
) -> float:
    """The full-year Grand Total shown on the sheet: the total row's Grand Total cell.

    Uses the value Excel saved with the file when there is one. A workbook
    the app has just written has no saved values (openpyxl writes formulas
    only), so the cell's own formula is calculated from the sheet - the same
    number Excel shows when the file is opened.
    """
    from shared.friendly_errors import PlainLanguageError
    from shared.sheet_formulas import SheetCalculator, UnsupportedFormula

    from .excel_service import _detail_range, _s4_structure

    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise PlainLanguageError(
            f'The workbook "{workbook_path.name}" is not in the Forecast Output Folder. '
            "Run Step 1 and Step 2 first, or check the folder."
        )

    formulas = load_workbook(workbook_path, data_only=False)
    values = load_workbook(workbook_path, data_only=True)
    try:
        ws = _find_excel_sheet(formulas, worksheet_name)
        structure = _s4_structure(ws)
        _first, _last, total_row = _detail_range(ws, structure)
        if total_row is None:
            raise PlainLanguageError(
                f'The "{ws.title}" tab has no total row under its lines (a row with SUM '
                "formulas), so its Grand Total could not be read."
            )
        column = structure["grand_total_column"]

        saved = values[ws.title].cell(total_row, column).value
        if isinstance(saved, (int, float)) and not isinstance(saved, bool):
            return float(saved)

        try:
            return SheetCalculator(ws).value(total_row, column)
        except UnsupportedFormula as exc:
            raise PlainLanguageError(
                f'The Grand Total on the "{ws.title}" tab (cell '
                f"{ws.cell(total_row, column).coordinate}) uses a formula the app can't "
                f"calculate ({exc}). Open the workbook in Excel, save it, and validate again."
            ) from exc
    finally:
        formulas.close()
        values.close()


# ============================================================
# INDIVIDUAL TABLEAU GRAND TOTAL
# ============================================================


def get_individual_tableau_grand_total(
    tableau_rows: list[dict],
) -> float:
    """
    Return Tableau's authoritative final Grand Total.

    Stage 5A discovery confirmed the individual Legal
    crosstabs contain a final row where:

        Company = Grand Total

    and the final value is stored in:

        Grand Total

    Intermediate Tableau subtotal rows are intentionally ignored.
    """

    if not tableau_rows:
        raise RuntimeError("Individual Tableau crosstab returned no rows.")

    grand_total_row = None

    for row in tableau_rows:
        # Tableau puts the "Grand Total" label in the first column of the row
        # (Company on most tabs, GL Account Mapping on ALIV Head Office, ...),
        # so look in every column except the Grand Total value itself.
        labels = [value for key, value in row.items() if key != "Grand Total"]
        if any(_norm(value) == "grand total" for value in labels):
            grand_total_row = row

    if grand_total_row is None:
        raise RuntimeError(
            "Individual Tableau crosstab does not contain "
            'a row where Company = "Grand Total" (no Grand Total row in any column).'
        )

    if "Grand Total" not in grand_total_row:
        raise RuntimeError(
            "Individual Tableau Grand Total row does not "
            'contain a "Grand Total" column.'
        )

    return _number(grand_total_row.get("Grand Total"))


# ============================================================
# BUSINESS SIGN CONVENTION
# ============================================================


def apply_validation_sign_rule(
    tableau_total: float,
    forecast_workbook_name: str,
) -> float:
    """
    Stage 4 stores Revenue forecast values as the opposite sign
    of the Tableau source.

    Apply the same convention before validation so Revenue is
    compared like-for-like.
    """

    if "revenue" in str(forecast_workbook_name).lower():
        return tableau_total * -1

    return tableau_total


# ============================================================
# COMPARE TOTALS
# ============================================================


def compare_forecast_totals(
    excel_total: float,
    tableau_total: float,
    tolerance: float = 1000.00,
) -> dict:
    """
    Compare Excel and Tableau using a monetary tolerance.
    """

    excel_total = float(excel_total)
    tableau_total = float(tableau_total)

    variance = excel_total - tableau_total

    status = "PASS" if abs(variance) <= tolerance else "FAIL"

    return {
        "excel_grand_total": excel_total,
        "tableau_grand_total": tableau_total,
        "variance": variance,
        "status": status,
    }


# ============================================================
# VALIDATE ONE ROUTE
# ============================================================


def validate_route_from_tableau_rows(
    route: dict,
    forecast_folder: str | Path,
    target_month_abbr: str,
    tableau_rows: list[dict],
    tolerance: float = 1000.00,
) -> dict:
    """
    Validate one completed forecast route against already
    extracted individual Tableau rows.

    Routes requiring additional filtering are deliberately
    blocked from PASS/FAIL validation.
    """

    validation_status = route.get("validation_status")

    if validation_status != "confirmed":
        return {
            "forecast_workbook": (
                resolve_forecast_workbook_name(
                    route["workbook"],
                    target_month_abbr,
                )
            ),
            "worksheet": route["excel_sheet"],
            "individual_tableau_workbook": (route.get("individual_tableau_workbook")),
            "individual_tableau_worksheet": (route.get("individual_tableau_worksheet")),
            "excel_grand_total": None,
            "tableau_grand_total": None,
            "variance": None,
            "status": "NOT VALIDATED",
            "note": (
                "Not checked: the individual Tableau tab covers more than this sheet. "
                "It needs a "
                + (route.get("validation_filter_note") or "department")
                + " filter before its total can be compared."
            ),
        }

    forecast_folder = Path(forecast_folder)

    forecast_filename = resolve_forecast_workbook_name(
        route["workbook"],
        target_month_abbr,
    )

    forecast_workbook = forecast_folder / forecast_filename

    excel_total = get_excel_forecast_grand_total(
        workbook_path=forecast_workbook,
        worksheet_name=route["excel_sheet"],
    )

    raw_tableau_total = get_individual_tableau_grand_total(tableau_rows)

    adjusted_tableau_total = apply_validation_sign_rule(
        tableau_total=raw_tableau_total,
        forecast_workbook_name=(forecast_filename),
    )

    comparison = compare_forecast_totals(
        excel_total=excel_total,
        tableau_total=adjusted_tableau_total,
        tolerance=tolerance,
    )

    return {
        "forecast_workbook": (forecast_filename),
        "worksheet": route["excel_sheet"],
        "individual_tableau_workbook": (route["individual_tableau_workbook"]),
        "individual_tableau_worksheet": (route["individual_tableau_worksheet"]),
        "excel_grand_total": (comparison["excel_grand_total"]),
        "tableau_grand_total": (comparison["tableau_grand_total"]),
        "variance": comparison["variance"],
        "status": comparison["status"],
        "note": "",
    }


# ============================================================
# STAGE 5B - INDIVIDUAL TABLEAU VALIDATION
# ============================================================
#
# V1 intentionally preserves the proven old validation behavior:
#
#   * one temporary TWB per validation route
#   * Forecast Start = selected month's first day
#   * Forecast Type = FYxx Forecast_MMYYYY when the tab uses it
#   * otherwise Forecast Measure = "Forecast"
#   * Company = route company
#   * preserve all other saved Tableau selections/filters
#
# The public runner name is retained so the existing Stage 5B
# test file does not need to change.
# ============================================================


def _validation_parameter_key(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def _prepare_individual_forecast_copy(
    source: str | Path,
    destination: str | Path,
    tab: str,
    company: str,
    forecast_start: date,
    forecast_type: str | None = None,
    filter_company: bool = False,
) -> dict:
    """
    Prepare one route-specific temporary individual Tableau workbook.

    This follows the proven old _prepare_individual_forecast_copy rules:
    update the period controls (Forecast Start and Forecast Type / Forecast
    Measure) and, only for a group tab shared by ALIV and CBL
    (``filter_company``), the Company filter. Every other saved filter stays.
    """
    import xml.etree.ElementTree as ET

    source = Path(source)
    destination = Path(destination)

    if not source.is_file() or source.suffix.lower() != ".twb":
        raise FileNotFoundError(f"Individual Tableau TWB not found: {source}")

    tree = ET.parse(source)
    root = tree.getroot()

    # --------------------------------------------------------
    # Find the exact mapped worksheet.
    # --------------------------------------------------------
    target_sheets = [
        sheet
        for sheet in root.findall("./worksheets/worksheet")
        if sheet.get("name") == tab
    ]

    if len(target_sheets) != 1:
        raise ValueError(f"Individual Tableau worksheet not found: {tab}")

    target_sheet = target_sheets[0]

    # --------------------------------------------------------
    # FORECAST START
    # --------------------------------------------------------
    parameters = [
        column
        for column in root.iter("column")
        if any(
            _validation_parameter_key(column.get(field)) == "forecaststart"
            for field in ("caption", "name")
        )
    ]

    declarations = [
        column for column in parameters if "param-domain-type" in column.attrib
    ]
    if declarations:
        parameters = declarations

    # Resolve duplicate Forecast Start captions through the target
    # worksheet's dependencies, exactly as the old working code did.
    target_names = {
        column.get("name")
        for column in target_sheet.iter("column")
        if _validation_parameter_key(column.get("caption")) == "forecaststart"
    }

    if target_names:
        parameters = [
            column for column in parameters if column.get("name") in target_names
        ]

    identities = {column.get("name") or column.get("caption") for column in parameters}

    if not parameters or len(identities) != 1:
        available = sorted(
            {
                column.get("caption") or column.get("name") or "<unnamed>"
                for column in root.iter("column")
                if "param-domain-type" in column.attrib
            }
        )
        raise ValueError(
            "Cannot identify Forecast Start parameter. Available parameters: "
            + (", ".join(available) or "none found")
            + ". Share the individual TWB file if this name differs."
        )

    forecast_start_value = f"#{forecast_start:%Y-%m-%d}#"

    for parameter in parameters:
        parameter.set("value", forecast_start_value)

        calculation = parameter.find("calculation")
        if calculation is not None:
            calculation.set("formula", forecast_start_value)

        members = parameter.find("members")
        if members is not None and not any(
            member.get("value") == forecast_start_value for member in members
        ):
            ET.SubElement(
                members,
                "member",
                value=forecast_start_value,
            )

    filters = target_sheet.findall(".//filter")

    # --------------------------------------------------------
    # COMPANY - only for a "group" tab that serves both companies
    # (the mapping uses it for ALIV and for CBL). A tab built for one
    # company keeps its saved filters exactly as they are.
    # --------------------------------------------------------
    company = str(company or "").strip().upper()
    if filter_company:
        if company not in {"ALIV", "CBL"}:
            raise ValueError(
                f'Confirmed route requires ALIV/CBL company_filter: "{company}".'
            )
        company_filters = [
            item for item in filters if item.get("column", "").endswith("[none:Company:nk]")
        ]
        if len(company_filters) != 1:
            raise ValueError(
                f'Expected one Company filter on individual forecast tab "{tab}".'
            )
        company_filter = company_filters[0]
        if company_filter.get("class") != "categorical":
            raise ValueError(
                f'Unsupported Company filter type on individual forecast tab "{tab}".'
            )
        for child in list(company_filter):
            company_filter.remove(child)
        ET.SubElement(
            company_filter,
            "groupfilter",
            function="member",
            level="[none:Company:nk]",
            member=f'"{company}"',
        )

    # --------------------------------------------------------
    # FORECAST TYPE / FORECAST MEASURE
    # --------------------------------------------------------
    forecast_filters = [
        item
        for item in filters
        if item.get("column", "").endswith("[none:Forecast Type:nk]")
    ]

    if len(forecast_filters) > 1:
        raise ValueError(
            f'Multiple Forecast Type filters on individual forecast tab "{tab}".'
        )

    # Forecast Type comes from the selected forecast month (FY27
    # Forecast_102026 for October).
    if forecast_type:
        forecast_type_value = str(forecast_type).strip()
    else:
        fiscal_end = (
            forecast_start.year + 1 if forecast_start.month >= 7 else forecast_start.year
        )
        forecast_type_value = f"FY{str(fiscal_end)[-2:]} Forecast_{forecast_start:%m%Y}"

    if forecast_filters:
        selected = forecast_filters[0].findall(".//groupfilter[@function='member']")

        forecast_members = [
            member
            for member in selected
            if re.fullmatch(
                r"FY\d{2} Forecast_\d{6}",
                member.get("member", "").strip('"'),
                re.I,
            )
        ]

        if len(forecast_members) != 1:
            raise ValueError(
                f"Expected one monthly Forecast Type selection "
                f'in saved tab "{tab}".'
            )

        forecast_members[0].set(
            "member",
            f'"{forecast_type_value}"',
        )

        period_mode = "Forecast Type"

    else:
        # Old working fallback for worksheets that use a Forecast
        # Measure parameter instead of a Forecast Type filter.
        measure_names = {
            column.get("name")
            for column in target_sheet.iter("column")
            if _validation_parameter_key(column.get("caption")) == "forecastmeasure"
        }

        if len(measure_names) != 1:
            raise ValueError(
                f"No Forecast Type filter or unique Forecast Measure "
                f'parameter on forecast tab "{tab}".'
            )

        measure_name = next(iter(measure_names))

        calculations = [
            calculation.get("formula", "")
            for calculation in target_sheet.iter("calculation")
        ]

        if not any(
            measure_name in formula and "[Forecast]" in formula
            for formula in calculations
        ):
            raise ValueError(
                f"Forecast Measure is not linked to the worksheet "
                f'forecast calculation on "{tab}".'
            )

        # This block used to sit outside the else, so every tab with a
        # Forecast Type filter failed ("measure_name" not defined) and was
        # reported ACTION REQUIRED without Tableau ever opening.
        for column in root.iter("column"):
            if (
                column.get("name") == measure_name
                and "param-domain-type" in column.attrib
            ):
                column.set("value", '"Forecast"')
                calculation = column.find("calculation")
                if calculation is not None:
                    calculation.set("formula", '"Forecast"')

        period_mode = "Forecast Measure"

    # --------------------------------------------------------
    # SAVE TEMP COPY
    # --------------------------------------------------------
    tree.write(
        destination,
        encoding="utf-8",
        xml_declaration=True,
    )

    # --------------------------------------------------------
    # VERIFY THE PERIOD/COMPANY ACTUALLY SAVED
    # --------------------------------------------------------
    verify_tree = ET.parse(destination)
    verify_root = verify_tree.getroot()

    saved_start_values = {
        column.get("value")
        for column in verify_root.iter("column")
        if any(
            _validation_parameter_key(column.get(field)) == "forecaststart"
            for field in ("caption", "name")
        )
        and "param-domain-type" in column.attrib
    }

    if forecast_start_value not in saved_start_values:
        raise RuntimeError(
            "Forecast Start could not be verified in temporary Tableau copy: "
            f"{destination}"
        )

    verify_sheets = [
        sheet
        for sheet in verify_root.findall("./worksheets/worksheet")
        if sheet.get("name") == tab
    ]

    if len(verify_sheets) != 1:
        raise RuntimeError(
            f'Could not verify worksheet "{tab}" after saving temporary TWB.'
        )

    verify_sheet = verify_sheets[0]
    verify_filters = verify_sheet.findall(".//filter")

    if filter_company:
        saved_company_members = [
            member.get("member", "").strip('"')
            for item in verify_filters
            if item.get("column", "").endswith("[none:Company:nk]")
            for member in item.findall(".//groupfilter[@function='member']")
        ]
        if saved_company_members != [company]:
            raise RuntimeError(
                f'Company "{company}" could not be verified in temporary Tableau copy.'
            )

    if period_mode == "Forecast Type":
        verify_forecast_filters = [
            item
            for item in verify_filters
            if item.get("column", "").endswith("[none:Forecast Type:nk]")
        ]

        saved_forecast_members = []

        for item in verify_forecast_filters:
            saved_forecast_members.extend(
                member.get("member", "").strip('"')
                for member in item.findall(".//groupfilter[@function='member']")
            )

        if forecast_type_value not in saved_forecast_members:
            raise RuntimeError(
                "Forecast Type could not be verified in temporary Tableau copy.\n"
                f"Expected: {forecast_type_value}\n"
                f"Workbook: {destination}"
            )

    return {
        "source": source,
        "temporary": destination,
        "worksheet": tab,
        "company": company,
        "forecast_start": forecast_start,
        "forecast_start_value": forecast_start_value,
        "forecast_type": forecast_type_value,
        "period_mode": period_mode,
    }


def close_individual_tableau_workbook(
    tableau_workbook_path: str | Path,
) -> dict:
    """Close only the exact validation Tableau workbook, discarding UI changes."""
    import time
    from .tableau_service import _get_tableau_window, _Desktop

    path = Path(tableau_workbook_path)

    result = {
        "closed": False,
        "warning": None,
    }

    try:
        window = _get_tableau_window(
            max_attempts=1,
            retry_wait_seconds=0.2,
            tableau_workbook_path=path,
        )
    except Exception:
        # If the exact temp workbook is already gone, cleanup is complete.
        result["closed"] = True
        return result

    try:
        window.close()
        time.sleep(1.0)

        desktop = _Desktop(backend="uia")

        for candidate in desktop.windows():
            title = ""

            try:
                title = candidate.window_text().strip().lower()
            except Exception:
                pass

            if "tableau" not in title and "save" not in title:
                continue

            for button_name in (
                "Don't Save",
                "Don’t Save",
                "Discard",
                "No",
            ):
                try:
                    button = candidate.child_window(
                        title=button_name,
                        control_type="Button",
                    )

                    if button.exists(timeout=0.2):
                        button.click_input()
                        break
                except Exception:
                    continue

        time.sleep(1.0)
        result["closed"] = True

    except Exception as exc:
        result["warning"] = str(exc)

    return result


def run_grouped_individual_validation(
    routes: list[dict],
    forecast_folder: str | Path,
    target_month_abbr: str,
    forecast_month: int,
    forecast_year: int,
    only_workbooks: list[str] | None = None,
    tolerance: float = 1000.00,
    open_wait_seconds: float = 8.0,
    forecast_start: date | None = None,
    forecast_type: str | None = None,
    raw_output_folder: str | Path | None = None,
) -> dict:
    """
    Stage 5B controlled runner.

    ``raw_output_folder``: where a copy of each individual Tableau crosstab is
    kept (Forecast Validation\\Tableau Raw Copies\\Step 3), so a variance can be
    checked against exactly what Tableau showed.

    ``forecast_start`` / ``forecast_type`` come from the selected period
    (October 2026 -> 10/1/2026 and FY27 Forecast_102026).

    The function name is retained for compatibility with the existing
    Stage 5B test. Internally V1 intentionally validates route-by-route
    using one temporary TWB per route, matching the old proven behavior.
    """
    import time
    from uuid import uuid4

    from .tableau_service import (
        open_tableau_workbook,
        get_tableau_crosstab_rows,
    )

    allowed = None

    if only_workbooks:
        allowed = {_norm(name) for name in only_workbooks}

    selected_routes = []
    not_validated = []

    for route in routes:
        if route.get("validation_status") != "confirmed":
            # Listed in the report so nothing goes unchecked without a reason.
            item = validate_route_from_tableau_rows(
                route=route, forecast_folder=forecast_folder,
                target_month_abbr=target_month_abbr, tableau_rows=[], tolerance=tolerance,
            )
            item["category"] = route.get("category", "")
            not_validated.append(item)
            continue

        folder = route.get("individual_tableau_folder")
        workbook = route.get("individual_tableau_workbook")
        tab = route.get("individual_tableau_worksheet")

        if not folder or not workbook or not tab:
            continue

        if allowed is not None and _norm(workbook) not in allowed:
            continue

        selected_routes.append(route)

    results = []
    workbook_results = []

    # Individual tabs the mapping uses for both companies show the group
    # total, so those (only those) get the line's Company filter.
    from .mapping import FORECAST_ROUTES

    companies_by_tab: dict[tuple, set] = {}
    for item in [*FORECAST_ROUTES, *routes]:
        key = (_norm(item.get("individual_tableau_workbook")), _norm(item.get("individual_tableau_worksheet")))
        companies_by_tab.setdefault(key, set()).add(str(item.get("company_filter") or "").upper())

    if forecast_start is None:
        forecast_start = date(forecast_year, forecast_month, 1)

    for route in selected_routes:
        source = (
            Path(route["individual_tableau_folder"])
            / route["individual_tableau_workbook"]
        )

        from shared.tableau_temp_copies import temp_twb_path

        temporary = temp_twb_path(source, "Validation")

        opened = False

        workbook_record = {
            "source": str(source),
            "temporary": str(temporary),
            "routes": 1,
            "worksheet": route["individual_tableau_worksheet"],
            "company": route.get("company_filter"),
            "opened": False,
            "closed": False,
            "status": "PENDING",
            "error": "",
        }

        try:
            preparation = _prepare_individual_forecast_copy(
                source=source,
                destination=temporary,
                tab=route["individual_tableau_worksheet"],
                company=route.get("company_filter"),
                forecast_start=forecast_start,
                forecast_type=forecast_type,
                filter_company=len(
                    companies_by_tab.get(
                        (_norm(route["individual_tableau_workbook"]),
                         _norm(route["individual_tableau_worksheet"])),
                        set(),
                    ) & {"ALIV", "CBL"}
                ) > 1,
            )

            workbook_record["forecast_start"] = preparation["forecast_start_value"]
            workbook_record["forecast_type"] = preparation["forecast_type"]
            workbook_record["period_mode"] = preparation["period_mode"]

            open_tableau_workbook(
                tableau_workbook_path=temporary,
            )

            opened = True
            workbook_record["opened"] = True

            if open_wait_seconds:
                time.sleep(open_wait_seconds)

            tableau_rows = get_tableau_crosstab_rows(
                worksheet_name=route["individual_tableau_worksheet"],
                tableau_workbook_path=temporary,
                **({"raw_output_folder": raw_output_folder,
                    "sequence_number": len(results) + 1} if raw_output_folder else {}),
            )

            result = validate_route_from_tableau_rows(
                route=route,
                forecast_folder=forecast_folder,
                target_month_abbr=target_month_abbr,
                tableau_rows=tableau_rows,
                tolerance=tolerance,
            )

            result["forecast_start"] = preparation["forecast_start_value"]
            result["forecast_type"] = preparation["forecast_type"]
            result["period_mode"] = preparation["period_mode"]

            results.append(result)

            workbook_record["status"] = (
                "PASS" if result["status"] == "PASS" else "VALIDATED"
            )

        except Exception as exc:
            error = explain(exc)

            workbook_record["status"] = "ACTION REQUIRED"
            workbook_record["error"] = error

            results.append(
                {
                    "forecast_workbook": resolve_forecast_workbook_name(
                        route["workbook"],
                        target_month_abbr,
                    ),
                    "worksheet": route["excel_sheet"],
                    "individual_tableau_workbook": route.get(
                        "individual_tableau_workbook"
                    ),
                    "individual_tableau_worksheet": route.get(
                        "individual_tableau_worksheet"
                    ),
                    "excel_grand_total": None,
                    "tableau_grand_total": None,
                    "variance": None,
                    "status": "ACTION REQUIRED",
                    "note": error,
                }
            )

        finally:
            if opened:
                close_result = close_individual_tableau_workbook(temporary)
            else:
                close_result = {
                    "closed": True,
                    "warning": None,
                }

            workbook_record["closed"] = bool(close_result.get("closed"))

            warning = close_result.get("warning")

            if warning:
                if workbook_record["error"]:
                    workbook_record["error"] += f" | Close warning: {warning}"
                else:
                    workbook_record["error"] = f"Close warning: {warning}"

            # Delete only after Tableau has released the temp TWB.
            if workbook_record["closed"] and temporary.exists():
                try:
                    temporary.unlink()
                except Exception as exc:
                    explain(exc)
                    cleanup_warning = (
                        "A temporary file could not be removed; it can be "
                        "deleted manually."
                    )

                    if workbook_record["error"]:
                        workbook_record["error"] += f" | {cleanup_warning}"
                    else:
                        workbook_record["error"] = cleanup_warning

            workbook_results.append(workbook_record)

    # One result per selected route, in order: tag each with its category.
    for route, item in zip(selected_routes, results):
        item["category"] = route.get("category", "")

    results.extend(not_validated)

    passed = sum(item.get("status") == "PASS" for item in results)

    failed = sum(item.get("status") == "FAIL" for item in results)

    action_required = sum(item.get("status") == "ACTION REQUIRED" for item in results)

    return {
        "results": results,
        "workbooks": workbook_results,
        "passed": passed,
        "failed": failed,
        "action_required": action_required,
        "status": (
            "PASS"
            if results and failed == 0 and action_required == 0
            else "REVIEW REQUIRED"
        ),
    }
