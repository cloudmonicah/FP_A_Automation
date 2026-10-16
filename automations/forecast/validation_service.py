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
    """
    Calculate the final forecast Grand Total directly from the
    populated Jul-Jun detail values.

    We deliberately do NOT depend on Excel's cached formula
    result because openpyxl writes formulas but does not
    calculate them.
    """

    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        raise FileNotFoundError(f"Forecast workbook not found:\n{workbook_path}")

    # --------------------------------------------------------
    # data_only=False is intentional.
    #
    # Stage 4 populated the monthly values directly. We read
    # those values and calculate the total ourselves instead
    # of depending on Excel's formula cache.
    # --------------------------------------------------------

    wb = load_workbook(
        workbook_path,
        data_only=False,
        read_only=False,
    )

    try:
        ws = _find_excel_sheet(
            wb,
            worksheet_name,
        )

        header_row = None
        forecast_type_col = None

        # ----------------------------------------------------
        # Find Forecast Type header.
        # ----------------------------------------------------

        for row in range(
            1,
            min(ws.max_row, 25) + 1,
        ):
            for col in range(
                1,
                ws.max_column + 1,
            ):
                if (
                    _norm(
                        ws.cell(
                            row,
                            col,
                        ).value
                    )
                    == "forecast type"
                ):
                    header_row = row
                    forecast_type_col = col
                    break

            if header_row is not None:
                break

        if header_row is None:
            raise RuntimeError(
                f'Could not find "Forecast Type" header '
                f'in worksheet "{worksheet_name}".'
            )

        # ----------------------------------------------------
        # Identify Jul-Jun month columns from the real Excel
        # dates written by Stage 4.
        # ----------------------------------------------------

        month_columns = []

        for col in range(
            1,
            ws.max_column + 1,
        ):
            value = ws.cell(
                header_row,
                col,
            ).value

            if isinstance(
                value,
                (datetime, date),
            ):
                month_columns.append(col)

        if len(month_columns) != 12:
            raise RuntimeError(
                f"Expected 12 fiscal month columns in "
                f'"{worksheet_name}", but found '
                f"{len(month_columns)}."
            )

        # ----------------------------------------------------
        # Find the actual detail rows.
        #
        # The bottom total row deliberately has no Forecast
        # Type value, so it is automatically excluded.
        # ----------------------------------------------------

        detail_rows = []

        for row in range(
            header_row + 1,
            ws.max_row + 1,
        ):
            forecast_type = ws.cell(
                row,
                forecast_type_col,
            ).value

            if forecast_type not in (
                None,
                "",
            ):
                detail_rows.append(row)

        if not detail_rows:
            raise RuntimeError(
                f"No forecast detail rows found in " f'"{worksheet_name}".'
            )

        # ----------------------------------------------------
        # Calculate the workbook total directly from Jul-Jun.
        # ----------------------------------------------------

        grand_total = 0.0

        for row in detail_rows:
            for col in month_columns:
                value = ws.cell(
                    row,
                    col,
                ).value

                if value in (
                    None,
                    "",
                ):
                    continue

                # Stage 4 monthly cells should contain values,
                # not formulas.
                if isinstance(value, str) and value.startswith("="):
                    raise RuntimeError(
                        "Unexpected formula found in a "
                        "forecast month cell.\n"
                        f"Workbook: {workbook_path.name}\n"
                        f"Worksheet: {worksheet_name}\n"
                        f"Cell: {ws.cell(row, col).coordinate}\n"
                        f"Value: {value}"
                    )

                grand_total += _number(value)

        return grand_total

    finally:
        wb.close()


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
        if _norm(row.get("Company")) == "grand total":
            grand_total_row = row

    if grand_total_row is None:
        raise RuntimeError(
            "Individual Tableau crosstab does not contain "
            'a row where Company = "Grand Total".'
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
    tolerance: float = 5.00,
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
    tolerance: float = 5.00,
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
                route.get("validation_filter_note")
                or "Additional Tableau filtering is required."
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
) -> dict:
    """
    Prepare one route-specific temporary individual Tableau workbook.

    This follows the proven old _prepare_individual_forecast_copy rules:
    update only the monthly period controls and Company selection while
    retaining the worksheet's other saved filters/selections.
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

    # --------------------------------------------------------
    # COMPANY
    #
    # Preserve the old rule for V1. Workbooks without this
    # structure remain ACTION REQUIRED until their actual filter
    # structure is inspected; do not invent a replacement rule.
    # --------------------------------------------------------
    company = str(company or "").strip().upper()

    if company not in {"ALIV", "CBL"}:
        raise ValueError(
            f'Confirmed route requires ALIV/CBL company_filter: "{company}".'
        )

    filters = target_sheet.findall(".//filter")

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

    # Replace only Company. Leave every other saved filter alone.
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
    # Forecast_102026).
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

    verify_company_filters = [
        item
        for item in verify_filters
        if item.get("column", "").endswith("[none:Company:nk]")
    ]

    if len(verify_company_filters) != 1:
        raise RuntimeError(f'Company filter could not be verified on "{tab}".')

    saved_company_members = [
        member.get("member", "").strip('"')
        for member in verify_company_filters[0].findall(
            ".//groupfilter[@function='member']"
        )
    ]

    if company not in saved_company_members:
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
    tolerance: float = 5.00,
    open_wait_seconds: float = 8.0,
    forecast_start: date | None = None,
    forecast_type: str | None = None,
) -> dict:
    """
    Stage 5B controlled runner.

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

    for route in routes:
        if route.get("validation_status") != "confirmed":
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
