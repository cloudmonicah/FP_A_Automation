from __future__ import annotations

from datetime import date, datetime

from pathlib import Path
import re

from shared.friendly_errors import PlainLanguageError, explain

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
# STATUSES AND ERRORS
# ============================================================
#
# PASS / FAIL       the route was compared.
# NOT VALIDATED     the forecast line itself is suspect (wrong month, a
#                   non-numeric or formula month cell) or the route is not
#                   configured for automated validation yet. Nothing was
#                   compared; the run carries on with the next route.
# ACTION REQUIRED   the automation could not finish this route (Tableau,
#                   missing file). The run carries on with the next route.


class SuspectForecastData(PlainLanguageError, ValueError):
    """A forecast line that must not be compared. The route is NOT VALIDATED."""


class WrongPeriodError(SuspectForecastData):
    """The workbook belongs to a different forecast month than the one selected."""


_FORECAST_TYPE = re.compile(r"FY(\d{2}) Forecast_(\d{2})(\d{4})", re.I)


def describe_forecast_type(value) -> str:
    """'FY27 Forecast_102026' -> 'October 2026'; anything else is returned as text."""
    match = _FORECAST_TYPE.fullmatch(str(value or "").strip())
    if not match:
        return f'"{value}"'
    month, year = int(match.group(2)), int(match.group(3))
    if not 1 <= month <= 12:
        return f'"{value}"'
    return f"{date(year, month, 1):%B %Y}"


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
    expected_forecast_type: str | None = None,
) -> float:
    """
    Calculate the final forecast Grand Total directly from the populated
    Jul-Jun detail values.

    Everything is found from headers: the "Forecast Type" header row, the
    twelve month columns (real dates written by Stage 4) and the detail rows
    (rows with a Forecast Type). No cell or column position is assumed.

    openpyxl does not calculate formulas, so Excel's cached totals are not
    used. A month cell that holds a formula or text, or a Forecast Type for a
    different month than ``expected_forecast_type``, raises
    SuspectForecastData and the route is reported NOT VALIDATED.
    """

    workbook_path = Path(workbook_path)

    if not workbook_path.exists():
        raise FileNotFoundError(f"Forecast workbook not found:\n{workbook_path}")

    wb = load_workbook(workbook_path, data_only=False, read_only=False)

    try:
        ws = _find_excel_sheet(wb, worksheet_name)

        header_row = forecast_type_col = None
        for row in range(1, min(ws.max_row, 50) + 1):
            for col in range(1, ws.max_column + 1):
                if _norm(ws.cell(row, col).value) == "forecast type":
                    header_row, forecast_type_col = row, col
                    break
            if header_row is not None:
                break

        if header_row is None:
            raise RuntimeError(
                f'Could not find "Forecast Type" header in worksheet "{worksheet_name}".'
            )

        month_columns = [
            col
            for col in range(1, ws.max_column + 1)
            if isinstance(ws.cell(header_row, col).value, (datetime, date))
        ]
        if len(month_columns) != 12:
            raise RuntimeError(
                f'Expected 12 fiscal month columns in "{worksheet_name}", '
                f"but found {len(month_columns)}. Run step 2 (Update Workbooks "
                "from Tableau) for this workbook first."
            )

        # The bottom total row has no Forecast Type, so it is excluded.
        detail_rows = [
            row
            for row in range(header_row + 1, ws.max_row + 1)
            if ws.cell(row, forecast_type_col).value not in (None, "")
        ]
        if not detail_rows:
            raise RuntimeError(f'No forecast detail rows found in "{worksheet_name}".')

        if expected_forecast_type:
            wanted = _norm(expected_forecast_type)
            found = {
                str(ws.cell(row, forecast_type_col).value).strip()
                for row in detail_rows
            }
            wrong = sorted(value for value in found if _norm(value) != wanted)
            if wrong:
                raise WrongPeriodError(
                    f"{workbook_path.name} / {worksheet_name} is for "
                    f"{describe_forecast_type(wrong[0])} ({wrong[0]}), but the "
                    f"selected forecast month is "
                    f"{describe_forecast_type(expected_forecast_type)} "
                    f"({expected_forecast_type}). Check the Forecast Output "
                    "Folder, or run steps 1 and 2 for the selected month."
                )

        grand_total = 0.0
        for row in detail_rows:
            for col in month_columns:
                cell = ws.cell(row, col)
                value = cell.value
                if value in (None, ""):
                    continue
                if isinstance(value, str) and value.startswith("="):
                    raise SuspectForecastData(
                        f"{workbook_path.name} / {worksheet_name} cell "
                        f"{cell.coordinate} holds a formula ({value}); the month "
                        "values should be numbers written by step 2."
                    )
                try:
                    grand_total += _number(value)
                except ValueError as exc:
                    raise SuspectForecastData(
                        f"{workbook_path.name} / {worksheet_name} cell "
                        f"{cell.coordinate} is not a number: {value!r}"
                    ) from exc

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
    Return Tableau's final Grand Total from an individual crosstab.

    The crosstab ends with a row labelled "Grand Total" (normally in the
    Company column) and the amount sits in the "Grand Total" (or "Total")
    column. Both are found by name, so extra or reordered columns in the
    Tableau view do not matter. Intermediate subtotal rows are ignored.
    """

    if not tableau_rows:
        raise RuntimeError("Individual Tableau crosstab returned no rows.")

    headers: list = []
    for row in tableau_rows:
        headers.extend(key for key in row if key not in headers)

    total_column = next(
        (h for h in headers if _norm(h) == "grand total"),
        next((h for h in headers if _norm(h) == "total"), None),
    )
    if total_column is None:
        raise RuntimeError(
            'Individual Tableau crosstab has no "Grand Total" column. '
            f"Columns found: {', '.join(map(str, headers))}"
        )

    grand_total_row = None
    for row in tableau_rows:
        if any(
            _norm(value) == "grand total"
            for key, value in row.items()
            if key != total_column
        ):
            grand_total_row = row

    if grand_total_row is None:
        raise RuntimeError(
            'Individual Tableau crosstab has no "Grand Total" row. Check that '
            "the worksheet shows column grand totals (Analysis > Totals)."
        )

    return _number(grand_total_row.get(total_column))


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


def _route_row(route: dict, target_month_abbr: str, status: str, note: str = "") -> dict:
    """A result row for a route that was not compared (no totals)."""
    try:
        workbook = resolve_forecast_workbook_name(route["workbook"], target_month_abbr)
    except Exception:
        workbook = route.get("workbook")
    return {
        "forecast_workbook": workbook,
        "worksheet": route.get("excel_sheet"),
        "individual_tableau_workbook": route.get("individual_tableau_workbook"),
        "individual_tableau_worksheet": route.get("individual_tableau_worksheet"),
        "excel_grand_total": None,
        "tableau_grand_total": None,
        "variance": None,
        "status": status,
        "note": note,
    }


def validate_route_from_tableau_rows(
    route: dict,
    forecast_folder: str | Path,
    target_month_abbr: str,
    tableau_rows: list[dict],
    tolerance: float = 5.00,
    expected_forecast_type: str | None = None,
    excel_total: float | None = None,
) -> dict:
    """
    Validate one completed forecast route against already
    extracted individual Tableau rows.

    Routes requiring additional filtering are deliberately
    blocked from PASS/FAIL validation.
    """

    if route.get("validation_status") != "confirmed":
        return _route_row(
            route,
            target_month_abbr,
            "NOT VALIDATED",
            route.get("validation_filter_note")
            or "Additional Tableau filtering is required.",
        )

    forecast_filename = resolve_forecast_workbook_name(
        route["workbook"],
        target_month_abbr,
    )

    if excel_total is None:
        excel_total = get_excel_forecast_grand_total(
            workbook_path=Path(forecast_folder) / forecast_filename,
            worksheet_name=route["excel_sheet"],
            expected_forecast_type=expected_forecast_type,
        )

    raw_tableau_total = get_individual_tableau_grand_total(tableau_rows)

    adjusted_tableau_total = apply_validation_sign_rule(
        tableau_total=raw_tableau_total,
        forecast_workbook_name=forecast_filename,
    )

    comparison = compare_forecast_totals(
        excel_total=excel_total,
        tableau_total=adjusted_tableau_total,
        tolerance=tolerance,
    )

    return {
        "forecast_workbook": forecast_filename,
        "worksheet": route["excel_sheet"],
        "individual_tableau_workbook": route["individual_tableau_workbook"],
        "individual_tableau_worksheet": route["individual_tableau_worksheet"],
        "excel_grand_total": comparison["excel_grand_total"],
        "tableau_grand_total": comparison["tableau_grand_total"],
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

    ``forecast_type`` is the period's "FYxx Forecast_MMYYYY" value; pass the
    one built from the GUI selection so Excel, the consolidated workbook and
    this copy always agree.

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

    if forecast_type:
        forecast_type_value = str(forecast_type).strip()
    else:
        fiscal_end = (
            forecast_start.year + 1
            if forecast_start.month >= 7
            else forecast_start.year
        )
        forecast_type_value = (
            f"FY{str(fiscal_end)[-2:]} Forecast_{forecast_start:%m%Y}"
        )

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

        # This block used to sit outside the else, so every Forecast Type
        # tab failed with "measure_name referenced before assignment".
        declarations_updated = 0

        for column in root.iter("column"):
            if (
                column.get("name") == measure_name
                and "param-domain-type" in column.attrib
            ):
                column.set("value", '"Forecast"')
                calculation = column.find("calculation")
                if calculation is not None:
                    calculation.set("formula", '"Forecast"')
                declarations_updated += 1

        if not declarations_updated:
            raise ValueError(
                f'Forecast Measure parameter declaration not found for "{tab}".'
            )

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


def _temp_twb_path(source: Path) -> Path:
    """Temp copy location (never beside the source workbook)."""
    from shared.tableau_temp_copies import temp_twb_path

    return temp_twb_path(source, "Validation")


def _default_waiter(settle_seconds: float):
    def wait(tableau_workbook_path):
        from .tableau_service import wait_for_tableau_workbook

        wait_for_tableau_workbook(
            tableau_workbook_path, settle_seconds=settle_seconds
        )

    return wait


def run_grouped_individual_validation(
    routes: list[dict],
    forecast_folder: str | Path,
    target_month_abbr: str,
    forecast_month: int,
    forecast_year: int,
    only_workbooks: list[str] | None = None,
    tolerance: float = 5.00,
    open_wait_seconds: float = 3.0,
    forecast_type: str | None = None,
    opener=None,
    waiter=None,
    capture=None,
    closer=None,
    progress=None,
) -> dict:
    """
    Stage 5B: validate every route, one temporary individual TWB per route.

    Every route in ``routes`` gets exactly one result row, in mapping order:

    * routes not configured for automated validation -> NOT VALIDATED
    * Excel line is suspect (wrong month, text/formula month cell)
      -> NOT VALIDATED, checked BEFORE Tableau is opened
    * Tableau or file problems -> ACTION REQUIRED
    * otherwise PASS / FAIL against ``tolerance``

    A problem with one route never stops the run.

    ``opener``, ``waiter``, ``capture`` and ``closer`` default to the live
    Tableau Desktop functions and can be replaced in tests. ``progress`` is
    called as progress(done, total, label) from this (worker) thread; the
    GUI must hand it to Tk with ``after``.
    """
    from .tableau_service import get_tableau_crosstab_rows, open_tableau_workbook

    opener = opener or open_tableau_workbook
    waiter = waiter or _default_waiter(open_wait_seconds)
    capture = capture or get_tableau_crosstab_rows
    closer = closer or close_individual_tableau_workbook

    allowed = {_norm(name) for name in only_workbooks} if only_workbooks else None
    selected = [
        route
        for route in routes
        if allowed is None
        or _norm(route.get("individual_tableau_workbook")) in allowed
        or _norm(route.get("workbook")) in allowed
    ]

    forecast_start = date(forecast_year, forecast_month, 1)
    results = []
    workbook_results = []
    total = len(selected)

    for index, route in enumerate(selected, start=1):
        label = f"{route.get('workbook', '?')} / {route.get('excel_sheet', '?')}"
        if progress:
            progress(index - 1, total, label)

        folder = route.get("individual_tableau_folder")
        workbook = route.get("individual_tableau_workbook")
        tab = route.get("individual_tableau_worksheet")

        if route.get("validation_status") != "confirmed":
            results.append(
                _route_row(
                    route,
                    target_month_abbr,
                    "NOT VALIDATED",
                    route.get("validation_filter_note")
                    or "This route is not set up for automated validation yet.",
                )
            )
            continue

        if not folder or not workbook or not tab:
            results.append(
                _route_row(
                    route,
                    target_month_abbr,
                    "NOT VALIDATED",
                    "The individual Tableau workbook/worksheet is missing from "
                    "the forecast mapping.",
                )
            )
            continue

        # Check the Excel side first: a wrong-month or suspect workbook is
        # reported without spending minutes in Tableau.
        try:
            forecast_filename = resolve_forecast_workbook_name(
                route["workbook"], target_month_abbr
            )
            excel_total = get_excel_forecast_grand_total(
                workbook_path=Path(forecast_folder) / forecast_filename,
                worksheet_name=route["excel_sheet"],
                expected_forecast_type=forecast_type,
            )
        except SuspectForecastData as exc:
            results.append(
                _route_row(route, target_month_abbr, "NOT VALIDATED", str(exc))
            )
            continue
        except Exception as exc:
            results.append(
                _route_row(
                    route,
                    target_month_abbr,
                    "ACTION REQUIRED",
                    explain(exc),
                )
            )
            continue

        source = Path(folder) / workbook
        temporary = None
        opened = False

        workbook_record = {
            "source": str(source),
            "temporary": "",
            "routes": 1,
            "worksheet": tab,
            "company": route.get("company_filter"),
            "opened": False,
            "closed": False,
            "status": "PENDING",
            "error": "",
        }

        try:
            temporary = _temp_twb_path(source)
            workbook_record["temporary"] = str(temporary)

            preparation = _prepare_individual_forecast_copy(
                source=source,
                destination=temporary,
                tab=tab,
                company=route.get("company_filter"),
                forecast_start=forecast_start,
                forecast_type=forecast_type,
            )

            workbook_record["forecast_start"] = preparation["forecast_start_value"]
            workbook_record["forecast_type"] = preparation["forecast_type"]
            workbook_record["period_mode"] = preparation["period_mode"]

            opener(tableau_workbook_path=temporary)
            opened = True
            workbook_record["opened"] = True
            waiter(tableau_workbook_path=temporary)

            tableau_rows = capture(
                worksheet_name=tab,
                tableau_workbook_path=temporary,
            )

            result = validate_route_from_tableau_rows(
                route=route,
                forecast_folder=forecast_folder,
                target_month_abbr=target_month_abbr,
                tableau_rows=tableau_rows,
                tolerance=tolerance,
                excel_total=excel_total,
            )

            result["forecast_start"] = preparation["forecast_start_value"]
            result["forecast_type"] = preparation["forecast_type"]
            result["period_mode"] = preparation["period_mode"]
            results.append(result)
            workbook_record["status"] = result["status"]

        except Exception as exc:
            error = explain(exc)
            workbook_record["status"] = "ACTION REQUIRED"
            workbook_record["error"] = error
            results.append(
                _route_row(route, target_month_abbr, "ACTION REQUIRED", error)
            )

        finally:
            warning = None
            if opened:
                try:
                    close_result = closer(temporary) or {}
                except Exception as exc:
                    close_result = {"closed": False, "warning": str(exc)}
                workbook_record["closed"] = bool(close_result.get("closed"))
                warning = close_result.get("warning")
            else:
                workbook_record["closed"] = True

            if warning:
                workbook_record["error"] = " | ".join(
                    part for part in (workbook_record["error"], f"Close warning: {warning}") if part
                )

            # Delete only after Tableau has released the temp TWB. If it
            # could not be closed, stale-copy cleanup removes it later.
            if temporary is not None and workbook_record["closed"]:
                try:
                    temporary.unlink(missing_ok=True)
                except Exception as exc:
                    explain(exc)  # technical detail to the log only
                    workbook_record["error"] = " | ".join(
                        part
                        for part in (
                            workbook_record["error"],
                            "A temporary Tableau copy could not be removed; "
                            "it is deleted automatically later.",
                        )
                        if part
                    )

            workbook_results.append(workbook_record)

    if progress:
        progress(total, total, "Validation finished")

    return summarize_results(results, workbook_results)


def summarize_results(results: list[dict], workbook_results: list[dict] | None = None) -> dict:
    counts = {
        status: sum(item.get("status") == status for item in results)
        for status in ("PASS", "FAIL", "ACTION REQUIRED", "NOT VALIDATED")
    }
    return {
        "results": results,
        "workbooks": workbook_results or [],
        "passed": counts["PASS"],
        "failed": counts["FAIL"],
        "action_required": counts["ACTION REQUIRED"],
        "not_validated": counts["NOT VALIDATED"],
        "status": (
            "PASS"
            if results and counts["PASS"] == len(results)
            else "REVIEW REQUIRED"
        ),
    }
