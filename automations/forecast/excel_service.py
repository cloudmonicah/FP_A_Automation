from __future__ import annotations
from shared.friendly_errors import PlainLanguageError, explain

from pathlib import Path
from typing import Any
from copy import copy
from datetime import datetime
import calendar
import re
import shutil

from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# MONTH / FISCAL CONFIGURATION
# ============================================================

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

FISCAL_MONTH_ORDER = {
    7: 1,
    8: 2,
    9: 3,
    10: 4,
    11: 5,
    12: 6,
    1: 7,
    2: 8,
    3: 9,
    4: 10,
    5: 11,
    6: 12,
}

MONTH_NUMBER_BY_ABBR = {
    abbreviation.lower(): month_number
    for month_number, abbreviation in MONTH_ABBR.items()
}


# ============================================================
# GENERAL HELPERS
# ============================================================


def normalize_text(value: Any) -> str:
    return str(value or "").strip().lower()


def is_formula(value: Any) -> bool:
    return isinstance(value, str) and value.startswith("=")


def get_previous_month(forecast_month: int) -> int:
    if forecast_month == 1:
        return 12

    return forecast_month - 1


# ============================================================
# FORECAST FILE DISCOVERY
# ============================================================


def parse_forecast_filename(
    file_path: Path,
    fiscal_year: str,
) -> dict | None:
    """
    Parse a forecast workbook filename.

    Example:
        FY27 OPEX Human Resources_Sep Forecast.xlsx
    """

    if file_path.name.startswith("~$"):
        return None

    if file_path.suffix.lower() not in {
        ".xlsx",
        ".xlsm",
        ".xls",
        ".xlsb",
    }:
        return None

    if not file_path.name.upper().startswith(fiscal_year.upper()):
        return None

    pattern = re.compile(
        r"^(?P<family>.+)_"
        r"(?P<month>"
        r"Jan|Feb|Mar|Apr|May|Jun|"
        r"Jul|Aug|Sep|Oct|Nov|Dec"
        r") Forecast$",
        re.IGNORECASE,
    )

    match = pattern.match(file_path.stem)

    if not match:
        return None

    family = match.group("family").strip()
    month_abbr = match.group("month").title()
    month_number = MONTH_NUMBER_BY_ABBR[month_abbr.lower()]

    return {
        "file": file_path,
        "family": family,
        "month": month_number,
        "month_abbr": month_abbr,
    }


def find_latest_unique_forecast_workbooks(
    source_folder: str | Path,
    fiscal_year: str,
    forecast_month: int,
) -> list[dict]:
    """
    Find the latest available workbook for each forecast family
    before the selected target month.
    """

    source_path = Path(source_folder)

    if not source_path.exists():
        raise FileNotFoundError(f"Forecast source folder not found: {source_path}")

    if forecast_month not in MONTH_ABBR:
        raise ValueError("Forecast month must be between 1 and 12.")

    if forecast_month == 7:
        raise ValueError(
            "July is the initial fiscal-year forecast month. "
            "Use Initial Forecast Month Setup to create July."
        )

    target_fiscal_position = FISCAL_MONTH_ORDER[forecast_month]

    workbook_families: dict[str, list[dict]] = {}

    for file_path in source_path.glob("*.xls*"):

        parsed = parse_forecast_filename(
            file_path=file_path,
            fiscal_year=fiscal_year,
        )

        if parsed is None:
            continue

        file_fiscal_position = FISCAL_MONTH_ORDER[parsed["month"]]

        # Do not use the target month or a later fiscal month
        # as the source.
        if file_fiscal_position >= target_fiscal_position:
            continue

        workbook_families.setdefault(
            parsed["family"],
            [],
        ).append(parsed)

    if not workbook_families:
        raise FileNotFoundError(
            f"No valid {fiscal_year} forecast workbooks "
            f"were found before "
            f"{MONTH_ABBR[forecast_month]} in "
            f"{source_path}"
        )

    latest_workbooks: list[dict] = []

    expected_previous_month = get_previous_month(forecast_month)

    for family, candidates in workbook_families.items():

        latest = max(
            candidates,
            key=lambda item: FISCAL_MONTH_ORDER[item["month"]],
        )

        latest["is_expected_previous_month"] = (
            latest["month"] == expected_previous_month
        )

        latest_workbooks.append(latest)

    latest_workbooks.sort(key=lambda item: item["family"].lower())

    return latest_workbooks


# ============================================================
# EXCEL HEADER HELPERS
# ============================================================


def find_header(
    worksheet,
    names: set[str],
) -> tuple[int | None, int | None]:

    for row_index in range(
        1,
        worksheet.max_row + 1,
    ):
        for column_index in range(
            1,
            worksheet.max_column + 1,
        ):
            value = normalize_text(
                worksheet.cell(
                    row=row_index,
                    column=column_index,
                ).value
            )

            if value in names:
                return row_index, column_index

    return None, None


# ============================================================
# FORECAST TYPE
# ============================================================


def update_forecast_workbook(
    workbook_path: str | Path,
    target_fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
) -> dict:
    """
    Prepare an existing forecast workbook for the selected
    monthly forecast cycle.

    At this stage we intentionally update only Forecast Type.

    Monthly values, formulas, formatting and workbook structure
    are preserved.
    """

    workbook_path = Path(workbook_path)

    workbook = load_workbook(
        workbook_path,
        data_only=False,
    )

    forecast_type_value = (
        f"{target_fiscal_year} Forecast_" f"{forecast_month:02d}" f"{forecast_year}"
    )

    worksheets_updated = 0
    forecast_values_updated = 0
    skipped_sheets: list[str] = []

    try:

        for worksheet in workbook.worksheets:

            header_row, forecast_column = find_header(
                worksheet,
                {
                    "forecast type",
                },
            )

            if header_row is None or forecast_column is None:
                skipped_sheets.append(f"{worksheet.title}: " "Forecast Type not found")
                continue

            rows_updated = 0

            for row_index in range(
                header_row + 1,
                worksheet.max_row + 1,
            ):
                cell = worksheet.cell(
                    row=row_index,
                    column=forecast_column,
                )

                if is_formula(cell.value):
                    continue

                current_value = normalize_text(cell.value)

                if not current_value:
                    continue

                if "forecast" in current_value:
                    cell.value = forecast_type_value
                    rows_updated += 1

            if rows_updated > 0:
                worksheets_updated += 1
                forecast_values_updated += rows_updated

        workbook.save(workbook_path)

    finally:
        workbook.close()

    return {
        "worksheets_updated": worksheets_updated,
        "forecast_values_updated": forecast_values_updated,
        "forecast_type": forecast_type_value,
        "skipped_sheets": skipped_sheets,
    }


# ============================================================
# FORMULA AUDIT
# ============================================================


def get_formula_cells(
    workbook_path: str | Path,
) -> dict[str, dict[str, str]]:
    """
    Capture all formulas in a workbook without modifying it.
    """

    workbook_path = Path(workbook_path)

    workbook = load_workbook(
        workbook_path,
        data_only=False,
    )

    formulas: dict[str, dict[str, str]] = {}

    try:

        for worksheet in workbook.worksheets:

            sheet_formulas: dict[str, str] = {}

            for row in worksheet.iter_rows():

                for cell in row:

                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        sheet_formulas[cell.coordinate] = cell.value

            formulas[worksheet.title] = sheet_formulas

    finally:
        workbook.close()

    return formulas


def count_formulas(
    formula_map: dict[str, dict[str, str]],
) -> int:

    return sum(len(sheet_formulas) for sheet_formulas in formula_map.values())


def compare_formula_preservation(
    source_formulas: dict[str, dict[str, str]],
    destination_formulas: dict[str, dict[str, str]],
) -> dict:

    missing_formulas = []
    changed_formulas = []

    for sheet_name, formulas in source_formulas.items():

        destination_sheet = destination_formulas.get(
            sheet_name,
            {},
        )

        for coordinate, source_formula in formulas.items():

            destination_formula = destination_sheet.get(coordinate)

            if destination_formula is None:

                missing_formulas.append(
                    {
                        "sheet": sheet_name,
                        "cell": coordinate,
                        "formula": source_formula,
                    }
                )

                continue

            if destination_formula != source_formula:

                changed_formulas.append(
                    {
                        "sheet": sheet_name,
                        "cell": coordinate,
                        "before": source_formula,
                        "after": destination_formula,
                    }
                )

    return {
        "source_formula_count": count_formulas(source_formulas),
        "destination_formula_count": count_formulas(destination_formulas),
        "missing_formula_count": len(missing_formulas),
        "changed_formula_count": len(changed_formulas),
        "missing_formulas": missing_formulas,
        "changed_formulas": changed_formulas,
        "formulas_preserved": not missing_formulas and not changed_formulas,
    }


# ============================================================
# CREATE MONTHLY FORECAST WORKBOOKS
# ============================================================


class ForecastFilesExistError(PlainLanguageError, FileExistsError):
    """Existing Files = cancel and some of the month's files already exist."""


def create_forecast_refresh(
    source_folder: str | Path,
    output_folder: str | Path,
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    overwrite_mode: str = "skip",
) -> dict:
    """
    Create the next month's forecast workbooks.

    Process:
        1. Find latest workbook for each family.
        2. Copy it to the output folder.
        3. Rename it for the selected month.
        4. Update Forecast Type.
        5. Verify existing formulas were preserved.
    """

    source_path = Path(source_folder)
    output_path = Path(output_folder)

    if not source_path.exists():
        raise FileNotFoundError(f"Forecast source folder not found: " f"{source_path}")

    if forecast_month not in MONTH_ABBR:
        raise ValueError("Forecast month must be between 1 and 12.")

    if overwrite_mode not in {
        "skip",
        "replace",
        "cancel",
    }:
        raise ValueError("overwrite_mode must be " "'skip', 'replace', or 'cancel'.")

    output_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    previous_month = get_previous_month(forecast_month)

    expected_source_month_abbr = MONTH_ABBR[previous_month]

    target_month_abbr = MONTH_ABBR[forecast_month]

    selected_workbooks = find_latest_unique_forecast_workbooks(
        source_folder=source_path,
        fiscal_year=fiscal_year,
        forecast_month=forecast_month,
    )

    # "cancel" means: change nothing if any of this month's files already
    # exist. Check them all before copying the first one.
    if overwrite_mode == "cancel":
        existing = [
            output_path / f"{info['family']}_{target_month_abbr} Forecast{info['file'].suffix}"
            for info in selected_workbooks
        ]
        existing = [path.name for path in existing if path.exists()]
        if existing:
            shown = "\n".join(f"  - {name}" for name in existing[:10])
            more = f"\n  ...and {len(existing) - 10} more" if len(existing) > 10 else ""
            raise ForecastFilesExistError(
                f"{len(existing)} of this month's forecast files already exist in "
                f"the output folder, so nothing was created (Existing Files = "
                f'"cancel"):\n\n{shown}{more}\n\nChoose "skip" to keep them or '
                '"replace" to recreate them.'
            )

    created: list[Path] = []
    skipped: list[Path] = []
    failed: list[dict[str, str]] = []

    formula_audits: list[dict] = []
    sources_used: list[dict] = []
    older_sources: list[dict] = []

    worksheets_updated = 0
    forecast_values_updated = 0

    for workbook_info in selected_workbooks:

        source_file: Path = workbook_info["file"]
        family = workbook_info["family"]
        source_month_abbr = workbook_info["month_abbr"]

        new_filename = (
            f"{family}_" f"{target_month_abbr} Forecast" f"{source_file.suffix}"
        )

        destination = output_path / new_filename

        try:

            if destination.exists():

                if overwrite_mode == "skip":

                    skipped.append(destination)

                    sources_used.append(
                        {
                            "family": family,
                            "source_file": source_file,
                            "source_month": source_month_abbr,
                            "destination": destination,
                            "status": "skipped",
                        }
                    )

                    continue

                if overwrite_mode == "cancel":
                    raise RuntimeError(
                        "Forecast file already exists: " f"{destination.name}"
                    )

                if overwrite_mode == "replace":
                    destination.unlink()

            # ----------------------------------------------
            # Capture formulas before copying.
            # ----------------------------------------------

            source_formulas = get_formula_cells(source_file)

            # ----------------------------------------------
            # Create new monthly workbook.
            # ----------------------------------------------

            shutil.copy2(
                source_file,
                destination,
            )

            # ----------------------------------------------
            # Update Forecast Type.
            # ----------------------------------------------

            result = update_forecast_workbook(
                workbook_path=destination,
                target_fiscal_year=fiscal_year,
                forecast_month=forecast_month,
                forecast_year=forecast_year,
            )

            # ----------------------------------------------
            # Confirm formulas survived the creation step.
            # ----------------------------------------------

            destination_formulas = get_formula_cells(destination)

            formula_audit = compare_formula_preservation(
                source_formulas=source_formulas,
                destination_formulas=destination_formulas,
            )

            formula_audits.append(
                {
                    "file": destination.name,
                    "source_file": source_file.name,
                    "source_month": source_month_abbr,
                    "target_month": target_month_abbr,
                    **formula_audit,
                }
            )

            worksheets_updated += result["worksheets_updated"]

            forecast_values_updated += result["forecast_values_updated"]

            created.append(destination)

            sources_used.append(
                {
                    "family": family,
                    "source_file": source_file,
                    "source_month": source_month_abbr,
                    "destination": destination,
                    "status": "created",
                }
            )

            if not workbook_info["is_expected_previous_month"]:
                older_sources.append(
                    {
                        "family": family,
                        "source_file": source_file.name,
                        "source_month": source_month_abbr,
                        "expected_month": expected_source_month_abbr,
                    }
                )

        except Exception as exc:

            failed.append(
                {
                    "file": source_file.name,
                    "path": str(source_file),
                    "error": explain(exc),
                }
            )

    return {
        "files_created": len(created),
        "files_skipped": len(skipped),
        "files_failed": len(failed),
        "unique_workbooks_found": len(selected_workbooks),
        "worksheets_updated": worksheets_updated,
        "forecast_values_updated": forecast_values_updated,
        "expected_source_month": expected_source_month_abbr,
        "target_month": target_month_abbr,
        "output_folder": str(output_path),
        "created_files": created,
        "skipped_files": skipped,
        "failures": failed,
        "sources_used": sources_used,
        "older_sources": older_sources,
        "formula_audits": formula_audits,
    }


# ============================================================
# STAGE 4 - POPULATE / FINALIZE FORECAST FROM TABLEAU
# ============================================================

TABLEAU_ALIASES = {
    "business_code": {"gl business code", "bus code name", "business code", "bus code"},
    "company": {"company", "company name", "entity"},
    "fs_category": {"gl fs category", "fs category", "financial statement category"},
    "account_mapping": {"gl account mapping", "account mapping", "gl mapping"},
    "sub_segment": {
        "gl sub segment code",
        "gl sub segment",
        "sub segment code",
        "sub segment",
        "account",
        "account code",
        "gl account",
    },
    "description": {
        "gl sub account description",
        "gl sub account desc",
        "sub account description",
        "sub account desc",
        "account description",
        "description",
    },
    "vendor": {"vendor name", "vendor"},
}

EXCEL_ALIASES = TABLEAU_ALIASES


def _s4norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _s4number(value: Any):
    if value is None or isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    if not text:
        return None
    n = text.replace(",", "").replace("$", "")
    if n.startswith("(") and n.endswith(")"):
        n = "-" + n[1:-1]
    try:
        x = float(n)
        return int(x) if x.is_integer() else x
    except ValueError:
        return text


def _s4_tableau_value(row: dict, field: str):
    aliases = TABLEAU_ALIASES[field]
    for header, value in row.items():
        if _s4norm(header) in aliases:
            return value
    return None


def _s4_detail_rows(rows: list[dict]) -> list[dict]:
    result = []
    for row in rows:
        business = _s4norm(_s4_tableau_value(row, "business_code"))
        segment = _s4norm(_s4_tableau_value(row, "sub_segment"))
        desc = _s4norm(_s4_tableau_value(row, "description"))
        if business == "grand total" or desc == "grand total":
            continue
        if segment == "total" and desc == "total":
            continue
        if any(v not in (None, "") for v in row.values()):
            result.append(row)
    return result


def _s4_month(value: Any) -> int | None:
    if isinstance(value, datetime):
        return value.month
    if hasattr(value, "month") and isinstance(getattr(value, "month"), int):
        return int(value.month)
    text = str(value or "").strip()
    if not text or text.startswith("="):
        return None
    low = text.lower()
    for month in range(1, 13):
        if re.search(rf"\b{calendar.month_name[month].lower()}\b", low):
            return month
        if re.search(rf"\b{calendar.month_abbr[month].lower()}\b", low):
            return month
    return None


def _s4_tableau_months(rows: list[dict]) -> dict[int, Any]:
    if not rows:
        raise RuntimeError("No Tableau detail rows supplied.")
    lookup = {}
    for header in rows[0].keys():
        month = _s4_month(header)
        if month is None:
            continue
        if month in lookup:
            raise RuntimeError(f"Duplicate Tableau month: {MONTH_ABBR[month]}")
        lookup[month] = header
    if not lookup:
        raise RuntimeError("No Tableau fiscal month columns were found.")
    missing = [MONTH_ABBR[m] for m in FISCAL_MONTH_ORDER if m not in lookup]
    if missing:
        print(
            "[PARTIAL MONTHS] Missing from Tableau: "
            + ", ".join(missing)
            + ". Excel values for those months will be blank."
        )
    return lookup


def _s4_structure(ws) -> dict:
    header_row, forecast_col = find_header(ws, {"forecast type"})
    if header_row is None or forecast_col is None:
        raise RuntimeError(f'Forecast Type header not found in "{ws.title}".')

    headers = {}
    for col in range(1, ws.max_column + 1):
        key = _s4norm(ws.cell(header_row, col).value)
        if key:
            headers[key] = col

    month_cols = {}
    for col in range(1, ws.max_column + 1):
        month = _s4_month(ws.cell(header_row, col).value)
        if month is not None and month not in month_cols:
            month_cols[month] = col

    # Old workbooks may have Jul as a date/value and Aug-Jun as =EDATE formulas.
    if len(month_cols) < 12:
        july_col = month_cols.get(7)
        if july_col is None:
            for col in range(1, ws.max_column + 1):
                if _s4norm(ws.cell(header_row, col).value).startswith("jul"):
                    july_col = col
                    break
        if july_col is not None and july_col + 11 <= ws.max_column:
            month_cols = {
                month: july_col + offset
                for offset, month in enumerate(FISCAL_MONTH_ORDER)
            }

    missing = [MONTH_ABBR[m] for m in FISCAL_MONTH_ORDER if m not in month_cols]
    if missing:
        raise RuntimeError(
            f'"{ws.title}" does not contain a complete Jul-Jun block. '
            f'Missing: {", ".join(missing)}'
        )

    desc_cols = {}
    for field, aliases in EXCEL_ALIASES.items():
        for header, col in headers.items():
            if header in aliases:
                desc_cols[field] = col
                break

    total_col = None
    for header, col in headers.items():
        if header in {"grand total", "total"}:
            total_col = col
            break
    if total_col is None:
        raise RuntimeError(f'Grand Total column not found in "{ws.title}".')

    return {
        "header_row": header_row,
        "forecast_type_column": forecast_col,
        "descriptive_columns": desc_cols,
        "month_columns": month_cols,
        "grand_total_column": total_col,
    }


def _s4_forecast_type(ws, structure):
    col = structure["forecast_type_column"]
    for row in range(structure["header_row"] + 1, ws.max_row + 1):
        value = ws.cell(row, col).value
        if value not in (None, ""):
            return value
    raise RuntimeError(f'Forecast Type value not found in "{ws.title}".')


def _s4_current_detail(ws, structure):
    first = structure["header_row"] + 1
    col = structure["forecast_type_column"]
    last = first - 1
    for row in range(first, ws.max_row + 1):
        if ws.cell(row, col).value not in (None, ""):
            last = row
    return first, last


def _s4_copy_style(source, target):
    target.font = copy(source.font)
    target.fill = copy(source.fill)
    target.border = copy(source.border)
    target.alignment = copy(source.alignment)
    target.number_format = source.number_format
    target.protection = copy(source.protection)


def _s4_write_real_dates(ws, structure, fiscal_year: str):
    match = re.search(r"(\d{2,4})", str(fiscal_year))
    if not match:
        raise ValueError(f"Invalid fiscal year: {fiscal_year}")
    fy = int(match.group(1))
    if fy < 100:
        fy += 2000
    start_year = fy - 1

    expected = []
    for month in FISCAL_MONTH_ORDER:
        year = start_year if month >= 7 else fy
        value = datetime(year, month, 1)
        cell = ws.cell(structure["header_row"], structure["month_columns"][month])
        cell.value = value  # REAL Excel date; never =EDATE(...)
        cell.number_format = "mmm-yy"
        expected.append(value)
    return expected


def _s4_total_border(ws, row: int, first_col: int, last_col: int):
    top = Side(style="medium")
    bottom = Side(style="double")
    for col in range(first_col, last_col + 1):
        cell = ws.cell(row, col)
        old = cell.border
        cell.border = Border(
            left=old.left,
            right=old.right,
            top=top,
            bottom=bottom,
            diagonal=old.diagonal,
            diagonal_direction=old.diagonal_direction,
            diagonalUp=old.diagonalUp,
            diagonalDown=old.diagonalDown,
            outline=old.outline,
            vertical=old.vertical,
            horizontal=old.horizontal,
        )


def _s4_finalize(ws, structure, fiscal_year: str, last_detail_row: int) -> int:
    _s4_write_real_dates(ws, structure, fiscal_year)

    first = structure["header_row"] + 1
    total_row = last_detail_row + 1
    month_cols = structure["month_columns"]
    gt_col = structure["grand_total_column"]

    # No literal "Grand Total" label.
    for col in {
        structure["forecast_type_column"],
        *structure["descriptive_columns"].values(),
    }:
        ws.cell(total_row, col).value = None

    # Detail Grand Total formulas.
    for row in range(first, last_detail_row + 1):
        refs = [f"{get_column_letter(month_cols[m])}{row}" for m in FISCAL_MONTH_ORDER]
        ws.cell(row, gt_col).value = "=SUM(" + ",".join(refs) + ")"

    # Bottom monthly totals.
    for month in FISCAL_MONTH_ORDER:
        col = month_cols[month]
        letter = get_column_letter(col)
        ws.cell(total_row, col).value = (
            f"=SUM({letter}{first}:{letter}{last_detail_row})"
        )

    # Bottom overall total.
    gt_letter = get_column_letter(gt_col)
    ws.cell(total_row, gt_col).value = (
        f"=SUM({gt_letter}{first}:{gt_letter}{last_detail_row})"
    )

    all_cols = [
        structure["forecast_type_column"],
        gt_col,
        *structure["descriptive_columns"].values(),
        *month_cols.values(),
    ]
    _s4_total_border(ws, total_row, min(all_cols), max(all_cols))
    return total_row


def read_tableau_audit_sheet(
    audit_workbook_path: str | Path,
    tableau_worksheet_name: str,
) -> list[dict]:

    wb = load_workbook(
        audit_workbook_path,
        read_only=True,
        data_only=False,
    )

    try:
        # ----------------------------------------------------
        # 1. Exact match
        # ----------------------------------------------------
        if tableau_worksheet_name in wb.sheetnames:
            ws = wb[tableau_worksheet_name]

        else:
            wanted = _s4norm(tableau_worksheet_name)

            # ------------------------------------------------
            # 2. Normalized exact match
            # ------------------------------------------------
            matches = [s for s in wb.worksheets if _s4norm(s.title) == wanted]

            if len(matches) == 1:
                ws = matches[0]

            else:
                # --------------------------------------------
                # 3. Excel audit tabs are limited to 31 chars.
                #    Stage 3 therefore truncates long Tableau
                #    worksheet names when writing the audit.
                #
                #    Match the beginning of the original
                #    Tableau worksheet name to the stored
                #    audit-sheet title.
                # --------------------------------------------
                prefix_matches = []

                for sheet in wb.worksheets:
                    sheet_norm = _s4norm(sheet.title)

                    if wanted.startswith(sheet_norm) or sheet_norm.startswith(wanted):
                        prefix_matches.append(sheet)

                if len(prefix_matches) == 1:
                    ws = prefix_matches[0]

                else:
                    # ----------------------------------------
                    # 4. Handle duplicate truncated names.
                    #
                    #    Excel audit creation may suffix
                    #    collisions with _2, _3, etc.
                    #    Remove that suffix and compare again.
                    # ----------------------------------------
                    collision_matches = []

                    for sheet in wb.worksheets:
                        base_title = re.sub(
                            r"_\d+$",
                            "",
                            sheet.title,
                        )

                        base_norm = _s4norm(base_title)

                        if wanted.startswith(base_norm):
                            collision_matches.append(sheet)

                    if len(collision_matches) == 1:
                        ws = collision_matches[0]

                    elif len(collision_matches) > 1:
                        raise RuntimeError(
                            f'Audit sheet "{tableau_worksheet_name}" '
                            f"matches multiple truncated audit sheets: "
                            f"{[s.title for s in collision_matches]}"
                        )

                    else:
                        raise RuntimeError(
                            f'Audit sheet "{tableau_worksheet_name}" '
                            f"not found. Available: {wb.sheetnames}"
                        )

        values = list(ws.iter_rows(values_only=True))

        if not values:
            raise RuntimeError(f'Audit sheet "{ws.title}" is empty.')

        headers = list(values[0])
        rows = []

        for values_row in values[1:]:
            row = {
                header: value
                for header, value in zip(
                    headers,
                    values_row,
                )
                if header not in (None, "")
            }

            if any(value not in (None, "") for value in row.values()):
                rows.append(row)

        return rows

    finally:
        wb.close()


def populate_forecast_sheet_from_tableau_rows(
    workbook_path: str | Path,
    excel_sheet_name: str,
    tableau_rows: list[dict],
    fiscal_year: str,
) -> dict:
    workbook_path = Path(workbook_path)
    wb = load_workbook(workbook_path, data_only=False)

    try:
        if excel_sheet_name in wb.sheetnames:
            ws = wb[excel_sheet_name]
        else:
            wanted = re.sub(r"[^a-z0-9]+", "", excel_sheet_name.lower())
            matches = [
                s
                for s in wb.worksheets
                if re.sub(r"[^a-z0-9]+", "", s.title.lower()) == wanted
            ]
            if len(matches) != 1:
                raise RuntimeError(
                    f'Excel sheet "{excel_sheet_name}" not found. '
                    f"Available: {wb.sheetnames}"
                )
            ws = matches[0]

        structure = _s4_structure(ws)
        forecast_type = _s4_forecast_type(ws, structure)
        detail_rows = _s4_detail_rows(tableau_rows)
        if not detail_rows:
            raise RuntimeError("No usable Tableau detail rows supplied.")
        tableau_months = _s4_tableau_months(detail_rows)

        first, old_last = _s4_current_detail(ws, structure)
        old_count = max(0, old_last - first + 1)
        new_count = len(detail_rows)

        style_row = first if old_count else structure["header_row"]
        styles = {col: ws.cell(style_row, col) for col in range(1, ws.max_column + 1)}
        row_height = ws.row_dimensions[style_row].height

        difference = new_count - old_count
        if difference > 0:
            insert_at = old_last + 1 if old_count else first
            ws.insert_rows(insert_at, amount=difference)
        elif difference < 0:
            ws.delete_rows(first + new_count, amount=abs(difference))

        last = first + new_count - 1

        for row in range(first, last + 1):
            for col, source in styles.items():
                _s4_copy_style(source, ws.cell(row, col))
            if row_height is not None:
                ws.row_dimensions[row].height = row_height

        revenue = "revenue" in workbook_path.name.lower()

        for offset, tableau_row in enumerate(detail_rows):
            excel_row = first + offset
            ws.cell(excel_row, structure["forecast_type_column"]).value = forecast_type

            for field, col in structure["descriptive_columns"].items():
                ws.cell(excel_row, col).value = _s4number(
                    _s4_tableau_value(tableau_row, field)
                )

            for month in FISCAL_MONTH_ORDER:
                col = structure["month_columns"][month]
                header = tableau_months.get(month)
                value = None if header is None else _s4number(tableau_row.get(header))
                if (
                    revenue
                    and isinstance(value, (int, float))
                    and not isinstance(value, bool)
                ):
                    value *= -1
                ws.cell(excel_row, col).value = value

        total_row = _s4_finalize(ws, structure, fiscal_year, last)
        wb.save(workbook_path)

        return {
            "workbook": str(workbook_path),
            "worksheet": ws.title,
            "tableau_detail_rows": new_count,
            "excel_detail_rows": new_count,
            "old_excel_detail_rows": old_count,
            "forecast_type": forecast_type,
            "revenue_sign_rule": revenue,
            "total_row": total_row,
        }
    finally:
        wb.close()


def populate_forecast_sheet_from_audit(
    workbook_path: str | Path,
    excel_sheet_name: str,
    audit_workbook_path: str | Path,
    tableau_worksheet_name: str,
    fiscal_year: str,
) -> dict:
    rows = read_tableau_audit_sheet(audit_workbook_path, tableau_worksheet_name)
    return populate_forecast_sheet_from_tableau_rows(
        workbook_path, excel_sheet_name, rows, fiscal_year
    )


# ============================================================
# STEP 2 - UPDATE THE CLOSED MONTHS, KEEP THE FORECAST MONTHS
# ============================================================
#
# The month's workbook was copied from last month's (step 1), so it already
# holds each department's forecast. Tableau only has actuals for the months
# that have closed (for October: Jul, Aug, Sep). Step 2 therefore:
#
#   * writes the closed months from Tableau (the just-closed month's actual
#     replaces its forecast),
#   * leaves the forecast months (Oct-Jun) exactly as they are,
#   * matches lines on the account columns, adds lines that are new in
#     Tableau, and keeps lines Tableau no longer has (their closed months
#     are cleared: no actuals were posted).


class ActualsUpdateError(PlainLanguageError, RuntimeError):
    """The sheet could not be updated safely; it was left unchanged."""


_MONTH_FORMATS = (
    "%b-%y", "%B %Y", "%b %Y", "%B-%Y", "%b-%Y", "%b %y",
    "%Y-%m", "%Y-%m-%d", "%m/%d/%Y", "%m/%Y",
)


def _month_year(value: Any) -> tuple[int | None, int] | None:
    """(year, month) of a month column header; year None if it has none."""
    if hasattr(value, "month") and hasattr(value, "year"):
        return int(value.year), int(value.month)
    text = " ".join(str(value or "").split())
    if not text or text.startswith("="):
        return None
    for fmt in _MONTH_FORMATS:
        try:
            parsed = datetime.strptime(text, fmt)
        except ValueError:
            continue
        return parsed.year, parsed.month
    month = _s4_month(text)
    return (None, month) if month else None


def closed_months(fiscal_year: str, forecast_month: int) -> list[tuple[int, int]]:
    """(year, month) of the months before the forecast month, July first.

    FY27 / October -> [(2026, 7), (2026, 8), (2026, 9)].
    """
    match = re.search(r"(\d{2,4})", str(fiscal_year))
    if not match:
        raise ValueError(f"Invalid fiscal year: {fiscal_year}")
    end = int(match.group(1))
    end = end + 2000 if end < 100 else end
    months = sorted(FISCAL_MONTH_ORDER, key=FISCAL_MONTH_ORDER.get)
    position = FISCAL_MONTH_ORDER[forecast_month]
    return [(end - 1 if m >= 7 else end, m) for m in months if FISCAL_MONTH_ORDER[m] < position]


def _key_part(value: Any) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return " ".join(str(value if value is not None else "").strip().lower().split())


def _find_sheet(wb, name: str):
    if name in wb.sheetnames:
        return wb[name]
    wanted = re.sub(r"[^a-z0-9]+", "", name.lower())
    matches = [s for s in wb.worksheets if re.sub(r"[^a-z0-9]+", "", s.title.lower()) == wanted]
    if len(matches) != 1:
        raise RuntimeError(f'Excel sheet "{name}" not found. Available: {wb.sheetnames}')
    return matches[0]


def update_actual_months(
    workbook_path: str | Path,
    excel_sheet_name: str,
    tableau_rows: list[dict],
    fiscal_year: str,
    forecast_month: int,
    expected_forecast_type: str,
) -> dict:
    """Write the closed months from Tableau into one forecast sheet.

    Forecast months are never touched. Raises ActualsUpdateError (sheet left
    unchanged) when the sheet is for another month or lines can't be matched.
    """
    workbook_path = Path(workbook_path)
    wb = load_workbook(workbook_path, data_only=False)
    try:
        ws = _find_sheet(wb, excel_sheet_name)
        label = f"{workbook_path.name} / {ws.title}"
        structure = _s4_structure(ws)

        sheet_type = str(_s4_forecast_type(ws, structure)).strip()
        if _s4norm(sheet_type) != _s4norm(expected_forecast_type):
            raise ActualsUpdateError(
                f"{label} is for {sheet_type}, not {expected_forecast_type}. It was "
                "not changed. If step 1 kept it because it already existed, run "
                'step 1 again with Existing Files = "replace".'
            )

        fields = list(structure["descriptive_columns"])
        if not fields:
            raise ActualsUpdateError(
                f"{label} has no account columns (GL Account, Sub Segment, "
                "Description...) to match lines on. It was not changed."
            )

        # Tableau column for each closed month (by month and year).
        wanted = closed_months(fiscal_year, forecast_month)
        tableau_detail = _s4_detail_rows(tableau_rows)
        columns = {}
        for header in (tableau_rows[0].keys() if tableau_rows else []):
            parsed = _month_year(header)
            if parsed is None:
                continue
            year, month = parsed
            for want_year, want_month in wanted:
                if month == want_month and year in (None, want_year):
                    columns[(want_year, want_month)] = header
        missing = [f"{calendar.month_abbr[m]} {y}" for y, m in wanted if (y, m) not in columns]

        revenue = "revenue" in workbook_path.name.lower()

        def tableau_value(row, ym):
            header = columns.get(ym)
            value = None if header is None else _s4number(row.get(header))
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                return 0
            return -value if revenue else value

        # Group Tableau lines by account key; split lines (e.g. by vendor when
        # the sheet has no vendor column) are added together.
        tableau_groups: dict[tuple, list[dict]] = {}
        for row in tableau_detail:
            key = tuple(_key_part(_s4_tableau_value(row, field)) for field in fields)
            tableau_groups.setdefault(key, []).append(row)

        first, last = _s4_current_detail(ws, structure)
        excel_rows: dict[tuple, list[int]] = {}
        for r in range(first, last + 1):
            if ws.cell(r, structure["forecast_type_column"]).value in (None, ""):
                continue
            key = tuple(
                _key_part(ws.cell(r, structure["descriptive_columns"][field]).value)
                for field in fields
            )
            excel_rows.setdefault(key, []).append(r)

        month_cols = structure["month_columns"]
        updates: dict[int, dict[int, Any]] = {}
        for key, rows in excel_rows.items():
            group = tableau_groups.get(key, [])
            if len(rows) == 1:
                updates[rows[0]] = {
                    m: sum(tableau_value(t, (y, m)) for t in group) if group else None
                    for y, m in wanted
                }
            elif len(rows) == len(group):
                for r, t in zip(rows, group):
                    updates[r] = {m: tableau_value(t, (y, m)) for y, m in wanted}
            elif not group:
                for r in rows:
                    updates[r] = {m: None for _, m in wanted}
            else:
                shown = " / ".join(part for part in key if part)
                raise ActualsUpdateError(
                    f"{label}: the line \"{shown}\" appears {len(rows)} times in Excel "
                    f"and {len(group)} times in Tableau, so the actuals can't be "
                    "matched. It was not changed."
                )

        new_keys = [key for key in tableau_groups if key not in excel_rows]

        # --- write (nothing above changed the sheet) ---------------------
        for r, values in updates.items():
            for m, value in values.items():
                ws.cell(r, month_cols[m]).value = value

        if new_keys:
            style_row = last if last >= first else structure["header_row"]
            styles = {c: ws.cell(style_row, c) for c in range(1, ws.max_column + 1)}
            ws.insert_rows(last + 1, amount=len(new_keys))
            for offset, key in enumerate(new_keys):
                r = last + 1 + offset
                for c, source in styles.items():
                    _s4_copy_style(source, ws.cell(r, c))
                ws.cell(r, structure["forecast_type_column"]).value = sheet_type
                sample = tableau_groups[key][0]
                for field, c in structure["descriptive_columns"].items():
                    ws.cell(r, c).value = _s4number(_s4_tableau_value(sample, field))
                for y, m in wanted:
                    ws.cell(r, month_cols[m]).value = sum(
                        tableau_value(t, (y, m)) for t in tableau_groups[key]
                    )
                for m, c in month_cols.items():
                    if (FISCAL_MONTH_ORDER[m] >= FISCAL_MONTH_ORDER[forecast_month]):
                        ws.cell(r, c).value = None
            last += len(new_keys)

        _s4_finalize(ws, structure, fiscal_year, last)
        wb.save(workbook_path)

        cleared = sum(1 for key, rows in excel_rows.items() if key not in tableau_groups for _ in rows)
        return {
            "workbook": str(workbook_path),
            "worksheet": ws.title,
            "months_updated": [f"{calendar.month_abbr[m]} {y}" for y, m in wanted],
            "months_missing_in_tableau": missing,
            "lines_updated": len(updates) - cleared,
            "lines_added": len(new_keys),
            "lines_without_actuals": cleared,
        }
    finally:
        wb.close()


def update_actual_months_from_audit(
    workbook_path: str | Path,
    excel_sheet_name: str,
    audit_workbook_path: str | Path,
    tableau_worksheet_name: str,
    fiscal_year: str,
    forecast_month: int,
    expected_forecast_type: str,
) -> dict:
    rows = read_tableau_audit_sheet(audit_workbook_path, tableau_worksheet_name)
    return update_actual_months(
        workbook_path, excel_sheet_name, rows, fiscal_year, forecast_month, expected_forecast_type
    )

