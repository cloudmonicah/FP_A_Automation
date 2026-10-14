from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

# ============================================================
# DIRECT COST FINANCIAL STATEMENT READER
# ============================================================
#
# PURPOSE
# -------
# Read finalized ALIV and CBL Financial Statements and extract
# Direct Cost / Cost of Goods Sold Actual values.
#
# BUSINESS RULES
# --------------
# Column B = Financial Statement line
# Column C = Actual
#
# Financial Statement values are already expressed in $000.
#
# IMPORTANT
# ---------
# Row numbers are NOT hard-coded.
#
# Required lines are located dynamically by their labels.
# This protects the automation if rows are inserted or moved.
# ============================================================


# ============================================================
# REQUIRED ALIV DIRECT COST LINES
# ============================================================

ALIV_DIRECT_COST_LINES = {
    "INTERCONNECT": "Interconnect",
    "RESELLER COMMISSIONS": "Reseller Commissions",
    "ROAMING SERVICE COSTS": "Roaming Service Costs",
    "SERVICE FEES": "Service Fees",
    "DIRECT COST": "Direct Cost",
}


# ============================================================
# REQUIRED CBL DIRECT COST LINES
# ============================================================

CBL_DIRECT_COST_LINES = {
    "PROGRAMMING FEES": "Programming fees",
    "HARDWARE/PBX": "Hardware/PBX",
    "VOICE TERMINATION COSTS": "Voice termination costs",
    "INTERNET BANDWIDTH": "Internet Bandwidth",
    "CLOUD SERVICES COS": "Cloud services COS",
    "DTH COST OF SALES": "DTH cost of sales",
    "SMART HOME PRODUCTS COS": "Smart Home Products COS",
    "COST OF GOODS SOLD": "Cost of Goods Sold",
}


# ============================================================
# LABEL NORMALIZATION
# ============================================================


def _canonical_statement_line(
    value: Any,
) -> str:
    """
    Normalize Financial Statement labels.

    Examples:

        "20 General and Administration"
            -> "GENERAL AND ADMINISTRATION"

        "Hardware/PBX"
            -> "HARDWARE/PBX"
    """

    if value is None:
        return ""

    text = str(value).strip()

    # --------------------------------------------------------
    # Remove a leading numeric Financial Statement code.
    #
    # Example:
    #     "20 General and Administration"
    # becomes:
    #     "General and Administration"
    # --------------------------------------------------------

    parts = text.split(
        " ",
        1,
    )

    if len(parts) == 2 and parts[0].isdigit():

        text = parts[1].strip()

    # --------------------------------------------------------
    # Standardize "&" versus "and".
    # --------------------------------------------------------

    text = text.replace(
        "&",
        "and",
    )

    # --------------------------------------------------------
    # Normalize case and whitespace.
    # --------------------------------------------------------

    return " ".join(text.upper().split())


# ============================================================
# GENERIC DIRECT COST SECTION READER
# ============================================================


def _read_direct_cost_section_by_label(
    *,
    workbook_path: str | Path,
    sheet_name: str,
    required_lines: dict[str, str],
    section_total_label: str,
    closed_period=None,
) -> dict[str, float]:
    """
    Read one Direct Cost section dynamically.

    The Financial Statement layout is:

        Column B = label
        Column C = Actual ($000)

    The function finds all occurrences of required labels and
    selects the occurrence belonging to the Direct Cost section.
    """

    workbook_path = Path(workbook_path)

    # ========================================================
    # SOURCE FILE CHECK
    # ========================================================

    if not workbook_path.exists():

        raise FileNotFoundError(
            "Direct Cost Financial Statement not found:\n" f"{workbook_path}"
        )

    # ========================================================
    # OPEN WORKBOOK
    # ========================================================

    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
    )

    try:

        # ====================================================
        # WORKSHEET CHECK
        # ====================================================

        if sheet_name not in workbook.sheetnames:

            raise ValueError(
                "Expected Direct Cost worksheet "
                f"'{sheet_name}' was not found in:\n"
                f"{workbook_path}"
            )

        sheet = workbook[sheet_name]

        # ====================================================
        # AUDIT INFORMATION
        # ====================================================

        print()
        print("=" * 70)
        print("DIRECT COST FINANCIAL STATEMENT SOURCE")
        print("=" * 70)

        print(
            "Workbook :",
            workbook_path,
        )

        print(
            "Worksheet:",
            sheet_name,
        )

        print(
            "Label Col:",
            "B",
        )

        # ----------------------------------------------------
        # ALIV statements change layout between months, so the
        # monthly Actual column is found from the headers (never
        # assumed to be column C). CBL keeps column C.
        # ----------------------------------------------------
        actual_column = 3
        if closed_period is not None and sheet_name == "P & L_":
            from automations.actuals.statement_columns import (
                column_letter,
                locate_aliv_columns,
            )

            _header, month_col, _ytd = locate_aliv_columns(
                list(sheet.iter_rows(values_only=True)),
                closed_period,
                workbook_path.name,
            )
            actual_column = month_col + 1

        print(
            "Actual Col:",
            openpyxl.utils.get_column_letter(actual_column),
        )

        print("=" * 70)

        # ====================================================
        # LOCATE REQUIRED LABELS
        # ====================================================

        occurrences: dict[
            str,
            list[dict[str, Any]],
        ] = {}

        for row_number in range(
            1,
            sheet.max_row + 1,
        ):

            raw_label = sheet.cell(
                row=row_number,
                column=2,
            ).value

            if raw_label is None:
                continue

            canonical_label = _canonical_statement_line(raw_label)

            if canonical_label not in required_lines:

                continue

            actual_value = sheet.cell(
                row=row_number,
                column=actual_column,
            ).value

            occurrences.setdefault(
                canonical_label,
                [],
            ).append(
                {
                    "row": row_number,
                    "label": raw_label,
                    "value": actual_value,
                }
            )

        # ====================================================
        # FIND DIRECT COST SECTION TOTAL
        # ====================================================

        total_key = _canonical_statement_line(section_total_label)

        total_occurrences = occurrences.get(
            total_key,
            [],
        )

        if not total_occurrences:

            raise RuntimeError(
                "Direct Cost section total "
                f"'{section_total_label}' "
                f"was not found in worksheet "
                f"'{sheet_name}'."
            )

        # ----------------------------------------------------
        # Use the final matching total in case the workbook
        # contains the same text elsewhere.
        # ----------------------------------------------------

        total_row = total_occurrences[-1]["row"]

        print(
            "Section Total:",
            section_total_label,
        )

        print(
            "Total Row    :",
            total_row,
        )

        print()
        print("DIRECT COST FINANCIAL STATEMENT VALUES")
        print("=" * 70)

        # ====================================================
        # EXTRACT REQUIRED VALUES
        # ====================================================

        results: dict[
            str,
            float,
        ] = {}

        for (
            canonical_label,
            output_name,
        ) in required_lines.items():

            matches = occurrences.get(
                canonical_label,
                [],
            )

            # ------------------------------------------------
            # We only want matches at or above the section
            # total.
            # ------------------------------------------------

            valid_matches = [match for match in matches if (match["row"] <= total_row)]

            if not valid_matches:

                raise RuntimeError(
                    "Required Direct Cost Financial Statement "
                    f"line '{output_name}' was not found "
                    f"before the section total on row "
                    f"{total_row}."
                )

            # ------------------------------------------------
            # Use the closest matching line to the section
            # total.
            #
            # This protects us when the same label exists
            # elsewhere in the Financial Statement.
            # ------------------------------------------------

            selected = max(
                valid_matches,
                key=lambda item: (item["row"]),
            )

            actual_value = selected["value"]

            # ------------------------------------------------
            # A blank Direct Cost actual is treated as zero.
            # ------------------------------------------------

            if actual_value is None:

                actual_value = 0.0

            try:

                results[output_name] = float(actual_value)

            except (
                TypeError,
                ValueError,
            ) as exc:

                raise RuntimeError(
                    "Direct Cost Financial Statement value "
                    f"for '{output_name}' on row "
                    f"{selected['row']} is not numeric: "
                    f"{actual_value!r}"
                ) from exc

            print(
                f"Row {selected['row']:<4} | "
                f"{output_name:<35} | "
                f"{float(actual_value):>12,.3f}"
            )

        print("=" * 70)

        return results

    finally:

        workbook.close()


# ============================================================
# ALIV DIRECT COST
# ============================================================


def read_aliv_direct_cost(
    workbook_path: str | Path,
    closed_period=None,
) -> dict[str, float]:
    """
    Read ALIV Direct Cost values.

    Worksheet:
        P & L_

    Section total:
        Direct Cost
    """

    return _read_direct_cost_section_by_label(
        workbook_path=(workbook_path),
        sheet_name=("P & L_"),
        required_lines=(ALIV_DIRECT_COST_LINES),
        section_total_label=("Direct Cost"),
        closed_period=closed_period,
    )


# ============================================================
# CBL DIRECT COST
# ============================================================


def read_cbl_direct_cost(
    workbook_path: str | Path,
) -> dict[str, float]:
    """
    Read CBL Direct Cost / Cost of Goods Sold values.

    Worksheet:
        Financials

    Section total:
        Cost of Goods Sold
    """

    return _read_direct_cost_section_by_label(
        workbook_path=(workbook_path),
        sheet_name=("Financials"),
        required_lines=(CBL_DIRECT_COST_LINES),
        section_total_label=("Cost of Goods Sold"),
    )


# ============================================================
# READ BOTH COMPANIES
# ============================================================


def read_direct_cost_financial_statements(
    *,
    aliv_workbook_path: str | Path,
    cbl_workbook_path: str | Path,
) -> dict[str, dict[str, float]]:
    """
    Read Direct Cost Financial Statement values for
    both ALIV and CBL.
    """

    return {
        "ALIV": (read_aliv_direct_cost(aliv_workbook_path)),
        "CBL": (read_cbl_direct_cost(cbl_workbook_path)),
    }
