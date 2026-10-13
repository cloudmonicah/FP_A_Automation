from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

# ============================================================
# REVENUE FINANCIAL STATEMENT READER
# ============================================================
#
# PURPOSE
# -------
# Read finalized ALIV and CBL Financial Statements and extract
# the Revenue Actual values required for Tableau validation.
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
# Required Revenue lines are located dynamically by label.
# This protects the automation when rows are inserted or moved
# in future Financial Statement workbooks.
#
# This file contains REVENUE logic only.
#
# OPEX and Direct Cost readers remain completely separate.
# ============================================================


# ============================================================
# ALIV REVENUE LINES
# ============================================================
#
# Based on the ALIV Financial Statement Revenue section:
#
# Consumer Prepaid
# Consumer Postpaid
# Corporate
# Visitor Roaming
# Interconnect
# Other
# Service Revenues
#
# "Service Revenues" is the section total/control line.
# ============================================================

ALIV_REVENUE_LINES = {
    "CONSUMER PREPAID": "Consumer Prepaid",
    "CONSUMER POSTPAID": "Consumer Postpaid",
    "CORPORATE": "Corporate",
    "VISITOR ROAMING": "Visitor Roaming",
    "INTERCONNECT": "Interconnect",
    "OTHER": "Other",
    "SERVICE REVENUES": "Service Revenues",
}


# ============================================================
# CBL REVENUE LINES
# ============================================================
#
# Based on the CBL Financial Statement Gross Revenue section:
#
# Residential
# Corporate - Recurring
# Corporate - One off
# Media
# Other
# MSA-ALIV
# Government
# Gross Revenue
#
# "Gross Revenue" is the section total/control line.
# ============================================================

CBL_REVENUE_LINES = {
    "RESIDENTIAL": "Residential",
    "CORPORATE - RECURRING": "Corporate - Recurring",
    "CORPORATE - ONE OFF": "Corporate - One off",
    "MEDIA": "Media",
    "OTHER": "Other",
    "MSA-ALIV": "MSA-ALIV",
    "GOVERNMENT": "Government",
    "GROSS REVENUE": "Gross Revenue",
}


# ============================================================
# LABEL NORMALIZATION
# ============================================================


def _canonical_statement_line(
    value: Any,
) -> str:
    """
    Normalize a Financial Statement line label.

    This allows the reader to tolerate differences such as:

        extra spaces
        capitalization
        leading numeric account/section codes
        "&" versus "and"

    Example:

        "20 General and Administration"
            -> "GENERAL AND ADMINISTRATION"
    """

    if value is None:
        return ""

    text = str(value).strip()

    # --------------------------------------------------------
    # Remove a leading numeric code when one exists.
    #
    # Example:
    #
    #     "20 General and Administration"
    #
    # becomes:
    #
    #     "General and Administration"
    # --------------------------------------------------------

    parts = text.split(
        " ",
        1,
    )

    if len(parts) == 2 and parts[0].isdigit():

        text = parts[1].strip()

    # --------------------------------------------------------
    # Normalize "&" versus "and".
    # --------------------------------------------------------

    text = text.replace(
        "&",
        "and",
    )

    # --------------------------------------------------------
    # Normalize common dash characters.
    #
    # This helps with labels such as:
    #
    #     Corporate - Recurring
    #     Corporate – Recurring
    # --------------------------------------------------------

    text = text.replace("–", "-").replace("—", "-")

    # --------------------------------------------------------
    # Normalize whitespace and capitalization.
    # --------------------------------------------------------

    return " ".join(text.upper().split())


# ============================================================
# GENERIC REVENUE SECTION READER
# ============================================================


def _read_revenue_section_by_label(
    *,
    workbook_path: str | Path,
    sheet_name: str,
    required_lines: dict[str, str],
    section_total_label: str,
) -> dict[str, float]:
    """
    Read one Revenue section from a Financial Statement.

    Expected layout:

        Column B = Financial Statement line
        Column C = Actual ($000)

    The function scans the worksheet dynamically and selects
    the required Revenue labels closest to the Revenue section
    total.

    This avoids hard-coding row numbers.
    """

    workbook_path = Path(workbook_path)

    # ========================================================
    # VALIDATE SOURCE FILE
    # ========================================================

    if not workbook_path.exists():

        raise FileNotFoundError(
            "Revenue Financial Statement was not found:\n" f"{workbook_path}"
        )

    # ========================================================
    # OPEN WORKBOOK
    # ========================================================
    #
    # data_only=True:
    #     read cached Excel formula results instead of formulas.
    #
    # read_only=True:
    #     improve performance because this is a source read.
    # ========================================================

    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
    )

    try:

        # ====================================================
        # VALIDATE WORKSHEET
        # ====================================================

        if sheet_name not in workbook.sheetnames:

            raise ValueError(
                "Expected Revenue worksheet "
                f"'{sheet_name}' was not found in:\n"
                f"{workbook_path}"
            )

        sheet = workbook[sheet_name]

        # ====================================================
        # AUDIT / DIAGNOSTIC OUTPUT
        # ====================================================

        print()
        print("=" * 70)
        print("REVENUE FINANCIAL STATEMENT SOURCE")
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

        print(
            "Actual Col:",
            "C",
        )

        print("=" * 70)

        # ====================================================
        # FIND ALL OCCURRENCES OF REQUIRED LABELS
        # ====================================================
        #
        # We collect all occurrences instead of stopping at the
        # first one because labels such as "Other" can appear
        # multiple times in a Financial Statement.
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
                column=3,
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
        # FIND REVENUE SECTION TOTAL
        # ====================================================

        total_key = _canonical_statement_line(section_total_label)

        total_occurrences = occurrences.get(
            total_key,
            [],
        )

        if not total_occurrences:

            raise RuntimeError(
                "Revenue section total "
                f"'{section_total_label}' "
                f"was not found in worksheet "
                f"'{sheet_name}'."
            )

        # ----------------------------------------------------
        # Use the first matching Revenue total.
        #
        # This is intentional for Revenue because the Revenue
        # section appears toward the top of both Financial
        # Statements and similarly named totals can potentially
        # appear elsewhere later in a workbook.
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
        print("REVENUE FINANCIAL STATEMENT VALUES")
        print("=" * 70)

        # ====================================================
        # EXTRACT REQUIRED REVENUE VALUES
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
            # Only consider matches at or above the Revenue
            # section total.
            # ------------------------------------------------

            valid_matches = [match for match in matches if (match["row"] <= total_row)]

            if not valid_matches:

                raise RuntimeError(
                    "Required Revenue Financial Statement "
                    f"line '{output_name}' was not found "
                    f"before the Revenue section total "
                    f"on row {total_row}."
                )

            # ------------------------------------------------
            # Choose the occurrence closest to the Revenue
            # section total.
            #
            # This is especially important for generic labels
            # such as "Other".
            # ------------------------------------------------

            selected = max(
                valid_matches,
                key=lambda item: (item["row"]),
            )

            actual_value = selected["value"]

            # ------------------------------------------------
            # A blank required Revenue actual is treated as 0.
            # ------------------------------------------------

            if actual_value is None:
                actual_value = 0.0

            try:

                numeric_value = float(actual_value)

            except (
                TypeError,
                ValueError,
            ) as exc:

                raise RuntimeError(
                    "Revenue Financial Statement value "
                    f"for '{output_name}' on row "
                    f"{selected['row']} is not numeric: "
                    f"{actual_value!r}"
                ) from exc

            results[output_name] = numeric_value

            print(
                f"Row {selected['row']:<4} | "
                f"{output_name:<30} | "
                f"{numeric_value:>12,.3f}"
            )

        print("=" * 70)

        return results

    finally:

        workbook.close()


# ============================================================
# ALIV REVENUE READER
# ============================================================


def read_aliv_revenue(
    workbook_path: str | Path,
) -> dict[str, float]:
    """
    Read the ALIV Revenue section.

    Worksheet:
        P & L_

    Section total:
        Service Revenues
    """

    return _read_revenue_section_by_label(
        workbook_path=(workbook_path),
        sheet_name=("P & L_"),
        required_lines=(ALIV_REVENUE_LINES),
        section_total_label=("Service Revenues"),
    )


# ============================================================
# CBL REVENUE READER
# ============================================================


def read_cbl_revenue(
    workbook_path: str | Path,
) -> dict[str, float]:
    """
    Read the CBL Gross Revenue section.

    Worksheet:
        Financials

    Section total:
        Gross Revenue
    """

    return _read_revenue_section_by_label(
        workbook_path=(workbook_path),
        sheet_name=("Financials"),
        required_lines=(CBL_REVENUE_LINES),
        section_total_label=("Gross Revenue"),
    )


# ============================================================
# READ BOTH COMPANIES
# ============================================================


def read_revenue_financial_statements(
    *,
    aliv_workbook_path: str | Path,
    cbl_workbook_path: str | Path,
) -> dict[str, dict[str, float]]:
    """
    Read Revenue Financial Statement values for both companies.

    Returned structure:

        {
            "ALIV": {
                "Consumer Prepaid": ...,
                "Consumer Postpaid": ...,
                ...
            },

            "CBL": {
                "Residential": ...,
                "Corporate - Recurring": ...,
                ...
            }
        }
    """

    return {
        "ALIV": read_aliv_revenue(aliv_workbook_path),
        "CBL": read_cbl_revenue(cbl_workbook_path),
    }
