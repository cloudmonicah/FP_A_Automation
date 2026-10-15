from __future__ import annotations

from pathlib import Path
from typing import Any

import openpyxl

# ============================================================
# FINANCIAL STATEMENT READER
# ============================================================
#
# PURPOSE
# -------
# Read finalized ALIV and CBL financial statements and extract
# OPEX Actual values for Actuals Validation.
#
# CURRENT BUSINESS RULES
# ----------------------
# Column B = Financial Statement line
# Column C = Actual
#
# Budget is ignored for this validation.
#
# IMPORTANT
# ---------
# We DO NOT depend on fixed row numbers.
#
# The reader scans Column B and matches required OPEX lines
# by name. This allows rows to move or new rows to be inserted
# without breaking the validation.
#
# Financial Statement values are already in $000s.
# Do NOT divide them by 1,000 here.
# ============================================================


# ============================================================
# REQUIRED ALIV OPEX LINES
# ============================================================

ALIV_OPEX_LINES = {
    "GENERAL AND ADMINISTRATION": "General and Administration",
    "HUMAN RESOURCES": "Human Resources",
    "FINANCIAL SERVICES": "Financial Services",
    "LEGAL": "Legal",
    "FACILITIES": "Facilities",
    "INFORMATION TECHNOLOGY": "Information Technology",
    "TECHNICAL OPERATIONS": "Technical Operations",
    "COMMERCIAL AND PRODUCT DEVELOPMENT": "Commercial and Product Development",
    "MARKETING": "Marketing",
    "CONSUMER SALES": "Consumer Sales",
    "CONTACT CENTER": "Contact Center",
    "ROAMING": "Roaming",
    "CORPORATE SALES": "Corporate Sales",
    "OPERATIONS": "Operations",
    "OPERATING EXPENSES": "Operating Expenses",
}

# ============================================================
# REQUIRED CBL OPEX LINES
# ============================================================

CBL_OPEX_LINES = {
    "GENERAL AND ADMINISTRATION": "General and Administration",
    "HUMAN RESOURCES": "Human Resources",
    "FINANCIAL SERVICES": "Financial Services",
    "LEGAL": "Legal",
    "FACILITIES": "Facilities",
    "INFORMATION TECHNOLOGY": "Information Technology",
    "TECHNICAL OPERATIONS": "Technical Operations",
    "COMMERCIAL AND PRODUCT DEVELOPMENT": "Commercial and Product Development",
    "MARKETING": "Marketing",
    "CONSUMER SALES": "Consumer Sales",
    "SERVICE DELIVERY": "Service Delivery",
    "CONTACT CENTER": "Contact Center",
    "CORPORATE SALES": "Corporate Sales",
    "OUTSIDE PLANT": "Outside Plant",
    "OPERATIONS": "Operations",
    "MEDIA": "Media",
    "Total Opex": "Total Operating Costs",
}


# ============================================================
# LABEL NORMALIZATION
# ============================================================


def _canonical_statement_line(
    value: Any,
) -> str:

    if value is None:
        return ""

    text = str(value).strip()

    parts = text.split(
        " ",
        1,
    )

    if len(parts) == 2 and parts[0].isdigit():
        text = parts[1].strip()
    elif len(parts) == 2 and parts[0].upper() == "MD" and parts[1].strip().upper() == "MEDIA":
        # The statement labels the Media line "MD Media" (MD is its code).
        text = "Media"

    # --------------------------------------------------------
    # Standardize "&" and "and".
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
# GENERIC OPEX READER
# ============================================================


def _read_opex_by_label(
    *,
    workbook_path,
    sheet_name,
    required_lines,
    section_total_label,
) -> dict[str, float]:

    workbook_path = Path(workbook_path)

    if not workbook_path.exists():

        raise FileNotFoundError("Financial Statement not found: " f"{workbook_path}")

    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
    )

    try:

        if sheet_name not in workbook.sheetnames:

            raise ValueError(f"Expected worksheet " f"'{sheet_name}' was not found.")

        sheet = workbook[sheet_name]

        print()
        print("=" * 70)
        print("FINANCIAL STATEMENT SOURCE")
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
        # FIND ALL REQUIRED OPEX LABEL OCCURRENCES
        # ====================================================

        occurrences = {}

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
        # FIND THE OPEX TOTAL / END OF SECTION
        # ====================================================

        total_key = _canonical_statement_line(section_total_label)

        total_occurrences = occurrences.get(
            total_key,
            [],
        )

        if not total_occurrences:

            raise RuntimeError(
                f"OPEX section total "
                f"'{section_total_label}' "
                f"was not found in "
                f"'{sheet_name}'."
            )

        total_row = total_occurrences[-1]["row"]

        print(
            "OPEX Total:",
            section_total_label,
        )

        print(
            "Total Row :",
            total_row,
        )

        print()
        print("FINANCIAL STATEMENT OPEX VALUES")
        print("=" * 70)

        results = {}

        for (
            canonical_label,
            output_name,
        ) in required_lines.items():

            matches = occurrences.get(
                canonical_label,
                [],
            )

            # ------------------------------------------------
            # Keep only matches at or above the OPEX total.
            # ------------------------------------------------

            valid_matches = [match for match in matches if match["row"] <= total_row]

            if not valid_matches:

                raise RuntimeError(
                    "Required OPEX Financial Statement "
                    f"line '{output_name}' was not found "
                    f"before the OPEX total on row "
                    f"{total_row}."
                )

            selected = max(
                valid_matches,
                key=lambda item: item["row"],
            )

            actual_value = selected["value"]

            if actual_value is None:

                actual_value = 0.0

            results[output_name] = float(actual_value)

            print(
                f"Row {selected['row']:<4} | "
                f"{output_name:<40} | "
                f"{float(actual_value):>12,.3f}"
            )

        print("=" * 70)

        return results

    finally:

        workbook.close()


# ============================================================
# ALIV OPEX
# ============================================================


def read_aliv_opex(workbook_path, closed_period=None) -> dict[str, float]:
    if closed_period is None:
        raise ValueError(
            "ALIV OPEX requires selected closed period to discover Actual columns"
        )
    from automations.actuals.adaptive_aliv_statement_reader import read_aliv_period
    from automations.actuals.adaptive_aliv_statement_reader import norm

    data = read_aliv_period(workbook_path, closed_period)
    output = {}
    for key, label in ALIV_OPEX_LINES.items():
        pair = data.get(norm(key))
        if pair is None:
            raise ValueError(
                f"ALIV OPEX statement missing/ambiguous actual for {label}"
            )
        output[label] = pair[0]
    return output


# ============================================================
# CBL OPEX
# ============================================================


def read_cbl_opex(
    workbook_path,
) -> dict[str, float]:

    return _read_opex_by_label(
        workbook_path=(workbook_path),
        sheet_name=("Financials"),
        required_lines=(CBL_OPEX_LINES),
        section_total_label=("Total Operating Costs"),
    )


# ============================================================
# READ BOTH COMPANIES
# ============================================================


def read_opex_financial_statements(
    *,
    aliv_workbook_path: str | Path,
    cbl_workbook_path: str | Path,
) -> dict[
    str,
    dict[
        str,
        float | None,
    ],
]:
    """
    Read the OPEX Actual values for both ALIV and CBL.
    """

    return {
        "ALIV": read_aliv_opex(aliv_workbook_path),
        "CBL": read_cbl_opex(cbl_workbook_path),
    }


def _read_financial_statement_section_by_label(
    *,
    workbook_path,
    sheet_name,
    required_lines,
    section_total_label,
) -> dict[str, float]:

    workbook_path = Path(workbook_path)

    if not workbook_path.exists():

        raise FileNotFoundError("Financial Statement not found: " f"{workbook_path}")

    workbook = openpyxl.load_workbook(
        workbook_path,
        data_only=True,
        read_only=True,
    )

    try:

        if sheet_name not in workbook.sheetnames:

            raise ValueError(f"Expected worksheet " f"'{sheet_name}' was not found.")

        sheet = workbook[sheet_name]

        occurrences = {}

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

        total_key = _canonical_statement_line(section_total_label)

        total_occurrences = occurrences.get(
            total_key,
            [],
        )

        if not total_occurrences:

            raise RuntimeError(
                f"Section total "
                f"'{section_total_label}' "
                f"was not found in "
                f"'{sheet_name}'."
            )

        total_row = total_occurrences[-1]["row"]

        results = {}

        print()
        print("=" * 70)
        print("DIRECT COST FINANCIAL STATEMENT SOURCE")
        print("=" * 70)
        print("Workbook :", workbook_path)
        print("Worksheet:", sheet_name)
        print("Label Col:", "B")
        print("Actual Col:", "C")
        print("Total Row :", total_row)
        print("=" * 70)

        for (
            canonical_label,
            output_name,
        ) in required_lines.items():

            matches = occurrences.get(
                canonical_label,
                [],
            )

            valid_matches = [match for match in matches if match["row"] <= total_row]

            if not valid_matches:

                raise RuntimeError(
                    "Required Direct Cost line " f"'{output_name}' was not found."
                )

            selected = max(
                valid_matches,
                key=lambda item: item["row"],
            )

            actual_value = selected["value"]

            if actual_value is None:
                actual_value = 0.0

            results[output_name] = float(actual_value)

            print(
                f"Row {selected['row']:<4} | "
                f"{output_name:<35} | "
                f"{float(actual_value):>12,.3f}"
            )

        print("=" * 70)

        return results

    finally:

        workbook.close()
