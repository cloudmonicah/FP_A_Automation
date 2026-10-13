from __future__ import annotations

from datetime import date
from pathlib import Path
import os
import shutil
import subprocess
import xml.etree.ElementTree as ET

# ============================================================
# TABLEAU FORECAST SERVICE
# ============================================================


def _clean_tableau_name(
    value: str | None,
) -> str:
    return str(value or "").strip().strip("[]").strip().lower()


# ============================================================
# FIND TABLEAU PARAMETER
# ============================================================


def find_tableau_parameter(
    root,
    parameter_caption: str,
):

    target_name = _clean_tableau_name(parameter_caption)

    candidates = []

    # --------------------------------------------------------
    # Search every Tableau column element.
    # --------------------------------------------------------
    for column in root.iter("column"):

        caption = _clean_tableau_name(column.attrib.get("caption"))

        name = _clean_tableau_name(column.attrib.get("name"))

        # ----------------------------------------------------
        # Remember parameter-like fields for diagnostics.
        # ----------------------------------------------------
        if "param-domain-type" in column.attrib or "parameter" in name:
            candidates.append(column)

        # ----------------------------------------------------
        # Exact match against the visible parameter caption.
        # ----------------------------------------------------
        if caption == target_name:
            return column

        # ----------------------------------------------------
        # Also allow the Tableau internal field name to match.
        # ----------------------------------------------------
        if name == target_name:
            return column

    print()
    print("=" * 70)
    print("TABLEAU PARAMETERS FOUND")
    print("=" * 70)

    for column in candidates:

        print(
            "Caption:",
            column.attrib.get(
                "caption",
                "<none>",
            ),
        )

        print(
            "Name:",
            column.attrib.get(
                "name",
                "<none>",
            ),
        )

        print(
            "Current Value:",
            column.attrib.get(
                "value",
                "<none>",
            ),
        )

        print("-" * 70)

    return None


# ============================================================
# SET FORECAST START
# ============================================================


def set_forecast_start_parameter(
    tableau_workbook_path: str | Path,
    forecast_month: int,
    forecast_year: int,
    parameter_caption: str = "Forecast Start",
) -> date:
    """
    Update Forecast Start in the Tableau TWB.

    Example:

        forecast_month = 8
        forecast_year = 2026

    Tableau screen:
        8/1/2026

    Tableau XML:
        #2026-08-01#
    """

    workbook_path = Path(tableau_workbook_path)

    # ========================================================
    # VALIDATION
    # ========================================================

    if not workbook_path.exists():
        raise FileNotFoundError("Tableau workbook was not found:\n" f"{workbook_path}")

    if workbook_path.suffix.lower() != ".twb":
        raise ValueError("The Tableau forecast workbook must " "be a .twb file.")

    if forecast_month not in range(1, 13):
        raise ValueError("Forecast month must be between 1 and 12.")

    # ========================================================
    # BUILD FORECAST START DATE
    # ========================================================

    forecast_start = date(
        forecast_year,
        forecast_month,
        1,
    )

    # --------------------------------------------------------
    # Tableau's internal representation of a DATE parameter.
    #
    # Example:
    #     #2026-08-01#
    # --------------------------------------------------------
    tableau_date_value = f"#{forecast_start:%Y-%m-%d}#"

    # ========================================================
    # BACKUP
    # ========================================================
    # Keep a backup while we are testing this automation.
    # The original backup is not recreated on every run.
    # ========================================================

    backup_path = workbook_path.with_name(
        f"{workbook_path.stem}" "_BEFORE_PARAMETER_AUTOMATION.twb"
    )

    if not backup_path.exists():

        shutil.copy2(
            workbook_path,
            backup_path,
        )

        print("Tableau backup created:")

        print(backup_path)

    # ========================================================
    # READ TABLEAU XML
    # ========================================================

    tree = ET.parse(workbook_path)

    root = tree.getroot()

    # ========================================================
    # FIND FORECAST START
    # ========================================================

    parameter = find_tableau_parameter(
        root=root,
        parameter_caption=parameter_caption,
    )

    if parameter is None:
        raise RuntimeError(
            f'Tableau parameter "{parameter_caption}" ' "could not be located."
        )

    # ========================================================
    # SHOW CURRENT VALUE
    # ========================================================

    old_value = parameter.attrib.get("value")

    print()
    print("=" * 70)
    print("TABLEAU FORECAST PARAMETER")
    print("=" * 70)

    print(f"Parameter     : {parameter_caption}")

    print(f"Current value : {old_value}")

    print(f"New value     : {tableau_date_value}")

    # ========================================================
    # UPDATE PARAMETER
    # ========================================================

    parameter.set(
        "value",
        tableau_date_value,
    )

    # ========================================================
    # SAVE TABLEAU WORKBOOK
    # ========================================================

    tree.write(
        workbook_path,
        encoding="utf-8",
        xml_declaration=True,
    )

    # ========================================================
    # VERIFY THE SAVED VALUE
    # ========================================================
    # Re-open the file from disk instead of assuming that the
    # previous write succeeded.
    # ========================================================

    verify_tree = ET.parse(workbook_path)

    verify_root = verify_tree.getroot()

    verify_parameter = find_tableau_parameter(
        root=verify_root,
        parameter_caption=parameter_caption,
    )

    if verify_parameter is None:
        raise RuntimeError(
            "Forecast Start disappeared during " "Tableau workbook verification."
        )

    saved_value = verify_parameter.attrib.get("value")

    if saved_value != tableau_date_value:

        raise RuntimeError(
            "Forecast Start could not be verified.\n\n"
            f"Expected: {tableau_date_value}\n"
            f"Saved: {saved_value}"
        )

    print(f"Verified      : {saved_value}")

    print(
        "Display value : "
        f"{forecast_start.month}/"
        f"{forecast_start.day}/"
        f"{forecast_start.year}"
    )

    print("=" * 70)
    print()

    return forecast_start


# ============================================================
# OPEN TABLEAU
# ============================================================


def open_tableau_workbook(
    tableau_workbook_path: str | Path,
    tableau_exe: str | Path | None = None,
) -> None:
    """
    Open the modified Tableau forecast workbook.

    IMPORTANT:
    The Forecast Start parameter must already have been updated
    before this function is called.
    """

    workbook_path = Path(tableau_workbook_path)

    if not workbook_path.exists():
        raise FileNotFoundError("Tableau workbook was not found:\n" f"{workbook_path}")

    # --------------------------------------------------------
    # Use a specific Tableau Desktop executable if supplied.
    # --------------------------------------------------------
    if tableau_exe:

        tableau_path = Path(tableau_exe)

        if not tableau_path.exists():
            raise FileNotFoundError(
                "Tableau Desktop executable " "was not found:\n" f"{tableau_path}"
            )

        subprocess.Popen(
            [
                str(tableau_path),
                str(workbook_path),
            ]
        )

        return

    # --------------------------------------------------------
    # Otherwise let Windows use the registered .twb
    # application, which should be Tableau Desktop.
    # --------------------------------------------------------
    os.startfile(str(workbook_path))
