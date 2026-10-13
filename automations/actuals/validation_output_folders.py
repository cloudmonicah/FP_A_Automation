from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import os

# ============================================================
# VALIDATION OUTPUT FOLDER CONFIGURATION
# ============================================================

VALIDATION_ROOT_FOLDER = "Finance_Validation_Results"
LEGACY_VALIDATION_ROOT_FOLDER = "Validation_Results"

MONTHLY_SUBFOLDERS = (
    "Forecast Validation",
    "OPEX",
    "Revenue",
    "Direct Cost",
    "Tableau Raw Copies",
    "Logs",
)


# ============================================================
# PERIOD HELPERS
# ============================================================


def normalize_period(
    closed_period: date | datetime | str,
) -> date:

    if isinstance(
        closed_period,
        datetime,
    ):
        return closed_period.date()

    if isinstance(
        closed_period,
        date,
    ):
        return closed_period

    text = str(closed_period).strip()

    accepted_formats = (
        "%Y-%m-%d",
        "%Y-%m",
        "%b-%y",
        "%B %Y",
        "%b %Y",
        "%m/%d/%Y",
    )

    for format_string in accepted_formats:

        try:

            parsed = datetime.strptime(
                text,
                format_string,
            )

            return parsed.date()

        except ValueError:

            continue

    raise ValueError(
        "Unsupported closed period. "
        "Examples accepted: "
        "'Aug-26', "
        "'August 2026', "
        "'2026-08', "
        "or '2026-08-31'."
    )


def get_fiscal_year_label(
    closed_period: date | datetime | str,
) -> str:
    """
    Return the fiscal year label.

    Cable Bahamas fiscal year:
        July -> June

    Examples:
        August 2026 -> FY27
        June 2027   -> FY27
        July 2027   -> FY28
    """

    period = normalize_period(closed_period)

    if period.month >= 7:

        fiscal_year_end = period.year + 1

    else:

        fiscal_year_end = period.year

    return f"FY" f"{str(fiscal_year_end)[-2:]}"


def get_month_folder_name(
    closed_period: date | datetime | str,
) -> str:
    """
    Build a sortable monthly folder name.

    Example:
        August 2026
            -> 2026-08 - August
    """

    period = normalize_period(closed_period)

    return f"{period:%Y-%m} - " f"{period:%B}"


# ============================================================
# ONEDRIVE ROOT DISCOVERY
# ============================================================


def _get_possible_onedrive_roots() -> list[Path]:
    """
    Find OneDrive locations for the CURRENT Windows user.

    Nothing in this function is tied to a specific username
    or computer.

    Windows commonly exposes OneDrive locations through:
        OneDriveCommercial
        OneDrive
        OneDriveConsumer

    We also inspect the current user's home folder for
    company OneDrive folders as a fallback.
    """

    possible_roots: list[Path] = []

    # --------------------------------------------------------
    # Check standard OneDrive environment variables.
    # --------------------------------------------------------

    environment_variables = (
        "OneDriveCommercial",
        "OneDrive",
        "OneDriveConsumer",
    )

    for variable_name in environment_variables:

        value = os.environ.get(variable_name)

        if not value:
            continue

        root = Path(value)

        if root.exists() and root not in possible_roots:

            possible_roots.append(root)

    # --------------------------------------------------------
    # Look inside the CURRENT user's home directory.
    #
    # Example:
    #
    # User A:
    #     C:\Users\UserA\OneDrive - Aliv
    #
    # User B:
    #     C:\Users\UserB\OneDrive - Aliv
    #
    # No username is hard-coded.
    # --------------------------------------------------------

    user_home = Path.home()

    if user_home.exists():

        try:

            for child in user_home.iterdir():

                if not child.is_dir():
                    continue

                if not child.name.casefold().startswith("onedrive"):
                    continue

                if child not in possible_roots:

                    possible_roots.append(child)

        except (
            PermissionError,
            OSError,
        ):

            pass

    return possible_roots


# ============================================================
# FIND SHARED VALIDATION ROOT
# ============================================================


def find_actuals_validation_results_folder() -> Path:

    expected_name_fragment = VALIDATION_ROOT_FOLDER

    possible_roots = _get_possible_onedrive_roots()

    # ========================================================
    # FIRST PASS
    # Check likely OneDrive shortcut locations.
    # ========================================================

    for root in possible_roots:

        shortcut_candidates = (
            root / "Shortcuts",
            root / "My files" / "Shortcuts",
        )

        for shortcuts_folder in shortcut_candidates:

            if not shortcuts_folder.exists():
                continue

            try:

                for child in shortcuts_folder.iterdir():

                    if not child.is_dir():
                        continue

                    if any(
                        name.casefold() in child.name.casefold()
                        for name in (
                            expected_name_fragment,
                            LEGACY_VALIDATION_ROOT_FOLDER,
                        )
                    ):

                        return child

            except (
                PermissionError,
                OSError,
            ):

                continue

    # ========================================================
    # SECOND PASS
    # Check immediate folders under each OneDrive root.
    #
    # This catches situations where the user synced the
    # SharePoint/OneDrive folder differently.
    # ========================================================

    for root in possible_roots:

        try:

            for child in root.iterdir():

                if not child.is_dir():
                    continue

                if any(
                    name.casefold() in child.name.casefold()
                    for name in (expected_name_fragment, LEGACY_VALIDATION_ROOT_FOLDER)
                ):

                    return child

        except (
            PermissionError,
            OSError,
        ):

            continue

    # ========================================================
    # NOTHING FOUND
    # ========================================================

    searched = "\n".join(f" - {root}" for root in possible_roots)

    raise FileNotFoundError(
        "The shared Finance Validation Results folder "
        "could not be located automatically.\n\n"
        "Use Browse in the Finance Automation Center to "
        "select the synced validation output folder.\n\n"
        "OneDrive locations checked:\n"
        f"{searched or ' - No OneDrive folders detected'}"
    )


# ============================================================
# CREATE VALIDATION OUTPUT FOLDERS
# ============================================================


def create_validation_output_folders(
    base_output_folder: str | Path,
    closed_period: date | datetime | str,
) -> dict[
    str,
    Path,
]:
    """
    Create the folder structure for the selected validation
    period.

    The caller supplies the base folder.

    This makes the function portable between users and PCs.

    Example supplied base:
        ...\Finance_Validation_Results

    Output:
        Finance_Validation_Results
        └── FY27
            └── 2026-08 - August
                ├── OPEX
                ├── Revenue
                ├── Direct Cost
                ├── Tableau Raw Copies
                └── Logs

    NOTE:
        No Summary folder is created.
    """

    if not str(base_output_folder).strip():
        raise ValueError("Validation output folder cannot be blank.")

    base_output_folder = Path(base_output_folder).expanduser()

    # --------------------------------------------------------
    # Protect against an empty / invalid base folder.
    # --------------------------------------------------------

    if not str(base_output_folder).strip():

        raise ValueError("Validation output folder " "cannot be blank.")

    fiscal_year = get_fiscal_year_label(closed_period)

    month_folder_name = get_month_folder_name(closed_period)

    # ========================================================
    # DETERMINE VALIDATION ROOT
    # ========================================================
    #
    # If the user already selected:
    #
    #     Finance_Validation_Results
    #
    # use it directly.
    #
    # If the user selected its parent folder, create/use:
    #
    #     <selected folder>\Finance_Validation_Results
    #
    # ========================================================

    # Accept the existing root until it is renamed in OneDrive.
    # This prevents nesting the new root inside the old shared folder.
    if any(
        name.casefold() in base_output_folder.name.casefold()
        for name in (VALIDATION_ROOT_FOLDER, LEGACY_VALIDATION_ROOT_FOLDER)
    ):

        validation_root = base_output_folder

    else:

        validation_root = base_output_folder / VALIDATION_ROOT_FOLDER

    validation_root.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # FISCAL YEAR
    # ========================================================

    fiscal_year_folder = validation_root / fiscal_year

    fiscal_year_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # CLOSED MONTH
    # ========================================================

    month_folder = fiscal_year_folder / month_folder_name

    month_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # RETURNED PATH DICTIONARY
    # ========================================================

    result: dict[
        str,
        Path,
    ] = {
        "base_output_folder": base_output_folder,
        "validation_root": validation_root,
        "fiscal_year_folder": fiscal_year_folder,
        "month_folder": month_folder,
    }

    # ========================================================
    # MONTHLY VALIDATION SUBFOLDERS
    # ========================================================

    for folder_name in MONTHLY_SUBFOLDERS:

        folder_path = month_folder / folder_name

        folder_path.mkdir(
            parents=True,
            exist_ok=True,
        )

        key = folder_name.lower().replace(
            " ",
            "_",
        )

        result[key] = folder_path

    return result


# ============================================================
# LOCAL TEST
# ============================================================
#
# This block runs ONLY when this file is executed directly.
#
# It does not run when imported by the Finance Automation
# Center.
# ============================================================

if __name__ == "__main__":

    print()
    print("=" * 70)
    print("VALIDATION OUTPUT FOLDER TEST")
    print("=" * 70)

    try:

        test_base_folder = find_actuals_validation_results_folder()

        test_closed_period = "Aug-26"

        paths = create_validation_output_folders(
            base_output_folder=(test_base_folder),
            closed_period=(test_closed_period),
        )

        print(
            "Closed Period:",
            test_closed_period,
        )

        print(
            "Fiscal Year:",
            get_fiscal_year_label(test_closed_period),
        )

        print()

        for name, path in paths.items():

            print(f"{name:24} -> {path}")

        print()
        print("Folder creation test completed.")

    except Exception as exc:

        print()
        print("Automatic validation folder " "discovery was not successful.")

        print()

        print(exc)

        print()
        print(
            "This is not a production failure. "
            "The GUI can still allow the user "
            "to Browse to the validation folder."
        )
