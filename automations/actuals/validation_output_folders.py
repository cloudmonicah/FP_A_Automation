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
# SHARED VALIDATION ROOT
# ============================================================


def find_actuals_validation_results_folder() -> Path:
    """The Teams "Validation Results" folder (or validation_output_root in settings.json).

    There is no search for other people's folders: the shared folder must be
    added to this person's OneDrive, or the run stops with a plain message.
    """
    from shared import settings
    from shared.friendly_errors import PlainLanguageError
    from shared.output_root import shared_output_root

    configured = settings.get_path("validation_output_root")
    if configured is not None:
        if not configured.is_dir():
            raise PlainLanguageError(
                "The validation output folder set in settings.json was not found:\n"
                f"{configured}\n\nCheck the folder is synced, or update settings.json."
            )
        return configured
    return shared_output_root()


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
        ...\\Finance_Validation_Results

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
    from shared.output_root import SHARED_FOLDER_NAME

    # The shared Teams folder holds the FY folders directly, like Forecast.
    if base_output_folder.name.casefold() == SHARED_FOLDER_NAME.casefold() or any(
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
