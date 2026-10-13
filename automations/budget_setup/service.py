from __future__ import annotations

from pathlib import Path
import shutil

from automations.budget_setup.workbook_service import (
    convert_budget_workbook,
    get_formula_cells,
    compare_formula_preservation,
)
from shared.filename_utils import build_forecast_filename


def create_initial_forecasts(
    source_folder: str | Path,
    output_folder: str | Path,
    source_fiscal_year: str,
    target_fiscal_year: str,
    overwrite_mode: str = "skip",
) -> dict:
    """
    Create the initial fiscal-year forecast workbooks from
    approved budget workbooks.

    PROCESS
    -------
    1. Find eligible budget workbooks.
    2. Capture the source workbook formulas.
    3. Copy the source workbook.
    4. Convert the copied workbook to the new forecast structure.
    5. Compare formulas before and after conversion.
    6. Report created, skipped, failed, and formula-audit results.

    IMPORTANT
    ---------
    The original source budget workbook is never modified.
    """

    # ========================================================
    # NORMALIZE INPUT PATHS
    # ========================================================

    source_path = Path(source_folder)
    output_path = Path(output_folder)

    # ========================================================
    # VALIDATE SOURCE FOLDER
    # ========================================================

    if not source_path.exists():
        raise FileNotFoundError(f"Budget source folder not found: {source_path}")

    # ========================================================
    # CREATE OUTPUT FOLDER IF NEEDED
    # ========================================================

    output_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # FIND ELIGIBLE BUDGET WORKBOOKS
    # ========================================================

    budget_files = [
        file
        for file in source_path.glob("*.xls*")
        if not file.name.startswith("~$")
        and source_fiscal_year.upper() in file.name.upper()
        and "BUDGET" in file.name.upper()
    ]

    # ========================================================
    # PROCESS RESULTS
    # ========================================================

    created = []
    skipped = []
    failed = []

    formula_audits = []

    worksheets_updated = 0

    # ========================================================
    # PROCESS EACH BUDGET WORKBOOK
    # ========================================================

    for source_file in budget_files:

        # ----------------------------------------------------
        # Build the new forecast filename.
        # ----------------------------------------------------

        new_name = build_forecast_filename(
            original_filename=source_file.name,
            source_fiscal_year=source_fiscal_year,
            target_fiscal_year=target_fiscal_year,
        )

        destination = output_path / new_name

        try:

            # ====================================================
            # HANDLE EXISTING OUTPUT FILE
            # ====================================================

            if destination.exists():

                if overwrite_mode == "skip":
                    skipped.append(destination)
                    continue

                if overwrite_mode == "cancel":
                    raise RuntimeError(
                        "Forecast file already exists: " f"{destination.name}"
                    )

                if overwrite_mode == "replace":
                    destination.unlink()

            # ====================================================
            # CAPTURE SOURCE FORMULAS BEFORE COPYING
            # ====================================================
            #
            # This allows us to prove whether formulas survive
            # the Budget -> Forecast conversion.
            # ====================================================

            source_formulas = get_formula_cells(source_file)

            # ====================================================
            # COPY ORIGINAL WORKBOOK
            # ====================================================
            #
            # shutil.copy2 creates a full file-level copy.
            # The original approved budget workbook is preserved.
            # ====================================================

            shutil.copy2(
                source_file,
                destination,
            )

            # ====================================================
            # CONVERT COPIED WORKBOOK TO FORECAST
            # ====================================================

            result = convert_budget_workbook(
                workbook_path=destination,
                target_fiscal_year=target_fiscal_year,
            )

            # ====================================================
            # AUDIT FORMULAS AFTER CONVERSION
            # ====================================================

            destination_formulas = get_formula_cells(destination)

            formula_audit = compare_formula_preservation(
                source_formulas=source_formulas,
                destination_formulas=destination_formulas,
            )

            # ----------------------------------------------------
            # Add the workbook name to the audit result.
            # ----------------------------------------------------

            formula_audits.append(
                {
                    "file": destination.name,
                    **formula_audit,
                }
            )

            # ====================================================
            # DISPLAY FORMULA AUDIT
            # ====================================================

            print()
            print("=" * 70)
            print("FORMULA PRESERVATION AUDIT")
            print("=" * 70)

            print(f"File: " f"{destination.name}")

            print(f"Source formulas: " f"{formula_audit['source_formula_count']}")

            print(f"Output formulas: " f"{formula_audit['destination_formula_count']}")

            print(f"Missing formulas: " f"{formula_audit['missing_formula_count']}")

            print(f"Changed formulas: " f"{formula_audit['changed_formula_count']}")

            print(f"Formulas preserved: " f"{formula_audit['formulas_preserved']}")

            print("=" * 70)

            # ====================================================
            # RECORD SUCCESS
            # ====================================================

            worksheets_updated += result["worksheets_updated"]

            created.append(destination)

        # ========================================================
        # RECORD FAILURE
        # ========================================================

        except Exception as exc:

            failed.append(
                {
                    "file": source_file.name,
                    "path": str(source_file),
                    "error": (f"{type(exc).__name__}: {exc}"),
                }
            )

    # ========================================================
    # RETURN PROCESS RESULTS
    # ========================================================

    return {
        "files_created": len(created),
        "files_skipped": len(skipped),
        "files_failed": len(failed),
        "worksheets_updated": worksheets_updated,
        "output_folder": str(output_path),
        "created_files": created,
        "skipped_files": skipped,
        "failures": failed,
        "formula_audits": formula_audits,
    }
