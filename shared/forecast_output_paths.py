from __future__ import annotations

from datetime import date
from pathlib import Path
import os

VALIDATION_SHORTCUT_NAME = "Stacy Brown-Johnson's files - Actuals_Validation_Results"


def get_forecast_validation_output_folder(
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    create: bool = False,
) -> Path:
    """Resolve the existing Stacy Brown-Johnson OneDrive forecast validation folder."""
    fiscal_year = str(fiscal_year).strip().upper()
    period = date(int(forecast_year), int(forecast_month), 1)

    user_profile = Path(os.environ.get("USERPROFILE", str(Path.home())))
    company_onedrive = user_profile / "OneDrive - Aliv"

    one_drive_commercial = os.environ.get("OneDriveCommercial")
    if one_drive_commercial:
        candidate = Path(one_drive_commercial)
        if candidate.exists():
            company_onedrive = candidate

    validation_root = company_onedrive / "Shortcuts" / VALIDATION_SHORTCUT_NAME

    output_folder = (
        validation_root
        / fiscal_year
        / f"{period:%Y-%m} - {period:%B}"
        / "Forecast Validation"
    )

    if create:
        if not validation_root.exists():
            raise FileNotFoundError(
                "Stacy's synced validation folder was not found.\n\n"
                f"Expected:\n{validation_root}\n\n"
                "Make sure the OneDrive shortcut is available and synced."
            )
        output_folder.mkdir(parents=True, exist_ok=True)

    return output_folder
