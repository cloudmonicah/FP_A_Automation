from __future__ import annotations

from datetime import date
from pathlib import Path

from shared import settings
from shared.output_root import shared_output_root


def get_forecast_validation_output_folder(
    fiscal_year: str,
    forecast_month: int,
    forecast_year: int,
    create: bool = False,
) -> Path:
    """<shared Teams folder>\\FY27\\2026-10 - October\\Forecast Validation.

    Raises a plain-language error when the shared folder is not synced.
    """
    fiscal_year = str(fiscal_year).strip().upper()
    period = date(int(forecast_year), int(forecast_month), 1)

    validation_root = settings.get_path("forecast_validation_output_root") or shared_output_root()
    if not validation_root.is_dir():
        from shared.friendly_errors import PlainLanguageError

        raise PlainLanguageError(
            "The forecast validation folder set in settings.json was not found:\n"
            f"{validation_root}"
        )

    output_folder = (
        validation_root
        / fiscal_year
        / f"{period:%Y-%m} - {period:%B}"
        / "Forecast Validation"
    )
    if create:
        output_folder.mkdir(parents=True, exist_ok=True)
    return output_folder
