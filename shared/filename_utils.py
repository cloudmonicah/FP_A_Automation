from pathlib import Path
import re


def clean_budget_title(
    filename: str,
    source_fiscal_year: str,
) -> str:
    title = Path(filename).stem

    title = title.replace("_", " ")
    title = title.replace("-", " ")

    patterns = [
        rf"\b{re.escape(source_fiscal_year)}\b",
        r"\bFINAL\b",
        r"\bBUDGET\s+SUBMISSION\s+TEMPLATE\b",
        r"\bBUDGET\s+SUBMISSION\b",
        r"\bBUDGET\s+TEMPLATE\b",
        r"\bTEMPLATE\b",
        r"\bBUDGET\b",
    ]

    for pattern in patterns:
        title = re.sub(
            pattern,
            " ",
            title,
            flags=re.IGNORECASE,
        )

    return re.sub(
        r"\s+",
        " ",
        title,
    ).strip()


def build_forecast_filename(
    original_filename: str,
    source_fiscal_year: str,
    target_fiscal_year: str,
) -> str:
    title = clean_budget_title(
        filename=original_filename,
        source_fiscal_year=source_fiscal_year,
    )

    extension = Path(original_filename).suffix

    return f"{target_fiscal_year} " f"{title}_Jul Forecast" f"{extension}"
