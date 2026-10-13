from datetime import date

import pytest

from automations.actuals.validation_output_folders import (
    create_validation_output_folders,
    get_fiscal_year_label,
    get_month_folder_name,
    normalize_period,
)


@pytest.mark.parametrize(
    "text", ["Aug-26", "August 2026", "2026-08", "2026-08-31", "08/31/2026"]
)
def test_normalize_period(text):
    assert normalize_period(text).replace(day=1) == date(2026, 8, 1)


def test_fiscal_year_runs_july_to_june():
    assert get_fiscal_year_label("2026-07") == "FY27"
    assert get_fiscal_year_label("2027-06") == "FY27"
    assert get_fiscal_year_label("2027-07") == "FY28"


def test_folders_created(tmp_path):
    paths = create_validation_output_folders(base_output_folder=tmp_path, closed_period="Aug-26")
    assert paths["month_folder"] == tmp_path / "Finance_Validation_Results" / "FY27" / get_month_folder_name("Aug-26")
    assert paths["opex"].is_dir() and paths["direct_cost"].is_dir()
