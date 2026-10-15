"""Step 1: create the month's forecast workbooks from last month's files."""

import pytest
from openpyxl import Workbook

from automations.forecast.excel_service import ForecastFilesExistError, create_forecast_refresh
from shared.friendly_errors import GENERIC, friendly


def _workbook(path, forecast_type="FY27 Forecast_092026"):
    wb = Workbook()
    ws = wb.active
    ws.title = "ALIV"
    ws.append(["Forecast Type", "GL Account", "Jul", "Grand Total"])
    ws.append([forecast_type, "6000", 5, "=SUM(C2:C2)"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


@pytest.fixture
def folders(tmp_path):
    source, out = tmp_path / "Sep", tmp_path / "Oct"
    _workbook(source / "FY27 OPEX Legal_Sep Forecast.xlsx")
    _workbook(source / "FY27 OPEX HR_Sep Forecast.xlsx")
    _workbook(source / "FY27 OPEX IT_Aug Forecast.xlsx", "FY27 Forecast_082026")
    return source, out


def _run(source, out, mode="skip"):
    return create_forecast_refresh(source, out, "FY27", 10, 2026, overwrite_mode=mode)


def test_creates_every_family_and_reports_older_sources(folders):
    source, out = folders
    result = _run(source, out)
    assert result["files_created"] == 3 and not result["failures"]
    assert sorted(p.name for p in out.iterdir()) == [
        "FY27 OPEX HR_Oct Forecast.xlsx",
        "FY27 OPEX IT_Oct Forecast.xlsx",
        "FY27 OPEX Legal_Oct Forecast.xlsx",
    ]
    assert [item["family"] for item in result["older_sources"]] == ["FY27 OPEX IT"]


def test_cancel_changes_nothing_when_files_exist(folders):
    source, out = folders
    _run(source, out)
    legal = out / "FY27 OPEX Legal_Oct Forecast.xlsx"
    (out / "FY27 OPEX HR_Oct Forecast.xlsx").unlink()
    before = legal.read_bytes()
    with pytest.raises(ForecastFilesExistError) as error:
        _run(source, out, "cancel")
    assert "2 of this month's forecast files already exist" in str(error.value)
    assert not (out / "FY27 OPEX HR_Oct Forecast.xlsx").exists()  # nothing was created
    assert legal.read_bytes() == before
    assert friendly(error.value) == str(error.value)  # shown as written


def test_skip_keeps_existing_files(folders):
    source, out = folders
    _run(source, out)
    result = _run(source, out, "skip")
    assert result["files_created"] == 0 and len(result["skipped_files"]) == 3


@pytest.mark.parametrize(
    "error, expected",
    [
        (ValueError("July is the initial fiscal-year forecast month. Use Initial Forecast Month Setup to create July."),
         "Initial Budget Setup"),
        (FileNotFoundError(r"No valid FY27 forecast workbooks were found before Oct in C:\Forecast\Sep"),
         r"C:\Forecast\Sep"),
        (FileNotFoundError(r"Forecast source folder not found: C:\Forecast\Sep"),
         r"Previous Forecast Folder was not found: C:\Forecast\Sep"),
    ],
)
def test_step1_errors_are_explained(error, expected):
    message = friendly(error)
    assert message != GENERIC and expected in message
