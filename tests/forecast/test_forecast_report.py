import os

from openpyxl import load_workbook

from automations.forecast.validation_service import summarize_results
from shared import forecast_validation_report as report

RESULTS = [
    {"forecast_workbook": "A_Oct Forecast.xlsx", "worksheet": "ALIV", "status": "PASS",
     "excel_grand_total": 10.0, "tableau_grand_total": 10.0, "variance": 0.0},
    {"forecast_workbook": "B_Oct Forecast.xlsx", "worksheet": "CBL", "status": "NOT VALIDATED",
     "note": "B_Oct Forecast.xlsx / CBL is for September 2026"},
    {"forecast_workbook": "C_Oct Forecast.xlsx", "worksheet": "ALIV", "status": "ACTION REQUIRED",
     "note": "RuntimeError: Tableau workbook not ready yet"},
]


def _write(folder):
    return report.write_forecast_validation_report(
        validation_result=summarize_results(RESULTS),
        output_folder=folder,
        fiscal_year="FY27",
        forecast_month=10,
        forecast_year=2026,
        forecast_type="FY27 Forecast_102026",
    )


def _summary(path):
    wb = load_workbook(path)
    try:
        ws = wb["Validation Summary"]
        values = {ws.cell(r, 1).value: ws.cell(r, 2).value for r in range(1, ws.max_row + 1)}
        details = [row[7] for row in wb["Validation Detail"].iter_rows(min_row=2, values_only=True)]
        return values, details
    finally:
        wb.close()


def test_every_route_and_status_is_reported(tmp_path):
    path = _write(tmp_path)
    assert path.name == "Forecast_Validation_Oct_2026.xlsx"
    values, details = _summary(path)
    assert details == ["PASS", "NOT VALIDATED", "ACTION REQUIRED"]
    assert values["Not Validated"] == 1
    assert values["Action Required"] == 1
    assert values["Overall Status"] == "REVIEW REQUIRED"
    assert not list(tmp_path.glob("~staging_*"))


def test_report_open_in_excel_is_saved_under_another_name(tmp_path, monkeypatch):
    first = _write(tmp_path)
    real_replace = os.replace

    def locked(src, dst):
        if os.fspath(dst) == os.fspath(first):
            raise PermissionError("in use")
        return real_replace(src, dst)

    monkeypatch.setattr(report.os, "replace", locked)
    second = _write(tmp_path)
    assert second != first and second.exists()
    assert second.name.startswith("Forecast_Validation_Oct_2026 (")
    assert not list(tmp_path.glob("~staging_*"))
