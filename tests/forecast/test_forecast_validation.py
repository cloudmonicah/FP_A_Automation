import xml.etree.ElementTree as ET
from datetime import date

from automations.forecast import validation_service as vs
from tests.forecast.conftest import (
    FORECAST_MEASURE_TAB,
    FORECAST_TYPE_TAB,
    write_forecast_workbook,
)

OCT_2026 = date(2026, 10, 1)


def _members(path, column_suffix):
    root = ET.parse(path).getroot()
    return [
        g.get("member").strip('"')
        for f in root.iter("filter")
        if f.get("column", "").endswith(column_suffix)
        for g in f.iter("groupfilter")
        if g.get("function") == "member"
    ]


def test_forecast_type_tab_is_prepared(twb, tmp_path):
    out = tmp_path / "copy.twb"
    prepared = vs._prepare_individual_forecast_copy(
        source=twb(FORECAST_TYPE_TAB),
        destination=out,
        tab="Monthly Forecast",
        company="CBL",
        forecast_start=OCT_2026,
        forecast_type="FY27 Forecast_102026",
    )
    assert prepared["period_mode"] == "Forecast Type"
    assert _members(out, "[none:Forecast Type:nk]") == ["FY27 Forecast_102026"]
    assert _members(out, "[none:Company:nk]") == ["CBL"]


def test_forecast_measure_tab_is_prepared(twb, tmp_path):
    out = tmp_path / "copy.twb"
    prepared = vs._prepare_individual_forecast_copy(
        source=twb(FORECAST_MEASURE_TAB),
        destination=out,
        tab="Monthly Forecast",
        company="ALIV",
        forecast_start=OCT_2026,
        forecast_type="FY27 Forecast_102026",
    )
    assert prepared["period_mode"] == "Forecast Measure"
    measure = [c for c in ET.parse(out).getroot().iter("column") if c.get("caption") == "Forecast Measure"][0]
    assert measure.get("value") == '"Forecast"'


def test_source_twb_is_never_modified(twb, tmp_path):
    source = twb(FORECAST_TYPE_TAB)
    before = source.read_bytes()
    vs._prepare_individual_forecast_copy(
        source=source,
        destination=tmp_path / "copy.twb",
        tab="Monthly Forecast",
        company="ALIV",
        forecast_start=OCT_2026,
        forecast_type="FY27 Forecast_102026",
    )
    assert source.read_bytes() == before


def test_excel_total_found_from_headers_not_positions(tmp_path):
    path = write_forecast_workbook(
        tmp_path / "x.xlsx", rows=((100, 200), (5,)), extra_columns=3
    )
    total = vs.get_excel_forecast_grand_total(path, "ALIV", "FY27 Forecast_102026")
    assert total == 305


def test_wrong_month_excel_file_is_rejected_naming_both_months(tmp_path):
    path = write_forecast_workbook(tmp_path / "x.xlsx", forecast_type="FY27 Forecast_092026")
    try:
        vs.get_excel_forecast_grand_total(path, "ALIV", "FY27 Forecast_102026")
    except vs.WrongPeriodError as exc:
        message = str(exc)
    else:
        raise AssertionError("wrong-month workbook was accepted")
    assert "September 2026" in message and "October 2026" in message


def test_tableau_grand_total_found_from_headers():
    rows = [
        {"GL Company": "ALIV", "Account": "6000", "Total": "10"},
        {"GL Company": "Grand Total", "Account": "", "Total": "(1,250)"},
    ]
    assert vs.get_individual_tableau_grand_total(rows) == -1250


ROUTES = [
    {
        "workbook": "FY27 OPEX Legal_Aug Forecast.xlsx",
        "excel_sheet": "ALIV",
        "individual_tableau_folder": "FOLDER",
        "individual_tableau_workbook": "Legal.twb",
        "individual_tableau_worksheet": "Monthly Forecast",
        "company_filter": "ALIV",
        "validation_status": "confirmed",
    },
    {
        "workbook": "FY27 OPEX Legal_Aug Forecast.xlsx",
        "excel_sheet": "CBL",
        "individual_tableau_folder": "FOLDER",
        "individual_tableau_workbook": "Legal.twb",
        "individual_tableau_worksheet": "Monthly Forecast",
        "company_filter": "CBL",
        "validation_status": "confirmed",
    },
    {
        "workbook": "FY27 OPEX HR_Aug Forecast.xlsx",
        "excel_sheet": "ALIV",
        "validation_status": "pending",
    },
]


def _run(tmp_path, twb, monkeypatch, capture, workbook_sheets=("ALIV", "CBL"), **excel):
    folder = tmp_path / "forecast"
    for sheet in workbook_sheets:
        path = folder / f"{sheet}.xlsx"
        write_forecast_workbook(path, sheet=sheet, **excel)
    # One workbook with both sheets is what production uses; two files keep
    # the fixture simple, so point each route at its own file.
    routes = [dict(r) for r in ROUTES]
    routes[0]["workbook"] = "ALIV_Aug Forecast.xlsx"
    routes[1]["workbook"] = "CBL_Aug Forecast.xlsx"
    for sheet in workbook_sheets:
        (folder / f"{sheet}.xlsx").rename(folder / f"{sheet}_Oct Forecast.xlsx")
    source = twb(FORECAST_TYPE_TAB)
    for route in routes[:2]:
        route["individual_tableau_folder"] = str(source.parent)
    monkeypatch.setattr(vs, "_temp_twb_path", lambda src: tmp_path / f"tmp_{len(list(tmp_path.glob('tmp_*')))}.twb")
    opened, closed, progress = [], [], []
    result = vs.run_grouped_individual_validation(
        routes=routes,
        forecast_folder=folder,
        target_month_abbr="Oct",
        forecast_month=10,
        forecast_year=2026,
        forecast_type="FY27 Forecast_102026",
        opener=lambda tableau_workbook_path: opened.append(tableau_workbook_path),
        waiter=lambda tableau_workbook_path: None,
        capture=capture,
        closer=lambda path: closed.append(path) or {"closed": True, "warning": None},
        progress=lambda done, total, label: progress.append((done, total)),
    )
    return result, opened, closed, progress


def _grand_total(value):
    return [{"Company": "Grand Total", "Grand Total": value}]


def test_one_failing_route_does_not_stop_the_run(tmp_path, twb, monkeypatch):
    calls = []

    def capture(worksheet_name, tableau_workbook_path):
        calls.append(tableau_workbook_path)
        if len(calls) == 1:
            raise RuntimeError("Tableau workbook not ready yet")
        return _grand_total("300")

    result, opened, closed, progress = _run(tmp_path, twb, monkeypatch, capture)
    statuses = [r["status"] for r in result["results"]]
    assert statuses == ["ACTION REQUIRED", "PASS", "NOT VALIDATED"]
    assert "Tableau did not open in time" in result["results"][0]["note"]
    assert len(opened) == len(closed) == 2
    assert not list(tmp_path.glob("tmp_*.twb"))  # temp copies cleaned up
    assert progress[-1] == (3, 3)
    assert result["not_validated"] == 1


def test_wrong_month_workbook_marked_not_validated_and_run_continues(tmp_path, twb, monkeypatch):
    result, *_ = _run(
        tmp_path, twb, monkeypatch, lambda **k: _grand_total("300"),
        forecast_type="FY27 Forecast_092026",
    )
    first = result["results"][0]
    assert first["status"] == "NOT VALIDATED"
    assert "September 2026" in first["note"] and "October 2026" in first["note"]
    assert result["results"][1]["status"] == "NOT VALIDATED"


def test_variance_over_tolerance_fails(tmp_path, twb, monkeypatch):
    result, *_ = _run(tmp_path, twb, monkeypatch, lambda **k: _grand_total("310"))
    assert [r["status"] for r in result["results"][:2]] == ["FAIL", "FAIL"]
    assert result["results"][0]["variance"] == -10
