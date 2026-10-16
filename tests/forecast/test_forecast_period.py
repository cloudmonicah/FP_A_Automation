from datetime import date

from automations.forecast import service


def test_forecast_start_is_the_first_of_the_forecast_month():
    october = service.build_forecast_period("FY27", 10, 2026)
    assert october["forecast_start"] == date(2026, 10, 1)
    assert october["forecast_start_tableau"] == "#2026-10-01#"
    assert october["forecast_type"] == "FY27 Forecast_102026"   # unchanged
    assert october["display"] == "October 2026"

    january = service.build_forecast_period("FY27", 1, 2027)
    assert january["forecast_start"] == date(2027, 1, 1)


def test_step_2_opens_tableau_only_when_it_is_not_open(monkeypatch):
    import automations.forecast.tableau_service as ts

    opened, waited = [], []
    monkeypatch.setattr(service, "open_tableau_workbook", lambda tableau_workbook_path: opened.append(tableau_workbook_path))
    monkeypatch.setattr(ts, "wait_for_tableau_workbook", lambda path: waited.append(path))

    monkeypatch.setattr(ts, "_get_tableau_window", lambda **k: "window")
    service._make_sure_tableau_is_open("FY27.twb")
    assert opened == [] and waited == []

    def not_open(**k):
        raise RuntimeError("Tableau Desktop did not become available")

    monkeypatch.setattr(ts, "_get_tableau_window", not_open)
    service._make_sure_tableau_is_open("FY27.twb")
    assert opened == ["FY27.twb"] and waited == ["FY27.twb"]
