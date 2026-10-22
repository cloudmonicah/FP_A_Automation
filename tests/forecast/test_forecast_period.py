from datetime import date

from automations.forecast import service


def test_october_uses_the_october_forecast_type_and_starts_10_1():
    october = service.build_forecast_period("FY27", 10, 2026)
    assert october["forecast_start"] == date(2026, 10, 1)
    assert october["forecast_start_tableau"] == "#2026-10-01#"
    assert october["forecast_type"] == "FY27 Forecast_102026"
    assert october["display"] == "October 2026"

    january = service.build_forecast_period("FY27", 1, 2027)
    assert january["forecast_start"] == date(2027, 1, 1)
    assert january["forecast_type"] == "FY27 Forecast_012027"


def test_forecast_start_can_be_entered():
    october = service.build_forecast_period("FY27", 10, 2026, forecast_start=date(2026, 9, 1))
    assert october["forecast_start_tableau"] == "#2026-09-01#"
    # An edited Forecast Start does not change Forecast Type.
    assert october["forecast_type"] == "FY27 Forecast_102026"


def test_forecast_start_must_be_the_first_of_a_month():
    import pytest

    from shared.friendly_errors import PlainLanguageError

    with pytest.raises(PlainLanguageError, match="1st of a month"):
        service.build_forecast_period("FY27", 10, 2026, forecast_start=date(2026, 9, 15))


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


def test_every_route_has_a_category():
    """A route without one stops the Forecast page from opening."""
    from automations.forecast.mapping import FORECAST_CATEGORIES, FORECAST_ROUTES

    missing = [r["tableau_worksheet"] for r in FORECAST_ROUTES if r.get("category") not in FORECAST_CATEGORIES]
    assert missing == []
