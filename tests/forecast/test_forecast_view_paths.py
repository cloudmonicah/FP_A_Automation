import pytest

pytest.importorskip("ttkbootstrap")

from shared import settings  # noqa: E402


@pytest.fixture
def isolated(tmp_path, monkeypatch):
    app = tmp_path / "app"
    app.mkdir()
    monkeypatch.setattr(settings, "app_dir", lambda: app)
    monkeypatch.setenv("APPDATA", str(tmp_path / "appdata"))
    settings.load(refresh=True)
    yield
    settings.load(refresh=True)


def test_paths_are_remembered_for_next_time(isolated, tmp_path):
    from ui.forecast_refresh_view import ForecastRefreshView as View

    assert View._remembered_paths() == {
        "forecast_source_folder": "",
        "forecast_output_folder": "",
        "forecast_tableau_workbook": "",
    }
    View._remember_paths(
        {"source_folder": str(tmp_path / "Sep"), "output_folder": str(tmp_path / "Oct"), "tableau_workbook": ""}
    )
    remembered = View._remembered_paths()
    assert remembered["forecast_source_folder"].endswith("Sep")
    assert remembered["forecast_output_folder"].endswith("Oct")
    assert remembered["forecast_tableau_workbook"] == ""
