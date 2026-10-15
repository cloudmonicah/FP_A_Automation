from automations.forecast import tableau_service as ts


def test_a_worksheet_that_fails_once_is_retried(tmp_path, monkeypatch):
    workbook = tmp_path / "FY27.twb"
    workbook.write_text("<workbook/>")
    calls = []

    def capture(worksheet_name, **kwargs):
        calls.append(worksheet_name)
        if worksheet_name == "Big" and calls.count("Big") == 1:
            raise RuntimeError('Tableau "Crosstab" menu item could not be found')
        return [{"GL Account": "1"}]

    monkeypatch.setattr(ts, "get_tableau_crosstab_rows", capture)
    monkeypatch.setattr(ts, "_clear_clipboard", lambda: None)
    monkeypatch.setattr(ts, "_get_tableau_window", lambda **k: (_ for _ in ()).throw(RuntimeError("no UI here")))
    monkeypatch.setattr(ts.time, "sleep", lambda s: None)

    results, errors = ts.extract_forecast_worksheets(workbook, ["Small", "Big", "Never"])
    assert calls == ["Small", "Big", "Big", "Never"]
    assert set(results) == {"Small", "Big", "Never"} and errors == {}


def test_a_worksheet_that_keeps_failing_is_reported(tmp_path, monkeypatch):
    workbook = tmp_path / "FY27.twb"
    workbook.write_text("<workbook/>")

    def capture(worksheet_name, **kwargs):
        raise RuntimeError("menu not found")

    monkeypatch.setattr(ts, "get_tableau_crosstab_rows", capture)
    monkeypatch.setattr(ts, "_clear_clipboard", lambda: None)
    monkeypatch.setattr(ts, "_get_tableau_window", lambda **k: (_ for _ in ()).throw(RuntimeError("no UI here")))
    monkeypatch.setattr(ts.time, "sleep", lambda s: None)

    results, errors = ts.extract_forecast_worksheets(workbook, ["Big"])
    assert results == {} and "menu not found" in errors["Big"]
