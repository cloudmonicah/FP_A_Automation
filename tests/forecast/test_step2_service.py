"""Step 2 service: uses Step 1's Tableau workbook, never half-updates a file, retries once."""

from datetime import date

import pytest
from openpyxl import Workbook

from automations.forecast import service
from shared.friendly_errors import PlainLanguageError

TWB = """<?xml version='1.0' encoding='utf-8' ?>
<workbook xmlns:user='http://www.tableausoftware.com/xml/user'>
  <datasources><datasource name='Parameters'>
    <column caption='Forecast Start' name='[Parameter 1]' param-domain-type='any' value='#{start}#' />
  </datasource></datasources>
  <worksheets><worksheet name='OPEX'><table><view>
    <filter><groupfilter function='member' member='&quot;{ftype}&quot;' /></filter>
  </view></table></worksheet></worksheets>
</workbook>
"""


def _twb(tmp_path, start="2026-10-01", ftype="FY27 Forecast_102026"):
    path = tmp_path / "FY Forecast.twb"
    path.write_text(TWB.format(start=start, ftype=ftype), encoding="utf-8")
    return path


def _period(start=date(2026, 10, 1)):
    return service.build_forecast_period("FY27", 10, 2026, forecast_start=start)


def test_matching_period_passes_and_the_twb_is_not_rewritten(tmp_path):
    twb = _twb(tmp_path)
    before = twb.read_bytes()
    service.check_consolidated_forecast_period(twb, _period())
    assert twb.read_bytes() == before


def test_changed_forecast_start_stops_step_2_with_a_clear_reason(tmp_path):
    twb = _twb(tmp_path)
    with pytest.raises(PlainLanguageError) as error:
        service.check_consolidated_forecast_period(twb, _period(date(2026, 9, 1)))
    text = str(error.value)
    assert "Forecast Start in Tableau is 2026-10-01, but this page has 2026-09-01" in text
    assert "Generate Forecast & Open Tableau" in text and "Nothing was updated" in text


def test_wrong_forecast_type_is_named(tmp_path):
    twb = _twb(tmp_path, ftype="FY27 Forecast_092026")
    with pytest.raises(PlainLanguageError, match="Forecast Type in Tableau is FY27 Forecast_092026"):
        service.check_consolidated_forecast_period(twb, _period())


def _target(tmp_path):
    path = tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx"
    Workbook().save(path)
    return path


def test_unexpected_failure_is_tried_again_and_flagged(tmp_path, monkeypatch):
    target = _target(tmp_path)
    calls = []

    def flaky(**kwargs):
        calls.append(kwargs["workbook_path"])
        if len(calls) == 1:
            raise OSError("disk hiccup")
        return {"tableau_detail_rows": 3}

    monkeypatch.setattr(service, "populate_forecast_sheet_from_audit", flaky)
    result, retried = service._paste_route_safely(target, "ALIV", tmp_path / "x.xlsx", "ALIV Legal", "FY27")
    assert retried and result["tableau_detail_rows"] == 3
    assert all(path.name.startswith("~updating") for path in calls)    # worked on a copy
    assert sorted(p.name for p in tmp_path.iterdir()) == [target.name]  # no leftovers


def test_a_failure_leaves_the_workbook_untouched(tmp_path, monkeypatch):
    target = _target(tmp_path)
    before = target.read_bytes()

    def broken(**kwargs):
        # Half-write the working copy, then crash.
        kwargs["workbook_path"].write_bytes(b"half written")
        raise ValueError("crashed mid-paste")

    monkeypatch.setattr(service, "populate_forecast_sheet_from_audit", broken)
    with pytest.raises(ValueError):
        service._paste_route_safely(target, "ALIV", tmp_path / "x.xlsx", "ALIV Legal", "FY27")
    assert target.read_bytes() == before
    assert sorted(p.name for p in tmp_path.iterdir()) == [target.name]


def test_problems_the_user_must_fix_are_not_retried(tmp_path, monkeypatch):
    target = _target(tmp_path)
    calls = []

    def missing_tab(**kwargs):
        calls.append(1)
        raise PlainLanguageError('no tab called "ALIV"')

    monkeypatch.setattr(service, "populate_forecast_sheet_from_audit", missing_tab)
    with pytest.raises(PlainLanguageError):
        service._paste_route_safely(target, "ALIV", tmp_path / "x.xlsx", "ALIV Legal", "FY27")
    assert calls == [1]


def test_workbook_open_in_excel_says_so(tmp_path, monkeypatch):
    target = _target(tmp_path)
    monkeypatch.setattr(service, "populate_forecast_sheet_from_audit", lambda **k: {})

    def locked(src, dst):
        raise PermissionError(13, "The process cannot access the file")

    monkeypatch.setattr(service.os, "replace", locked)
    with pytest.raises(PlainLanguageError, match="is open in Excel"):
        service._paste_route_safely(target, "ALIV", tmp_path / "x.xlsx", "ALIV Legal", "FY27")


def test_missing_workbook_says_to_run_step_1(tmp_path):
    with pytest.raises(PlainLanguageError, match="Run Step 1 first"):
        service._paste_route_safely(tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx",
                                    "ALIV", tmp_path / "x.xlsx", "ALIV Legal", "FY27")


def test_unknown_errors_still_give_a_specific_reason():
    reason = service._reason(KeyError("Jul-26"))
    assert "Something unexpected" not in reason and "Jul-26" in reason


TABS = """<?xml version='1.0' encoding='utf-8' ?>
<workbook>
  <worksheets>
    <worksheet name='ALIV Head Office' />
    <worksheet name='CBL  Head Office' />
    <worksheet name='ALIV Legal' />
  </worksheets>
  <windows>
    <window class='worksheet' name='ALIV Legal' />
    <window class='worksheet' hidden='true' name='ALIV Head Office' />
    <window class='dashboard' name='Overview' />
  </windows>
</workbook>
"""


def test_missing_or_hidden_tabs_are_named_precisely(tmp_path):
    twb = tmp_path / "FY27 Forecast.twb"
    twb.write_text(TABS, encoding="utf-8")
    problems = service.consolidated_tab_problems(
        twb, ["ALIV Legal", "cbl head office", "ALIV Head Office", "ALIV Head Offices", "Board"])
    assert set(problems) == {"ALIV Head Office", "ALIV Head Offices", "Board"}
    assert "is hidden in FY27 Forecast.twb" in problems["ALIV Head Office"] and "Unhide" in problems["ALIV Head Office"]
    assert 'no tab called "ALIV Head Offices". The closest tab names are: "ALIV Head Office"' in problems["ALIV Head Offices"]
    assert problems["Board"].endswith("tell support to change the mapping.")
