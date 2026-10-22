"""Step 1: setting Forecast Start / Forecast Type in the consolidated .twb."""

from datetime import date

import pytest

from automations.forecast import service, tableau_service

TWB = """<?xml version='1.0' encoding='utf-8' ?>
<workbook source-build='2024.2' version='18.1' xmlns:user='http://www.tableausoftware.com/xml/user'>
  <datasources>
    <datasource hasconnection='false' inline='true' name='Parameters'>
      <column caption='Forecast Start' datatype='date' name='[Parameter 1]' param-domain-type='any' role='measure' type='quantitative' value='#2026-08-01#'>
        <calculation class='tableau' formula='#2026-08-01#' />
      </column>
    </datasource>
  </datasources>
  <worksheets>
    <worksheet name='OPEX'>
      <table><view><filter class='categorical' column='[x].[Forecast Type]'>
        <groupfilter function='member' level='[Forecast Type]' member='&quot;FY27 Forecast_082026&quot;' user:ui-domain='database' />
      </filter></view></table>
    </worksheet>
  </worksheets>
</workbook>
"""


@pytest.fixture
def twb(tmp_path, monkeypatch):
    monkeypatch.setattr("shared.tableau_temp_copies.tableau_temp_folder", lambda: tmp_path)
    path = tmp_path / "FY Forecast.twb"
    path.write_text(TWB, encoding="utf-8")
    return path


def test_sets_forecast_start_and_leaves_the_forecast_type_filter_alone(twb):
    period = service.set_consolidated_forecast_period(twb, "FY27", 10, 2026, date(2026, 9, 1))
    text = twb.read_text(encoding="utf-8")
    assert "value=\"#2026-09-01#\"" in text
    # The user picks Forecast Type in Tableau; the saved filter is untouched.
    assert "FY27 Forecast_082026" in text and "102026" not in text
    assert "user:ui-domain" in text and "ns0:" not in text
    assert period["forecast_type"] == "FY27 Forecast_102026"   # still used for Excel
    assert period["forecast_type_updates"] == 0


def test_a_backup_is_kept_once(twb, tmp_path):
    service.set_consolidated_forecast_period(twb, "FY27", 10, 2026)
    backup = tmp_path / "FY Forecast_BEFORE_PARAMETER_AUTOMATION.twb"
    assert "#2026-08-01#" in backup.read_text(encoding="utf-8")
    service.set_consolidated_forecast_period(twb, "FY27", 11, 2026)
    assert "#2026-08-01#" in backup.read_text(encoding="utf-8")


def test_missing_parameter_is_reported(twb):
    twb.write_text(TWB.replace("Forecast Start", "Something Else"), encoding="utf-8")
    with pytest.raises(RuntimeError, match="Forecast Start"):
        tableau_service.set_forecast_start_parameter(twb, 10, 2026)


def test_raw_copy_names_stay_inside_the_windows_path_limit(tmp_path):
    folder = tmp_path / ("x" * (175 - len(str(tmp_path))))
    long_tab = "ALIV Direct Cost (CONSUMER SOLUTIONS) - ROAMING_INTERCONNECT_OTHER Forecast"
    path = tableau_service.raw_copy_path(folder, long_tab, 7)
    assert len(str(path)) <= 250
    assert path.name.startswith("007_ALIV_Direct_Cost") and path.suffix == ".txt"
    other = tableau_service.raw_copy_path(folder, long_tab.replace("ALIV", "CBL"), 7)
    assert path.name[:-20] != other.name[:-20]   # kept apart (timestamps aside)
