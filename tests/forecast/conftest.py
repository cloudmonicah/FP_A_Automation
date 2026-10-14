"""Small synthetic Tableau (.twb) and forecast Excel files for the forecast tests."""

from datetime import datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

FORECAST_TYPE_TAB = """<?xml version='1.0' encoding='utf-8' ?>
<workbook>
  <datasources>
    <datasource name='Parameters'>
      <column caption='Forecast Start' name='[Parameter 1]' param-domain-type='range' value='#2026-08-01#'>
        <calculation class='tableau' formula='#2026-08-01#' />
      </column>
    </datasource>
  </datasources>
  <worksheets>
    <worksheet name='Monthly Forecast'>
      <table><view>
        <datasource-dependencies datasource='Parameters'>
          <column caption='Forecast Start' name='[Parameter 1]' />
        </datasource-dependencies>
        <filter class='categorical' column='[ds].[none:Company:nk]'>
          <groupfilter function='member' level='[none:Company:nk]' member='"ALIV"' />
        </filter>
        <filter class='categorical' column='[ds].[none:Forecast Type:nk]'>
          <groupfilter function='union'>
            <groupfilter function='member' level='[none:Forecast Type:nk]' member='"FY27 Forecast_082026"' />
          </groupfilter>
        </filter>
      </view></table>
    </worksheet>
  </worksheets>
</workbook>
"""

FORECAST_MEASURE_TAB = """<?xml version='1.0' encoding='utf-8' ?>
<workbook>
  <datasources>
    <datasource name='Parameters'>
      <column caption='Forecast Start' name='[Parameter 1]' param-domain-type='range' value='#2026-08-01#' />
      <column caption='Forecast Measure' name='[Parameter 2]' param-domain-type='list' value='"Budget"'>
        <calculation class='tableau' formula='"Budget"' />
      </column>
    </datasource>
  </datasources>
  <worksheets>
    <worksheet name='Monthly Forecast'>
      <table><view>
        <datasource-dependencies datasource='Parameters'>
          <column caption='Forecast Start' name='[Parameter 1]' />
          <column caption='Forecast Measure' name='[Parameter 2]'>
            <calculation class='tableau' formula='IF [Parameter 2] = "Forecast" THEN [Forecast] END' />
          </column>
        </datasource-dependencies>
        <filter class='categorical' column='[ds].[none:Company:nk]'>
          <groupfilter function='member' level='[none:Company:nk]' member='"CBL"' />
        </filter>
      </view></table>
    </worksheet>
  </worksheets>
</workbook>
"""


@pytest.fixture
def twb(tmp_path):
    def make(text: str, name: str = "Legal.twb") -> Path:
        path = tmp_path / "tableau" / name
        path.parent.mkdir(exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    return make


def write_forecast_workbook(
    path: Path,
    sheet: str = "ALIV",
    forecast_type: str = "FY27 Forecast_102026",
    rows=((100, 200),),
    extra_columns: int = 0,
) -> Path:
    """A forecast sheet in the Stage 4 layout: header row, Jul-Jun dates, detail rows.

    ``extra_columns`` inserts blank columns before the block, so tests can prove
    nothing depends on fixed column positions.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(["Monthly forecast"])
    ws.append([])
    offset = [None] * extra_columns
    months = [datetime(2026 if m >= 7 else 2027, m, 1) for m in (7, 8, 9, 10, 11, 12, 1, 2, 3, 4, 5, 6)]
    ws.append(offset + ["Forecast Type", "GL Account", *months, "Grand Total"])
    for values in rows:
        monthly = list(values) + [0] * (12 - len(values))
        ws.append(offset + [forecast_type, "6000", *monthly, "=SUM()"])
    ws.append(offset + [None, None, *(["=SUM()"] * 12), "=SUM()"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
