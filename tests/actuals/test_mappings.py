from collections import Counter
from pathlib import PureWindowsPath

import pytest

from automations.actuals import mapping
from automations.forecast.mapping import FORECAST_ROUTES
from shared import settings

ACTUALS_TABLES = {
    "OPEX_MAPPINGS": mapping.OPEX_MAPPINGS,
    "DIRECT_COST_MAPPINGS": mapping.DIRECT_COST_MAPPINGS,
    "REVENUE_MAPPINGS": mapping.REVENUE_MAPPINGS,
    "INDIVIDUAL_DIRECT_COST_MAPPINGS": mapping.INDIVIDUAL_DIRECT_COST_MAPPINGS,
    "INDIVIDUAL_REVENUE_MAPPINGS": mapping.INDIVIDUAL_REVENUE_MAPPINGS,
}


@pytest.mark.parametrize("name", ACTUALS_TABLES)
def test_actuals_mapping_rows_are_complete_and_unique(name):
    rows = ACTUALS_TABLES[name]
    assert rows
    for row in rows:
        assert row["company"] in ("ALIV", "CBL"), row
        assert str(row["fs_line"]).strip(), row
    counts = Counter((r["company"], r["fs_line"]) for r in rows)
    assert [k for k, n in counts.items() if n > 1] == []


@pytest.mark.parametrize(
    "rows", [mapping.OPEX_MAPPINGS, mapping.INDIVIDUAL_DIRECT_COST_MAPPINGS, mapping.INDIVIDUAL_REVENUE_MAPPINGS]
)
def test_workbooks_live_under_the_configured_root(rows):
    root = PureWindowsPath(settings.get("workbooks_root").replace("/", "\\"))
    for row in rows:
        path = row.get("individual_tableau_workbook")
        if path:
            assert PureWindowsPath(path.replace("/", "\\")).is_relative_to(root), path
            assert path.lower().endswith(".twb"), path


def test_forecast_routes_are_complete():
    required = {"tableau_worksheet", "workbook", "excel_sheet"}
    for route in FORECAST_ROUTES:
        assert required <= route.keys(), route
        assert route["workbook"].endswith(".xlsx"), route
