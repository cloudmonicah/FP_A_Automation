from datetime import date

from openpyxl import Workbook

from automations.actuals.direct_cost_individual_merge import _extract


def _crosstab():
    wb = Workbook()
    ws = wb.active
    ws.append([None, None, "August 2026", "Grand Total"])
    ws.append(["Company", "GL Account Mapping", "Actual", "Actual"])
    ws.append(["ALIV", "INTERCONNECT COST", 1200, 2500])
    ws.append(["ALIV", "INTERCONNECT COST", 300, 500])
    ws.append(["CBL", "INTERCONNECT COST", 999, 999])
    ws.append(["ALIV", "Grand Total", 9999, 9999])
    return ws


def test_extract_sums_matching_rows_in_thousands():
    route = {"company": "ALIV", "fs_line": "Interconnect", "tableau_categories": ["INTERCONNECT COST"]}
    monthly, ytd, note = _extract(_crosstab(), route, date(2026, 8, 31))
    assert (monthly, ytd) == (2, 3)  # 1,500 and 3,000 dollars, rounded to $000
    assert "Summed 2" in note


def test_extract_reports_missing_category():
    route = {"company": "ALIV", "fs_line": "Media", "tableau_categories": ["MEDIA"]}
    monthly, ytd, note = _extract(_crosstab(), route, date(2026, 8, 31))
    assert monthly is None and ytd is None
    assert "No exact matching" in note
