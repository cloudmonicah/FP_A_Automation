from datetime import date

from openpyxl import Workbook

from automations.actuals.direct_cost_individual_merge import _extract

AUG = date(2026, 8, 1)


def _sheet(rows):
    wb = Workbook()
    ws = wb.active
    for r in rows:
        ws.append(r)
    return ws


def test_reseller_total_row_is_read_directly():
    ws = _sheet([
        ["", "", "July 2026", "August 2026", "Grand Total"],
        ["Company", "GL Account Mapping", "Actual", "Actual", "Actual"],
        ["ALIV", "COMMISSIONS - KIOSKS", 100000, 200000, 300000],
        ["ALIV", "COMMISSIONS - PUSH TO TALK", 266000, 152000, 418000],
        ["Grand Total", "", 366000, 352000, 718000],
    ])
    route = {"company": "ALIV", "fs_line": "Reseller Commissions", "aggregation": "TOTAL",
             "dedicated_total": True, "tableau_categories": []}
    month, ytd, _ = _extract(ws, route, AUG)
    assert (month, ytd) == (352, 718)
