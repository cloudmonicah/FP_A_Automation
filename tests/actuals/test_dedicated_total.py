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


def test_media_is_added_to_programming_fees():
    import tempfile
    from pathlib import Path
    from openpyxl import load_workbook
    from automations.actuals import direct_cost_individual_merge as m

    route_pf = {"company": "CBL", "fs_line": "Programming fees"}
    media = _sheet([
        ["", "", "July 2026", "August 2026", "Grand Total"],
        ["Company", "GL Account Mapping", "Actual", "Actual", "Actual YTD"],
        ["CBL", "Total", 100000, 62000, 162000],
        ["Grand Total", "", 100000, 62000, 162000],
    ])
    month, ytd, _ = m._extract_company_total(media, {"company": "CBL", "fs_line": "Programming fees - Media"}, AUG)
    assert (month, ytd) == (62, 162), (month, ytd)


def test_media_content_fees_single_row_view_is_added():
    from automations.actuals import direct_cost_individual_merge as m

    media = _sheet([
        ["", "", "July 2026", "August 2026", "Grand Total"],
        ["", "", "Actual", "Actual", "Actual"],
        ["", "", 70000, 62000, 132000],
    ])
    route = {"company": "CBL", "fs_line": "Programming fees - Media", "add_to": "Programming fees",
             "tableau_categories": []}
    month, ytd, _ = m._extract_company_total(media, route, AUG)
    assert (month, ytd) == (62, 132), (month, ytd)
