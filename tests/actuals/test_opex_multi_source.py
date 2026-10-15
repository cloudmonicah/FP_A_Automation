from datetime import date
from pathlib import Path
import tempfile

from automations.actuals.individual_opex_crosstab_parser import extract_monthly_ytd
from automations.actuals.individual_actuals_validation_service import _combine_source_rows

AUG = date(2026, 8, 1)


def _write(rows):
    path = Path(tempfile.mkdtemp()) / "c.txt"
    path.write_text("\n".join("\t".join(r) for r in rows), encoding="utf-8")
    return path


def test_tab_without_company_column_uses_grand_total():
    periods = ["", "", "July 2026", "August 2026", "Grand Total"]
    measures = ["", "", "Actual", "Actual", "Actual"]
    rows = [
        periods,
        measures,
        ["PAYROLL", "SALARIES", "60", "40", "100"],
        ["Grand Total", "", "60", "40", "100"],
    ]
    out = extract_monthly_ytd(_write(rows), "ALIV", AUG)
    assert float(out["MONTH"]) == 40 and float(out["YTD"]) == 100


def test_two_workbooks_are_added_and_status_recomputed():
    def pair(month, ytd, fs_m, fs_y, cons_m, cons_y):
        mk = lambda basis, amount, fs, cons: [
            basis, fs, cons, amount, None if amount is None else round(amount / 1000), None, None,
            "ACTION REQUIRED", "wb", "tab", "sheet", "2026-09-01", "ft", "",
        ]
        return [mk("MONTH", month, fs_m, cons_m), mk("YTD", ytd, fs_y, cons_y)]

    a = pair(600_000, 1_100_000, 989, 1872, 989, 1872)
    b = pair(389_000, 772_000, 989, 1872, 989, 1872)
    month, ytd = _combine_source_rows([a, b])
    assert month[3] == 989_000 and month[4] == 989 and month[7] == "PASS"
    assert ytd[3] == 1_872_000 and ytd[7] == "PASS"
    bad = pair(None, None, 989, 1872, 989, 1872)
    bad[0][3] = bad[1][3] = None
    month, _ = _combine_source_rows([a, bad])
    assert month[7] == "ACTION REQUIRED"
