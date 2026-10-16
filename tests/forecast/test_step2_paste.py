"""Step 2 pastes each consolidated Tableau tab into its sheet and keeps every formula."""

from datetime import datetime

from openpyxl import Workbook, load_workbook

from automations.forecast.excel_service import populate_forecast_sheet_from_tableau_rows

OCT = "FY27 Forecast_102026"


def _sheet(path, lines):
    """A forecast sheet like the real ones: =EDATE month headers, formula forecast months."""
    wb = Workbook()
    ws = wb.active
    ws.title = "ALIV"
    ws.append(["Legal forecast"])
    ws.append([
        "Forecast Type", "GL Sub Segment code", "GL Sub Account Description",
        datetime(2026, 7, 1), *[f"=EDATE(D2,{i})" for i in range(1, 12)], "Grand Total",
    ])
    for i, (code, desc) in enumerate(lines, start=3):
        # Jul-Sep values; Oct-Jun formulas (the department's forecast); row total.
        ws.append([OCT, code, desc, 1, 1, 1, *[f"=AVERAGE(D{i}:F{i})"] * 9, f"=SUM(D{i}:O{i})"])
    end = 2 + len(lines)
    ws.append([None, None, None, *[f"=SUM({c}3:{c}{end})" for c in "DEFGHIJKLMNO"], f"=SUM(P3:P{end})"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


LATER = ["Oct-26", "Nov-26", "Dec-26", "Jan-27", "Feb-27", "Mar-27", "Apr-27", "May-27", "Jun-27"]


def _row(code, desc, jul, aug, sep, later="9"):
    """A line as Tableau copies it, with all 12 fiscal months."""
    return {"GL Business Code": "LEGAL", "GL Sub Segment code": code, "GL Sub Account Description": desc,
            "Jul-26": jul, "Aug-26": aug, "Sep-26": sep, **{m: later for m in LATER}, "Grand Total": "999"}


def test_tableau_values_are_pasted_and_formulas_stay(tmp_path):
    path = _sheet(tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx", [(6000, "Salaries"), (6100, "Travel")])
    rows = [_row("6000", "Salaries", "1,000", "2,000", "(300)"), _row("6100", "Travel", "5", "6", "7")]
    populate_forecast_sheet_from_tableau_rows(path, "ALIV", rows, "FY27")

    ws = load_workbook(path)["ALIV"]
    assert [ws.cell(3, c).value for c in (4, 5, 6)] == [1000, 2000, -300]   # pasted as Tableau shows
    assert [ws.cell(3, c).value for c in range(7, 16)] == [9] * 9            # all 12 periods present
    assert ws.cell(3, 16).value == "=SUM(D3:O3)"                           # row total formula kept
    assert ws.cell(2, 5).value == "=EDATE(D2,1)"                           # header formulas kept
    assert ws.cell(5, 4).value == "=SUM(D3:D4)"                            # bottom total


def test_extra_tableau_lines_get_the_formulas_filled_down(tmp_path):
    path = _sheet(tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx", [(6000, "Salaries")])
    rows = [_row("6000", "Salaries", "1", "2", "3"), _row("6200", "Training", "4", "5", "6")]
    populate_forecast_sheet_from_tableau_rows(path, "ALIV", rows, "FY27")

    ws = load_workbook(path)["ALIV"]
    assert [ws.cell(4, c).value for c in (1, 2, 3, 4, 5, 6)] == [OCT, 6200, "Training", 4, 5, 6]
    assert [ws.cell(4, c).value for c in range(7, 16)] == [9] * 9
    assert ws.cell(4, 16).value == "=SUM(D4:O4)"
    assert ws.cell(5, 4).value == "=SUM(D3:D4)"       # total now covers the added line
    assert ws.cell(5, 16).value == "=SUM(P3:P4)"


def test_fewer_tableau_lines_remove_rows(tmp_path):
    path = _sheet(tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx", [(6000, "Salaries"), (6100, "Travel"), (6200, "x")])
    populate_forecast_sheet_from_tableau_rows(path, "ALIV", [_row("6000", "Salaries", "1", "2", "3")], "FY27")

    ws = load_workbook(path)["ALIV"]
    assert ws.cell(3, 2).value == 6000 and ws.cell(4, 1).value is None
    assert ws.cell(4, 4).value == "=SUM(D3:D3)"


def test_revenue_keeps_the_existing_sign_rule(tmp_path):
    path = _sheet(tmp_path / "FY27 CONSUMER REVENUE_Oct Forecast.xlsx", [(4000, "Sales")])
    populate_forecast_sheet_from_tableau_rows(path, "ALIV", [_row("4000", "Sales", "-10", "-20", "-30")], "FY27")
    ws = load_workbook(path)["ALIV"]
    assert [ws.cell(3, c).value for c in (4, 5, 6)] == [10, 20, 30]


def test_month_from_another_fiscal_year_is_not_pasted(tmp_path):
    path = _sheet(tmp_path / "FY27 OPEX Legal_Oct Forecast.xlsx", [(6000, "Salaries")])
    row = _row("6000", "Salaries", "1", "2", "3")
    row["Jul-27"] = "777"  # FY28 - must not land in Jul 2026
    populate_forecast_sheet_from_tableau_rows(path, "ALIV", [row], "FY27")
    assert load_workbook(path)["ALIV"].cell(3, 4).value == 1


def test_months_the_tableau_tab_leaves_out_are_pasted_blank(tmp_path):
    # Live run: "INTERNET BANDWIDTH (ALIV BUSINESS)" shows only Oct-26..Jun-27.
    path = _sheet(tmp_path / "FY27 CBL BANDWIDTH_Oct Forecast.xlsx", [(6000, "Bandwidth")])
    row = {k: v for k, v in _row("6000", "Bandwidth", "1", "2", "3").items() if k not in ("Jul-26", "Aug-26", "Sep-26")}
    result = populate_forecast_sheet_from_tableau_rows(path, "ALIV", [row], "FY27")
    ws = load_workbook(path)["ALIV"]
    assert [ws.cell(3, c).value for c in range(4, 16)] == [None, None, None] + [9] * 9
    assert result["months_blank_in_tableau"] == ["Jul", "Aug", "Sep"]
