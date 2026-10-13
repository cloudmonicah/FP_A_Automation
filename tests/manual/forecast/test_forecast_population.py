from datetime import datetime
from pathlib import Path
import shutil
from openpyxl import load_workbook

from automations.forecast.excel_service import (
    FISCAL_MONTH_ORDER,
    _s4_structure,
    populate_forecast_sheet_from_audit,
)

FORECAST_FOLDER = Path(r"C:\Users\moncloud\OneDrive - Aliv\Monicah\ForecastTest")
SOURCE = FORECAST_FOLDER / "FY27 OPEX Legal_Oct Forecast.xlsx"
AUDIT = FORECAST_FOLDER / "Tableau_Forecast_Extract_2026-10.xlsx"
TEST = FORECAST_FOLDER / "FY27 OPEX Legal_Oct Forecast_STAGE4A_TEST.xlsx"

EXCEL_SHEET = "ALIV"
TABLEAU_SHEET = "ALIV Legal"
FISCAL_YEAR = "FY27"

EXPECTED = [
    datetime(2026, 7, 1),
    datetime(2026, 8, 1),
    datetime(2026, 9, 1),
    datetime(2026, 10, 1),
    datetime(2026, 11, 1),
    datetime(2026, 12, 1),
    datetime(2027, 1, 1),
    datetime(2027, 2, 1),
    datetime(2027, 3, 1),
    datetime(2027, 4, 1),
    datetime(2027, 5, 1),
    datetime(2027, 6, 1),
]


def main():
    if not SOURCE.exists():
        raise FileNotFoundError(SOURCE)
    if not AUDIT.exists():
        raise FileNotFoundError(AUDIT)

    shutil.copy2(SOURCE, TEST)

    result = populate_forecast_sheet_from_audit(
        workbook_path=TEST,
        excel_sheet_name=EXCEL_SHEET,
        audit_workbook_path=AUDIT,
        tableau_worksheet_name=TABLEAU_SHEET,
        fiscal_year=FISCAL_YEAR,
    )

    wb = load_workbook(TEST, data_only=False)
    try:
        ws = wb[EXCEL_SHEET]
        structure = _s4_structure(ws)
        header_row = structure["header_row"]

        actual = []
        for month in FISCAL_MONTH_ORDER:
            cell = ws.cell(header_row, structure["month_columns"][month])
            assert not (
                isinstance(cell.value, str) and cell.value.startswith("=")
            ), f"{cell.coordinate} is still a formula: {cell.value}"
            assert isinstance(
                cell.value, datetime
            ), f"{cell.coordinate} is not a real date: {cell.value!r}"
            assert (
                cell.number_format == "mmm-yy"
            ), f"{cell.coordinate} format is {cell.number_format!r}"
            actual.append(cell.value)

        assert actual == EXPECTED, f"Wrong fiscal date order: {actual}"
        assert result["tableau_detail_rows"] == result["excel_detail_rows"]

        total_row = result["total_row"]
        gt_col = structure["grand_total_column"]
        first = header_row + 1
        last = total_row - 1

        # No literal Grand Total label.
        for col in {
            structure["forecast_type_column"],
            *structure["descriptive_columns"].values(),
        }:
            assert (
                str(ws.cell(total_row, col).value or "").strip().lower()
                != "grand total"
            )

        # Detail Grand Total formulas.
        for row in range(first, last + 1):
            value = ws.cell(row, gt_col).value
            assert isinstance(value, str) and value.startswith(
                "="
            ), f"Missing detail total formula row {row}"

        # Bottom monthly formulas.
        for month in FISCAL_MONTH_ORDER:
            value = ws.cell(total_row, structure["month_columns"][month]).value
            assert isinstance(value, str) and value.startswith(
                "="
            ), f"Missing bottom formula for {month}"

        value = ws.cell(total_row, gt_col).value
        assert isinstance(value, str) and value.startswith(
            "="
        ), "Missing bottom overall Grand Total formula"

        # New border.
        july_col = structure["month_columns"][7]
        border = ws.cell(total_row, july_col).border
        assert border.top.style == "medium", border.top.style
        assert border.bottom.style == "double", border.bottom.style

        print("=" * 72)
        print("STAGE 4A - ALIV LEGAL")
        print("=" * 72)
        print(f"PASS - Tableau rows copied: {result['tableau_detail_rows']}")
        print("PASS - Jul-26 -> Jun-27 are REAL Excel dates")
        print("PASS - Dates are in correct fiscal order")
        print("PASS - No =EDATE/header formulas remain")
        print("PASS - Detail Grand Total formulas rebuilt")
        print("PASS - Bottom monthly totals rebuilt")
        print("PASS - Bottom overall Grand Total rebuilt")
        print('PASS - No literal "Grand Total" label')
        print("PASS - Medium top + double bottom total-row border")
        print()
        print("STAGE 4A RESULT: PASS")
        print(f"Review: {TEST}")
    finally:
        wb.close()


if __name__ == "__main__":
    main()
