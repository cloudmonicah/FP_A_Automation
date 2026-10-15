"""Forecast errors must reach the user in plain language, not the generic fallback."""

import pytest

from automations.forecast.validation_service import WrongPeriodError
from shared.friendly_errors import GENERIC, friendly

CASES = [
    (ValueError("July is the initial fiscal-year forecast month. Use Initial Forecast Month Setup to create July."),
     "Initial Budget Setup"),
    (FileNotFoundError(r"No valid FY27 forecast workbooks were found before Oct in C:\Forecast\Sep"),
     r"C:\Forecast\Sep"),
    (RuntimeError("Forecast file already exists: FY27 OPEX Legal_Oct Forecast.xlsx"),
     '"replace"'),
    (RuntimeError('Tableau parameter "Forecast Start" could not be located in FY27.twb.'),
     '"Forecast Start" parameter'),
    (RuntimeError('Expected 12 fiscal month columns in "ALIV", but found 0. Run step 2 (Update Workbooks from Tableau) for this workbook first.'),
     "Run step 2"),
    (KeyError('Excel worksheet "CBL" was not found. Available sheets: [\'ALIV\']'),
     'no sheet named "CBL"'),
    (RuntimeError('Excel sheet "CBL" not found. Available: [\'ALIV\']'),
     'no sheet named "CBL"'),
    (RuntimeError('Individual Tableau crosstab has no "Grand Total" row. Check that the worksheet shows column grand totals (Analysis > Totals).'),
     "Analysis > Totals"),
    (ValueError('Expected one Company filter on individual forecast tab "Monthly Forecast".'),
     '"Monthly Forecast" is not set up'),
    (FileNotFoundError("Forecast workbook not found:\nC:\\Out\\FY27 OPEX Legal_Oct Forecast.xlsx"),
     "FY27 OPEX Legal_Oct Forecast.xlsx is not in the Forecast Output Folder"),
    (RuntimeError('Tableau "Crosstab" menu item could not be found after 5 complete menu attempts.'),
     "did not respond"),
    (TimeoutError('Tableau did not finish opening "Legal" within 240 seconds. Check that Tableau Desktop started.'),
     'did not finish opening "Legal"'),
    (RuntimeError("Tableau extraction failed: RuntimeError: Tableau copied an empty crosstab."),
     "empty table"),
]


@pytest.mark.parametrize("error, expected", CASES, ids=[str(e)[:40] for e, _ in CASES])
def test_forecast_errors_are_explained(error, expected):
    message = friendly(error)
    assert message != GENERIC
    assert expected in message


def test_messages_written_for_users_are_shown_as_written():
    error = WrongPeriodError("Legal / ALIV is for September 2026, but the selected month is October 2026.")
    assert friendly(error) == str(error)
