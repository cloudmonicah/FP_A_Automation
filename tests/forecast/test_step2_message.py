"""Step 2 pop-up: problems look like problems, housekeeping looks like information."""

from ui.forecast_messages import step2_message

FOLDER = r"C:\Users\x\OneDrive - Aliv\Shortcuts\FP&A\FY27\2026-10 - October\Forecast Validation"


def _ok(workbook, sheet, **extra):
    return {"workbook": workbook, "worksheet": sheet, "unmerged": [], "comments_dropped": [],
            "months_blank_in_tableau": [], "date_formats_fixed": 0, **extra}


def test_housekeeping_is_information_not_a_warning():
    attention, text = step2_message("OPEX", {
        "passed": [
            _ok("FY27 Legal BOARD EXPENSES_Oct Forecast.xlsx", "ALIV", months_blank_in_tableau=["Jul", "Aug"]),
            _ok("FY27 OPEX Engineering & Technology_Oct Forecast.xlsx", "Facilities- CBL",
                unmerged=["U41:U48"], date_formats_fixed=21),
        ],
        "failed": [], "workbooks_updated": 2, "validation_folder": FOLDER})
    assert not attention
    assert text.startswith("OPEX: all 2 sheets updated in 2 workbooks.")
    assert "For your information" in text and "Needs your attention" not in text
    assert "FY27 Legal BOARD EXPENSES / ALIV: Jul, Aug" in text           # short names
    assert "21 amount(s) were showing as dates" in text
    assert "Merged" not in text and "unmerged" not in text.lower()       # run log only
    assert text.rstrip().endswith("Next: click Validate under OPEX (step 3).")


def test_failures_are_numbered_under_needs_your_attention():
    attention, text = step2_message("Revenue", {
        "passed": [_ok("FY27 CONSUMER REVENUE_Oct Forecast.xlsx", "CBL")],
        "failed": [{"workbook": "FY27 OTHER REVENUE_Oct Forecast.xlsx", "worksheet": "ALIV",
                    "error": 'No tab called "ALIV OTHER REVENUE DETAILS (FINANCE) - Monthly Forecast".'}],
        "workbooks_updated": 1, "validation_folder": FOLDER})
    assert attention
    assert text.startswith("Revenue: 1 of 2 sheets updated. 1 sheet was not updated and is unchanged.")
    assert "Needs your attention\n1. FY27 OTHER REVENUE / ALIV\n    No tab called" in text
    assert "Next:" not in text


def test_retries_and_removed_comments_need_attention():
    attention, text = step2_message("Direct Cost", {
        "passed": [_ok("FY27 CBL DTH COSTS_Oct Forecast.xlsx", "CBL",
                       comments_dropped=[{"account": "AF / CBL / 610", "text": "Moved"}])],
        "failed": [], "check_manually": [_ok("FY27 CBL DTH COSTS_Oct Forecast.xlsx", "CBL")],
        "workbooks_updated": 1})
    assert attention
    assert "worked on the second try" in text and "comments were removed" in text
