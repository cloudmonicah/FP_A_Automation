from datetime import date

from automations.actuals.revenue_validation_service import (
    parse_capture,
    source_values,
)

AUG = date(2026, 8, 31)


def _tsv(*rows):
    return "\n".join("\t".join(r) for r in rows)


def test_government_two_header_month_and_ytd():
    dims = ["GL Bus Code Name", "GL FS Category", "GL Account Mapping", "GL Sub Segment code", "GL Sub Account Description"]
    h1 = dims + ["Jul-26", "", "", "Aug-26", "", ""]
    h2 = [""] * 5 + ["Actual", "Budget", "Var $", "Actual", "Budget", "Var $"]
    body = [
        ["ALIV BUSINESS", "CORPORATE GOVERNMENT", "TV SERVICE REVENUE", "400005", "TV", "1,000", "1", "1", "2,000", "1", "1"],
        ["", "", "Total", "", "", "852,954", "9", "9", "900,001", "9", "9"],
    ]
    ytd_text = _tsv(h1, h2, *body)
    rows = parse_capture(ytd_text, "YTD", AUG)
    value = source_values("CBL Government", rows)
    assert value == {"CBL": {"Government": 852954.0 + 900001.0}}
    # MONTH basis: only the Aug columns are shown.
    month_text = _tsv(dims + ["Aug-26", "", ""], [""] * 5 + ["Actual", "Budget", "Var $"],
                      ["ALIV BUSINESS", "CORPORATE GOVERNMENT", "Total", "", "", "900,001", "9", "9"])
    rows = parse_capture(month_text, "MONTH", AUG)
    assert source_values("CBL Government", rows) == {"CBL": {"Government": 900001.0}}


def test_finance_view_month_columns_sum_for_ytd():
    dims = ["GL Bus Code Name", "GL FS Category", "GL Account Mapping", "GL Sub Segment code", "GL Sub Account Description"]
    text = _tsv(
        dims + ["Jul-26", "Aug-26"],
        ["HFC FIXED NETWORK CONSUMER", "OTHER SUBSCRIBER", "SMART HOME PRODUCTS REVENUE", "408001", "RING REVENUE", "2,090", "3,959"],
        ["INTERCOMPANY", "OTHER SUBSCRIBER", "MSA REVENUE", "420001", "X", "5", "6"],
        ["", "Total", "", "", "", "1,098,554", "1,097,582"],
    )
    rows = parse_capture(text, "YTD", AUG)
    result = source_values("CBL Finance View", rows)
    assert result["CBL"]["Other"] == 2090 + 3959
    assert result["CBL"]["MSA-ALIV"] == 1098554 + 1097582


def test_aliv_other_from_rollup_sums_months():
    text = _tsv(
        ["Company", "GL FS Revenue by Type", "GL Account Mapping", "Jul-26", "Aug-26"],
        ["ALIV", "OTHER REVENUE", "OTHER REVENUE", "11,208", "19,200"],
        ["", "", "Total", "11,208", "19,200"],
    )
    rows = parse_capture(text, "YTD", AUG)
    assert source_values("ALIV Group Other", rows) == {"ALIV": {"Other": 30408.0}}


def test_wrong_month_headers_are_rejected():
    text = _tsv(
        ["Company", "GL FS Revenue by Type", "GL Account Mapping", "Jul-26", "Aug-26", "Sep-26"],
        ["ALIV", "OTHER REVENUE", "Total", "1", "2", "3"],
    )
    try:
        parse_capture(text, "YTD", AUG)
    except ValueError as exc:
        assert "month headers" in str(exc)
    else:
        raise AssertionError("expected the extra month to be rejected")
