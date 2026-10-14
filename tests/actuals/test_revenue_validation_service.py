from automations.actuals.revenue_validation_service import source_values


def test_cbl_group_routes_use_exported_dimensions():
    rows = [
        {
            "GL Bus Code Name": "ALIV BUSINESS",
            "GL FS Category": "CORPORATE GOVERNMENT",
            "_actual": "852954",
        },
        {
            "GL Bus Code Name": "HFC FIXED NETWORK CONSUMER",
            "GL FS Category": "OTHER SUBSCRIBER",
            "_actual": "3959",
        },
        {
            "GL Bus Code Name": "INTERCOMPANY",
            "GL FS Category": "OTHER SUBSCRIBER",
            "_actual": "1097582",
        },
        {
            "GL Bus Code Name": "MEDIA",
            "GL FS Category": "CORPORATE POST PAID RECURRING",
            "_actual": "112164",
        },
        {
            "GL Bus Code Name": "Grand Total",
            "GL FS Category": "Total",
            "_actual": "12107614",
        },
    ]

    assert source_values("CBL Group", rows) == {
        "CBL": {
            "Government": 852954.0,
            "Other": 3959.0,
            "MSA-ALIV": 1097582.0,
            "Media": 112164.0,
            "Gross Revenue": 12107614.0,
        }
    }
