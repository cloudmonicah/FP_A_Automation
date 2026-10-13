from pathlib import Path

from automations.forecast.mapping import FORECAST_ROUTES
from automations.forecast.validation_service import run_grouped_individual_validation

# ============================================================
# STAGE 5B CONTROLLED TEST
# ============================================================

FORECAST_FOLDER = Path(
    r"C:\Users\moncloud\OneDrive - Aliv\Monicah\ForecastTest\Stage4B_Full_Population_Test_2"
)

TARGET_MONTH_ABBR = "Oct"
FORECAST_MONTH = 10
FORECAST_YEAR = 2026

# Controlled set only. Do not expand this list until Stage 5B passes.
TEST_WORKBOOKS = [
    "Legal.twb",
    "Human Resources.twb",
    "Service Delivery.twb",
]


def main():
    print("=" * 78)
    print("STAGE 5B - GROUPED INDIVIDUAL TABLEAU VALIDATION")
    print("=" * 78)
    print(f"Forecast folder : {FORECAST_FOLDER}")
    print(f"Forecast period : {TARGET_MONTH_ABBR} {FORECAST_YEAR}")
    print("Individual TWBs :")
    for name in TEST_WORKBOOKS:
        print(f"  - {name}")
    print()

    if not FORECAST_FOLDER.exists():
        raise FileNotFoundError(f"Forecast test folder not found: {FORECAST_FOLDER}")

    result = run_grouped_individual_validation(
        routes=FORECAST_ROUTES,
        forecast_folder=FORECAST_FOLDER,
        target_month_abbr=TARGET_MONTH_ABBR,
        forecast_month=FORECAST_MONTH,
        forecast_year=FORECAST_YEAR,
        only_workbooks=[
            "Legal.twb",
            "Human Resources.twb",
        ],
        tolerance=5.00,
        open_wait_seconds=8.0,
    )

    print()
    print("=" * 78)
    print("WORKBOOK OPEN/CLOSE RESULTS")
    print("=" * 78)
    for item in result["workbooks"]:
        print(f"Workbook : {Path(item['source']).name}")
        print(f"Routes   : {item['routes']}")
        print(f"Opened   : {item['opened']}")
        print(f"Closed   : {item['closed']}")
        print(f"Status   : {item['status']}")
        if item["error"]:
            print(f"Message  : {item['error']}")
        print("-" * 78)

    print()
    print("=" * 78)
    print("ROUTE VALIDATION RESULTS")
    print("=" * 78)

    for item in result["results"]:
        print(f"Excel    : {item['forecast_workbook']}")
        print(f"Sheet    : {item['worksheet']}")
        print(f"Tableau  : {item['individual_tableau_workbook']}")
        print(f"Tab      : {item['individual_tableau_worksheet']}")

        if item["excel_grand_total"] is not None:
            print(f"Excel GT : {item['excel_grand_total']:,.2f}")
        else:
            print("Excel GT : --")

        if item["tableau_grand_total"] is not None:
            print(f"Tab GT   : {item['tableau_grand_total']:,.2f}")
        else:
            print("Tab GT   : --")

        if item["variance"] is not None:
            print(f"Variance : {item['variance']:+,.2f}")
        else:
            print("Variance : --")

        print(f"Status   : {item['status']}")
        if item.get("note"):
            print(f"Note     : {item['note']}")
        print("-" * 78)

    print()
    print("=" * 78)
    print("STAGE 5B SUMMARY")
    print("=" * 78)
    print(f"Passed          : {result['passed']}")
    print(f"Failed          : {result['failed']}")
    print(f"Action Required : {result['action_required']}")

    workbook_ok = all(
        item["opened"] and item["closed"] and item["status"] == "PASS"
        for item in result["workbooks"]
    )
    route_ok = result["failed"] == 0 and result["action_required"] == 0

    print()
    print("STAGE 5B RESULT:", "PASS" if workbook_ok and route_ok else "REVIEW REQUIRED")


if __name__ == "__main__":
    main()
