from pathlib import Path

from automations.forecast.mapping import (
    FORECAST_ROUTES,
)

from automations.forecast.tableau_service import (
    get_tableau_crosstab_rows,
)

from automations.forecast.validation_service import (
    validate_route_from_tableau_rows,
)

FORECAST_FOLDER = Path(r"C:\Users\moncloud\OneDrive - Aliv\Monicah\ForecastTest")

LEGAL_TABLEAU = Path(r"E:\Workbooks\OPEX\Legal.twb")

TARGET_MONTH_ABBR = "Oct"


def get_legal_routes():
    return [
        route
        for route in FORECAST_ROUTES
        if (
            route.get("individual_tableau_workbook") == "Legal.twb"
            and route.get("validation_status") == "confirmed"
        )
    ]


def money(value):
    if value is None:
        return ""

    return f"{value:,.2f}"


def main():
    print("=" * 80)
    print("STAGE 5A - LEGAL GRAND TOTAL VALIDATION")
    print("=" * 80)

    routes = get_legal_routes()

    if len(routes) != 2:
        raise RuntimeError(
            "Expected exactly 2 confirmed " f"Legal routes; found {len(routes)}."
        )

    results = []

    for position, route in enumerate(
        routes,
        start=1,
    ):
        print()
        print("=" * 80)
        print(f"[{position}/{len(routes)}] " f"{route['excel_sheet']}")
        print("=" * 80)

        rows = get_tableau_crosstab_rows(
            worksheet_name=route["individual_tableau_worksheet"],
            tableau_workbook_path=(LEGAL_TABLEAU),
        )

        result = validate_route_from_tableau_rows(
            route=route,
            forecast_folder=(FORECAST_FOLDER),
            target_month_abbr=(TARGET_MONTH_ABBR),
            tableau_rows=rows,
        )

        results.append(result)

        print()
        print("Excel Grand Total   : " f"{money(result['excel_grand_total'])}")

        print("Tableau Grand Total : " f"{money(result['tableau_grand_total'])}")

        print("Variance            : " f"{money(result['variance'])}")

        print("STATUS              : " f"{result['status']}")

    print()
    print("=" * 80)
    print("STAGE 5A RESULTS")
    print("=" * 80)

    passed = 0
    failed = 0

    for result in results:
        print(
            f"{result['worksheet']:<10} | "
            f"Excel: "
            f"{money(result['excel_grand_total']):>15} | "
            f"Tableau: "
            f"{money(result['tableau_grand_total']):>15} | "
            f"Variance: "
            f"{money(result['variance']):>12} | "
            f"{result['status']}"
        )

        if result["status"] == "PASS":
            passed += 1
        else:
            failed += 1

    print()
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")

    if failed:
        print()
        print("STAGE 5A RESULT: FAIL")
    else:
        print()
        print("STAGE 5A RESULT: PASS")


if __name__ == "__main__":
    main()
