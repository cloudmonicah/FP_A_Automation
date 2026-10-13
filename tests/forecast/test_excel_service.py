from pathlib import Path

from automations.forecast.excel_service import create_forecast_refresh

# ============================================================
# TEST SETTINGS
# ============================================================

SOURCE_FOLDER = Path(r"C:\Users\moncloud\OneDrive - Aliv\Monicah\Forecast")

OUTPUT_FOLDER = Path(r"C:\Users\moncloud\OneDrive - Aliv\Monicah\ForecastTest")


# ============================================================
# RUN TEST
# ============================================================

result = create_forecast_refresh(
    source_folder=SOURCE_FOLDER,
    output_folder=OUTPUT_FOLDER,
    fiscal_year="FY27",
    forecast_month=10,
    forecast_year=2026,
    overwrite_mode="replace",
)


# ============================================================
# RESULTS
# ============================================================

print("\n" + "=" * 60)
print("MONTHLY FORECAST CREATION TEST")
print("=" * 60)

print(f"Workbooks found: " f"{result['unique_workbooks_found']}")

print(f"Files created: " f"{result['files_created']}")

print(f"Files skipped: " f"{result['files_skipped']}")

print(f"Files failed: " f"{result['files_failed']}")

print(f"Worksheets updated: " f"{result['worksheets_updated']}")

print(f"Forecast Type values updated: " f"{result['forecast_values_updated']}")

print(f"Expected source month: " f"{result['expected_source_month']}")

print(f"Target month: " f"{result['target_month']}")


print("\nFORMULA AUDIT")
print("-" * 60)

formula_failures = []

for audit in result["formula_audits"]:

    status = "PASS" if audit["formulas_preserved"] else "FAIL"

    print(
        f"{status}: "
        f"{audit['file']} "
        f"| Source formulas: "
        f"{audit['source_formula_count']} "
        f"| Output formulas: "
        f"{audit['destination_formula_count']}"
    )

    if not audit["formulas_preserved"]:
        formula_failures.append(audit)


print("\nOLDER SOURCE WORKBOOKS")
print("-" * 60)

if result["older_sources"]:

    for item in result["older_sources"]:

        print(
            f"WARNING: "
            f"{item['family']} "
            f"used {item['source_month']} "
            f"instead of "
            f"{item['expected_month']}"
        )

else:
    print("All workbooks came from the expected " "previous forecast month.")


print("\nFAILURES")
print("-" * 60)

if result["failures"]:

    for failure in result["failures"]:

        print(f"{failure['file']}: " f"{failure['error']}")

else:
    print("No workbook creation failures.")


print("\n" + "=" * 60)

if result["files_failed"] == 0 and not formula_failures:

    print("STAGE 1 RESULT: PASS")

else:

    print("STAGE 1 RESULT: REVIEW REQUIRED")

print("=" * 60)
