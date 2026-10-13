from pathlib import Path
import xml.etree.ElementTree as ET

from automations.actuals.mapping import INDIVIDUAL_DIRECT_COST_MAPPINGS

passed = 0
failed = 0

for route in INDIVIDUAL_DIRECT_COST_MAPPINGS:
    company = route["company"]
    fs_line = route["fs_line"]
    path = Path(route["individual_tableau_workbook"])
    worksheet = route.get("individual_tableau_worksheet")

    print(f"\n{company} | {fs_line}")
    print("Workbook:", path.name)
    print("Worksheet:", worksheet or "NOT CONFIGURED")

    if not path.is_file():
        print("FAIL: Workbook missing")
        failed += 1
        continue

    try:
        root = ET.parse(path).getroot()
        names = {ws.get("name") for ws in root.findall("./worksheets/worksheet")}

        if not worksheet:
            print("REVIEW: Worksheet name required")
            failed += 1
        elif worksheet not in names:
            print("FAIL: Worksheet not found")
            print("Available worksheets:")
            for name in sorted(names):
                print(" -", name)
            failed += 1
        else:
            print("PASS: Workbook and worksheet verified")
            passed += 1

    except Exception as exc:
        print("FAIL:", exc)
        failed += 1

print("\n" + "=" * 60)
print("TOTAL MAPPINGS:", len(INDIVIDUAL_DIRECT_COST_MAPPINGS))
print("PASSED:", passed)
print("NEEDS ATTENTION:", failed)
