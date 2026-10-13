from pathlib import Path
import xml.etree.ElementTree as ET

from automations.actuals.mapping import INDIVIDUAL_DIRECT_COST_MAPPINGS

route = next(
    r
    for r in INDIVIDUAL_DIRECT_COST_MAPPINGS
    if r["company"] == "CBL" and r["fs_line"].casefold() == "programming fees"
)

source = Path(route["individual_tableau_workbook"])
tab = route["individual_tableau_worksheet"]
root = ET.parse(source).getroot()

sheet = next(s for s in root.findall("./worksheets/worksheet") if s.get("name") == tab)

print("SOURCE:", source)
print("WORKSHEET:", tab)

print("\nACTUAL AND DATE DEPENDENCIES")
for dependency in sheet.findall(".//datasource-dependencies"):
    print("DATASOURCE:", dependency.get("datasource"))

    for column in dependency.findall("column"):
        label = str(column.get("caption") or column.get("name") or "")

        if any(
            word in label.lower() for word in ("actual", "date", "forecast", "company")
        ):
            calculation = column.find("calculation")
            print("FIELD:", label)
            print(
                "FORMULA:",
                calculation.get("formula") if calculation is not None else "",
            )

print("\nWORKSHEET FILTERS")
for filter_element in sheet.findall(".//filter"):
    print("FILTER:", filter_element.get("column"))

    for group in filter_element.findall(".//groupfilter"):
        print("  ", group.attrib)

print("\nGLOBAL ACTUAL YTD CALCULATIONS")
for column in root.iter("column"):
    caption = str(column.get("caption") or "")

    if "actual ytd" in caption.lower():
        calculation = column.find("calculation")
        print(
            caption,
            column.get("name"),
            calculation.get("formula") if calculation is not None else "",
        )

print("\nCOMPLETE: No workbook changes made.")
