from pathlib import Path
import xml.etree.ElementTree as ET

workbooks = {
    "CONSOLIDATED": Path(
        r"E:\Workbooks\Direct Cost\Cable Bahamas Group Direct Costs.twb"
    ),
    "INDIVIDUAL": Path(r"E:\Workbooks\Direct Cost\CBL Content Fees.twb"),
}

keywords = (
    "GL Business Code",
    "Fiscal Year",
    "GL Posting Date",
    "Type",
    "Company",
    "Business",
)

for label, path in workbooks.items():
    print("\n" + "=" * 70)
    print(label, path)
    print("=" * 70)

    root = ET.parse(path).getroot()

    for worksheet in root.findall("./worksheets/worksheet"):
        name = worksheet.get("name", "")

        if label == "CONSOLIDATED" and name != "Group Direct Costs Details":
            continue

        if label == "INDIVIDUAL" and name != "CBL Content Fees Summary":
            continue

        print("WORKSHEET:", name)

        for filt in worksheet.findall(".//filter"):
            column = filt.get("column", "")

            if not any(k.lower() in column.lower() for k in keywords):
                continue

            print("\nFILTER:", column)

            for group in filt.findall(".//groupfilter"):
                print(" ", group.attrib)

print("\nCOMPLETE: Read-only comparison.")
