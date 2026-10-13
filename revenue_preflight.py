"""Read-only Revenue source preflight. No Tableau Desktop or workbook writes."""
from pathlib import Path
import argparse
import xml.etree.ElementTree as ET

SOURCES = {
    'group': ('CBL Group Revenue.twb', 'CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY'),
    'consumer': ('Cable Bahamas Group_CONSUMER REVENUE.twb', 'ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD'),
    'corporate': ('Cable Bahamas Group_CORPORATE REVENUE.twb', 'CBL COMMERCIAL REVENUE DETAILS - Monthly Actuals'),
}

def check(root: Path):
    failures = 0
    for key, (filename, worksheet) in SOURCES.items():
        file = root / filename
        if not file.is_file():
            print(f'[MISSING] {key}: {file}')
            failures += 1
            continue
        try:
            document = ET.parse(file).getroot()
            sheets = {node.get('name') for node in document.findall('./worksheets/worksheet')}
            if worksheet not in sheets:
                print(f'[MISSING WORKSHEET] {key}: {worksheet}')
                failures += 1
            else:
                print(f'[OK] {key}: {filename} -> {worksheet}')
                sheet = next(n for n in document.findall('./worksheets/worksheet') if n.get('name') == worksheet)
                fields = sorted({x.get('caption') for x in sheet.iter('column') if x.get('caption')})
                print('  Worksheet-local fields:', ', '.join(fields[:35]) or '(defined in datasource)')
        except ET.ParseError as exc:
            print(f'[INVALID TWB] {file}: {exc}')
            failures += 1
    print(f'\nPreflight: {len(SOURCES)-failures}/{len(SOURCES)} source worksheets verified')
    return failures

if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--revenue-folder', default=r'E:\Workbooks\Revenue')
    args = parser.parse_args()
    raise SystemExit(1 if check(Path(args.revenue_folder)) else 0)
