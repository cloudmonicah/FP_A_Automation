"""Read-only preflight for individual Direct Cost Tableau TWB mappings.
Run from the FPA_Automation project root. Does not open Tableau or alter TWBs.
"""
from __future__ import annotations
import argparse
import csv
import importlib.util
from pathlib import Path
import xml.etree.ElementTree as ET


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--mapping', default='automations/actuals/mapping.py')
    parser.add_argument('--company', choices=['ALL', 'ALIV', 'CBL'], default='ALL')
    parser.add_argument('--output', default='C:/Temp/DirectCostTest/individual_twb_preflight.csv')
    args = parser.parse_args()
    mapping_path = Path(args.mapping).resolve()
    spec = importlib.util.spec_from_file_location('local_actuals_mapping', mapping_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    rows = []
    for route in module.INDIVIDUAL_DIRECT_COST_MAPPINGS:
        company = route['company']
        if args.company != 'ALL' and company != args.company:
            continue
        path = Path(route['individual_tableau_workbook'])
        tab = route['individual_tableau_worksheet']
        result = dict(company=company, fs_line=route['fs_line'], workbook=str(path), worksheet=tab,
                      exists=path.is_file(), worksheet_found=False, status='PENDING',
                      measure_captions='', filter_fields='', note=route.get('note', ''), error='')
        if not path.is_file():
            result['error'] = 'Workbook path not found or inaccessible'
        else:
            try:
                root = ET.parse(path).getroot()
                sheets = root.findall('./worksheets/worksheet')
                matches = [s for s in sheets if s.get('name') == tab]
                if len(matches) != 1:
                    result['error'] = ('Expected worksheet not found; available: ' +
                                       ' | '.join(s.get('name', '') for s in sheets))
                else:
                    sheet = matches[0]
                    result['worksheet_found'] = True
                    captions = sorted({c.get('caption', '') for c in sheet.iter('column') if c.get('caption')})
                    filters = sorted({f.get('column', '') for f in sheet.iter('filter') if f.get('column')})
                    result['measure_captions'] = ' | '.join(captions)
                    result['filter_fields'] = ' | '.join(filters)
                    result['status'] = 'READY FOR CROSSTAB TEST'
            except (ET.ParseError, OSError, ValueError) as exc:
                result['error'] = f'{type(exc).__name__}: {exc}'
        rows.append(result)
        print(f"{company:4} | {route['fs_line']:<28} | {result['status']:<25} | {result['error']}")
    destination = Path(args.output)
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open('w', newline='', encoding='utf-8-sig') as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0]) if rows else [])
        if rows:
            writer.writeheader()
            writer.writerows(rows)
    print('Preflight report:', destination)
    print('NOTE: READY means only path and worksheet verified, not reconciled.')

if __name__ == '__main__':
    main()
