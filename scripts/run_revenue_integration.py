"""Run existing FP&A Revenue validation with explicit source paths.
Place at FPA_Automation project root. No source TWB is modified.
"""
import argparse
import csv
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

from automations.actuals.mapping import INDIVIDUAL_REVENUE_MAPPINGS
from automations.actuals.revenue_validation_service import run_live_revenue_validation

SOURCES = {
    'consumer': 'Cable Bahamas Group_CONSUMER REVENUE.twb',
    'corporate': 'Cable Bahamas Group_CORPORATE REVENUE.twb',
    'group': 'CBL Group Revenue.twb',
}

def audit_mappings(folder, output):
    records = []
    for route in INDIVIDUAL_REVENUE_MAPPINGS:
        workbook = str(route.get('individual_tableau_workbook') or '')
        worksheet = str(route.get('individual_tableau_worksheet') or '')
        enabled = bool(route.get('individual_actuals_enabled'))
        resolved = folder / Path(workbook.replace('\\', '/')).name if workbook else None
        exists = bool(resolved and resolved.is_file())
        sheet_exists = False
        if exists and worksheet:
            sheet_exists = any(s.get('name') == worksheet for s in ET.parse(resolved).getroot().findall('./worksheets/worksheet'))
        status = ('CONTROL TOTAL' if route.get('aggregation') == 'TOTAL' and not enabled
                  else 'PENDING MAPPING' if not enabled
                  else 'READY FOR CAPTURE' if exists and sheet_exists
                  else 'MISSING WORKBOOK/WORKSHEET')
        records.append({'company': route.get('company',''), 'fs_line':route.get('fs_line',''),
            'workbook':str(resolved or ''), 'worksheet':worksheet,
            'mapping_values':' | '.join(route.get('mapping_values', [])),
            'enabled':enabled, 'status':status})
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open('w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=list(records[0])); writer.writeheader(); writer.writerows(records)
    for r in records:
        print(f"[{r['status']}] {r['company']} / {r['fs_line']}")
    print('Mapping audit:', output)
    return records

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--revenue-folder', type=Path, default=Path(r'E:\Workbooks\Revenue'))
    parser.add_argument('--aliv-fs', type=Path)
    parser.add_argument('--cbl-fs', type=Path)
    parser.add_argument('--period', default='2026-08')
    parser.add_argument('--output', type=Path, default=Path('Revenue_Integration_Results'))
    parser.add_argument('--audit-only', action='store_true')
    args = parser.parse_args()
    audit = audit_mappings(args.revenue_folder, args.output / 'Individual_Revenue_Mapping_Audit.csv')
    if args.audit_only:
        return
    if not args.aliv_fs or not args.cbl_fs:
        parser.error('--aliv-fs and --cbl-fs are required unless --audit-only is set')
    for path in [args.aliv_fs, args.cbl_fs] + [args.revenue_folder / name for name in SOURCES.values()]:
        if not path.is_file(): parser.error(f'File not found: {path}')
    period = datetime.strptime(args.period, '%Y-%m').date()
    out = args.output
    result = run_live_revenue_validation(
        aliv_fs_path=str(args.aliv_fs), cbl_fs_path=str(args.cbl_fs),
        tableau_sources={key:str(args.revenue_folder / name) for key,name in SOURCES.items()},
        output_paths={'revenue':str(out / 'Validation_Results'),
            'tableau_raw_copies':str(out / 'Tableau_Raw_Copies'), 'logs':str(out / 'Run_Logs')},
        closed_period=period,
    )
    print('Validation workbook:', result['validation_workbook'])
    print('Overall status:', result['overall_status'])
    print('NOTE: Pending individual mappings remain unverified; do not interpret report as signed off.')

if __name__ == '__main__': main()
