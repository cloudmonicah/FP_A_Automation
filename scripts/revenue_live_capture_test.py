"""Safe Revenue live Tableau capture; no validation report modifications."""
import argparse
import csv
import io
import sys
import time
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from automations.actuals.revenue_validation_service import prepare_revenue_copy
from automations.forecast.tableau_service import open_tableau_workbook
from automations.actuals.tableau_desktop_export_service import get_tableau_crosstab_text
from automations.actuals.tableau_desktop_close_service import close_tableau_forecast_workbook

SOURCES = {
    'group': ('CBL Group Revenue.twb', 'CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY'),
    'consumer': ('Cable Bahamas Group_CONSUMER REVENUE.twb', 'ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD'),
    'corporate': ('Cable Bahamas Group_CORPORATE REVENUE.twb', 'CBL COMMERCIAL REVENUE DETAILS - Monthly Actuals'),
}

def main():
    parser = argparse.ArgumentParser(description='Capture one Revenue Tableau worksheet safely')
    parser.add_argument('--source', choices=SOURCES, default='group')
    parser.add_argument('--revenue-folder', default=r'E:\Workbooks\Revenue')
    parser.add_argument('--period', default='2026-08', help='Closed month YYYY-MM')
    parser.add_argument('--basis', choices=('MONTH','YTD'), default='MONTH')
    parser.add_argument('--output', default='Revenue_Live_Test_Results')
    parser.add_argument('--wait', type=float, default=8)
    args = parser.parse_args()
    period = datetime.strptime(args.period, '%Y-%m').date()
    filename, worksheet = SOURCES[args.source]
    source = Path(args.revenue_folder) / filename
    if not source.is_file():
        parser.error(f'Source workbook not found: {source}')
    out = Path(args.output).resolve() / f'{args.period}_{args.source}_{args.basis}'
    out.mkdir(parents=True, exist_ok=True)
    temp = source.with_name(f'{source.stem}_Revenue_Live_Test_{uuid4().hex[:8]}.twb')
    opened = False
    try:
        print(f'Preparing temporary TWB: {source}', flush=True)
        prepare_revenue_copy(source, temp, [worksheet], period, args.basis)
        print(f'Opening Tableau: {worksheet}', flush=True)
        open_tableau_workbook(tableau_workbook_path=str(temp))
        opened = True
        time.sleep(max(args.wait, 0))
        captured = get_tableau_crosstab_text(worksheet_name=worksheet, tableau_workbook_path=str(temp))
        if not captured.strip():
            raise ValueError('Tableau returned an empty crosstab')
        raw_path = out / 'Revenue_Raw_Crosstab.tsv'
        raw_path.write_text(captured, encoding='utf-8-sig')
        rows = list(csv.reader(io.StringIO(captured), delimiter='\t'))
        print(f'CAPTURED: {len(rows)} rows, max {max(map(len,rows), default=0)} columns')
        print('FIRST THREE ROWS (truncated to 15 cells):')
        for row in rows[:3]:
            print(repr(row[:15]))
        print(f'RAW CAPTURE: {raw_path}')
        print('LIVE CAPTURE PASS — amounts and mapping NOT YET validated')
    finally:
        if opened:
            try:
                close_tableau_forecast_workbook(tableau_workbook_path=str(temp))
            except Exception as exc:
                print(f'WARNING: Tableau close failed: {exc}', file=sys.stderr)
        try:
            temp.unlink(missing_ok=True)
        except OSError as exc:
            print(f'WARNING: Temporary TWB could not be deleted: {temp}: {exc}', file=sys.stderr)

if __name__ == '__main__':
    main()
