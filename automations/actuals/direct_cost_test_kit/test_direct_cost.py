"""Run consolidated Direct Cost validation for a chosen closed month.
Example:
 python test_direct_cost.py --aliv "C:\\Statements\\August 2026.xlsx" --cbl "C:\\Statements\\1 - CBL Standalone Financials 08.31.26.xlsx" --month 2026-08 --output "C:\\Temp\\DirectCostTest"
"""
import argparse
from pathlib import Path
from datetime import date
from automations.actuals.validation_config import DIRECT_COST_MAPPINGS, DIRECT_COST_TABLEAU_WORKSHEET
from automations.actuals.direct_cost_validation_service import run_live_direct_cost_validation

def main():
    p=argparse.ArgumentParser()
    p.add_argument('--aliv', required=True)
    p.add_argument('--cbl', required=True)
    p.add_argument('--month', default='2026-08', help='YYYY-MM')
    p.add_argument('--tableau', default=r'E:\Workbooks\Direct Cost\Cable Bahamas Group Direct Costs.twb')
    p.add_argument('--output', default=r'C:\Temp\DirectCostTest')
    p.add_argument('--open-tableau', action='store_true')
    args=p.parse_args()
    period=date.fromisoformat(args.month+'-01')
    import calendar
    period=date(period.year,period.month,calendar.monthrange(period.year,period.month)[1])
    for filename in (args.aliv,args.cbl,args.tableau):
        if not Path(filename).is_file():
            raise FileNotFoundError(filename)
    out=Path(args.output)
    paths={'tableau_raw_copies':out/'Tableau_Raw_Copies','direct_cost':out/'Direct_Cost','logs':out/'Logs'}
    for folder in paths.values():folder.mkdir(parents=True,exist_ok=True)
    print('Period:',period,'Mappings:',len(DIRECT_COST_MAPPINGS),'Worksheet:',DIRECT_COST_TABLEAU_WORKSHEET)
    print('Individual Tableau capture is NOT part of this consolidated smoke test.')
    result=run_live_direct_cost_validation(aliv_fs_path=args.aliv,cbl_fs_path=args.cbl,output_paths=paths,closed_period=period,tableau_workbook_path=args.tableau,open_tableau_first=args.open_tableau)
    print('STATUS:',result.get('overall_status'))
    print('REPORT:',result.get('validation_workbook'))
    print('RAW:',result.get('raw_tableau_workbook'))
if __name__=='__main__':main()
