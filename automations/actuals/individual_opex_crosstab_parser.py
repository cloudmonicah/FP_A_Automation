"""Extract numeric Actual YTD from full Tableau clipboard crosstabs."""
import csv
from datetime import datetime
from decimal import Decimal
from pathlib import Path


def _number(value):
    s=str(value or "").strip().replace(",", "").replace("$", "")
    if not s: return Decimal(0)
    if "%" in s: raise ValueError("Percentage is not an Actual YTD amount")
    if s.startswith("(") and s.endswith(")"): s="-"+s[1:-1]
    return Decimal(s)


def _is_month(value):
    try:
        datetime.strptime(value,"%B %Y")
        return True
    except ValueError:
        return False


def extract_monthly_ytd(raw_path, company, closed_period, fiscal_start_month=7):
    with Path(raw_path).open(encoding="utf-8-sig",newline="") as fh:
        rows=[[str(v).strip() for v in row] for row in csv.reader(fh,delimiter="\t")]
    rows=[r for r in rows if any(r)]
    if len(rows)<3: raise ValueError("Incomplete Tableau crosstab")
    if rows[0][0].casefold()=="company":
        # Single-header Tableau exports: each month is already the Actual value.
        periods=rows[0]
        measures=["actual ytd" if (v.casefold()=="grand total" or _is_month(v)) else v for v in periods]
        data_rows=rows[1:]
    else:
        periods,measures=rows[:2]
        data_rows=rows[2:]
    if len(periods)!=len(measures): raise ValueError("Misaligned Tableau crosstab headers")
    if measures[0].casefold()!="company": raise ValueError("Missing Company header")
    actual_cols={}; grand=[]
    for i,(period,measure) in enumerate(zip(periods,measures)):
        if measure.casefold().strip() not in ("actual ytd","actual"):
            continue
        if period.casefold()=="grand total": grand.append(i); continue
        try: dt=datetime.strptime(period,"%B %Y")
        except ValueError: continue
        key=(dt.year,dt.month)
        if key in actual_cols: raise ValueError(f"Duplicate actual column for {period}")
        actual_cols[key]=i
    selected=(closed_period.year,closed_period.month)
    year=closed_period.year if closed_period.month>=fiscal_start_month else closed_period.year-1
    month=fiscal_start_month; expected=[]
    while (year,month)<=selected:
        expected.append((year,month)); month+=1
        if month==13: year,month=year+1,1
    missing=[k for k in expected if k not in actual_cols]
    if missing: raise ValueError(f"Missing Actual YTD reporting months: {missing}; found {list(actual_cols)}")
    if len(grand)!=1: raise ValueError("Expected one Grand Total Actual YTD column")
    company_rows=[]
    grand_rows=[]
    for row in data_rows:
        row=row+[""]*(len(periods)-len(row))
        if row[0].casefold()=="grand total":
            grand_rows.append(row)
        if row[0].upper()!=company.upper(): continue
        # Prefer company Total; ignore subtotals for GL categories.
        dims=row[1:next((i for i,p in enumerate(periods) if p),len(row))]
        if dims and all(str(v).casefold()=="total" for v in dims): company_rows.append(row)
    if len(company_rows)==1:
        row=company_rows[0]
    elif len(company_rows)==0 and len(grand_rows)==1:
        # Tableau often labels its sole-company final row Grand Total.
        # Only accept it when all detailed company rows belong to the requested company.
        other_companies={r[0].upper() for r in data_rows if r and r[0].strip().upper() not in ("GRAND TOTAL",company.upper())}
        if other_companies:
            raise ValueError(f"Grand Total includes other companies: {other_companies}")
        row=grand_rows[0]
    else:
        raise ValueError(f"Expected one unambiguous company total for {company}; found {len(company_rows)}")
    monthly=_number(row[actual_cols[selected]])
    calculated=sum((_number(row[actual_cols[k]]) for k in expected),Decimal(0))
    total=_number(row[grand[0]])
    difference=abs(calculated-total)
    if difference>Decimal(2): raise ValueError(f"YTD does not reconcile: months={calculated} total={total}")
    return {"MONTH":monthly,"YTD":total,"YTD_CHECK":calculated,"YTD_DIFFERENCE":difference}
