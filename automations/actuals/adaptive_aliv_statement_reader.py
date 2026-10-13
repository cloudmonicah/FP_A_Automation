"""Discover ALIV monthly/YTD actuals by headings and labels (no fixed cells)."""
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import re
from openpyxl import load_workbook


def norm(v):
    return ' '.join(re.sub(r'^\d+\s+', '', str(v or '')).upper().replace('&',' AND ').split())


def _date(v):
    return v if isinstance(v,(date,datetime)) else None


def read_aliv_period(path, closed_period):
    wb=load_workbook(Path(path),read_only=True,data_only=True)
    try:
        if 'P & L_' not in wb: raise ValueError('ALIV P & L_ worksheet not found')
        ws=wb['P & L_']
        rows=list(ws.iter_rows(values_only=True))
        dates=[(r,c,v) for r,row in enumerate(rows[:30]) for c,v in enumerate(row) if _date(v)]
        if not any((v.year,v.month)==(closed_period.year,closed_period.month) for _,_,v in dates):
            raise ValueError(f'ALIV statement period does not match {closed_period:%Y-%m}')
        candidates=[]
        for r,row in enumerate(rows[:35]):
            for c,v in enumerate(row):
                if norm(v)!='ACTUAL': continue
                preceding=' '.join(norm(rows[rr][cc]) for rr in range(max(0,r-4),r) for cc in range(max(0,c-1),min(len(rows[rr]),c+2)))
                # Detect YTD from section heading above, even if headings move.
                is_ytd='YTD' in preceding
                candidates.append((r,c,is_ytd))
        # A monthly header must precede a YTD header in both provided layouts.
        pairs=[]
        for r,c,ytd in candidates:
            if ytd: continue
            next_ytd=[(rr,cc) for rr,cc,flag in candidates if flag and rr==r and cc>c]
            if next_ytd: pairs.append((r,c,next_ytd[0][1]))
        if len(pairs)!=1:
            raise ValueError(f'ALIV Actual/Actual YTD header pairing ambiguous: {pairs}')
        header,month_col,ytd_col=pairs[0]
        out={}
        for r in range(header+1,len(rows)):
            row=rows[r]
            label=norm(row[1] if len(row)>1 else None)
            if not label: continue
            month=row[month_col] if len(row)>month_col else None
            ytd=row[ytd_col] if len(row)>ytd_col else None
            if label in out:
                # duplicate labels are ambiguous; do not silently pick one
                out[label]=None
            elif isinstance(month,(int,float,Decimal)) and isinstance(ytd,(int,float,Decimal)):
                out[label]=(float(month),float(ytd))
        if closed_period.month == 7:
            inconsistent=[k for k,pair in out.items() if pair is not None and abs(pair[0]-pair[1])>1.0]
            if inconsistent:
                raise ValueError("ALIV July statement YTD does not equal July monthly actual; confirm fiscal-year basis before validation. Examples: " + ", ".join(inconsistent[:4]))
        return out
    finally: wb.close()


def get_aliv_line(path,closed_period,label):
    values=read_aliv_period(path,closed_period)
    key=norm(label)
    if key not in values or values[key] is None:
        raise ValueError(f'ALIV statement line missing/ambiguous: {label}')
    return values[key]
