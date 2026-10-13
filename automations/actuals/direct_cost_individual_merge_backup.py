"""Safely merge mapped individual Direct Cost crosstab actuals into compact report.

No inference of monthly actual from YTD, budgets or forecasts. Values in Tableau
are dollars; approved validation report displays thousands.
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from automations.actuals.mapping import INDIVIDUAL_DIRECT_COST_MAPPINGS


def _key(value):
    return re.sub(r"[^A-Z0-9]", "", str(value or "").upper())


def _number(value):
    if isinstance(value, (int, float, Decimal)) and not isinstance(value, bool):
        return Decimal(str(value))
    s = str(value or '').strip().replace(',', '').replace('$', '')
    if not s or '%' in s or s in ('-', '—'):
        return None
    if s.startswith('(') and s.endswith(')'):
        s = '-' + s[1:-1]
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def _thousands(value):
    return int((value / Decimal(1000)).quantize(Decimal('1'), rounding=ROUND_HALF_UP))


def _find_header(rows, closed_period):
    """Return (header_index, category_col, company_col, month_col, ytd_col)."""
    for index, row in enumerate(rows[:7]):
        headers = [_key(v) for v in row]
        cat = next((i for i, h in enumerate(headers) if h in ('GLACCOUNTMAPPING', 'DIRECTCOSTCATEGORY', 'CATEGORY', 'ACCOUNTMAPPING')), None)
        if cat is None:
            continue
        company = next((i for i, h in enumerate(headers) if h == 'COMPANY'), None)
        ytd = [i for i, h in enumerate(headers) if h in ('ACTUALYTD', 'YTDACTUAL', 'YEARToDATEACTUAL'.upper())]
        # Accept explicit monthly actual columns only. Never treat generic 'Actual' as monthly.
        month = [i for i, h in enumerate(headers) if h in ('ACTUALMONTH', 'MONTHLYACTUAL', 'MONTHACTUAL', 'ACTUALMTD', 'MTDACTUAL')]
        # A period-labelled column is accepted only if it explicitly identifies actuals.
        period = closed_period.strftime('%b%Y').upper()
        month += [i for i, h in enumerate(headers) if 'ACTUAL' in h and period in h and i not in month]
        return index, cat, company, (month[0] if len(set(month)) == 1 else None), (ytd[0] if len(ytd) == 1 else None)
    return None


def _extract(ws, route, closed_period):
    rows = list(ws.values)
    found = _find_header(rows, closed_period)
    if not found:
        return None, None, 'No recognized category/header row'
    header, cat_col, company_col, month_col, ytd_col = found
    if month_col is None and ytd_col is None:
        return None, None, 'No unambiguous Monthly Actual or Actual YTD measure'
    if str(route.get('aggregation', '')).upper() == 'TOTAL':
        return None, None, 'Control total: not inferred from component sums'
    categories = {_key(x) for x in route.get('tableau_categories', []) if x}
    if not categories:
        return None, None, 'No verified Tableau category mapping'
    matched = []
    for row in rows[header + 1:]:
        if cat_col >= len(row) or _key(row[cat_col]) not in categories:
            continue
        if company_col is not None and company_col < len(row):
            if _key(row[company_col]) != _key(route['company']):
                continue
        matched.append(row)
    if not matched:
        return None, None, 'No exact matching Tableau category/company rows'
    def aggregate(column):
        if column is None:
            return None
        values = [_number(row[column] if column < len(row) else None) for row in matched]
        return sum(values, Decimal(0)) if all(v is not None for v in values) else None
    monthly = aggregate(month_col)
    ytd = aggregate(ytd_col)
    note = f'Matched {len(matched)} rows; Monthly: {"available" if monthly is not None else "unavailable"}; YTD: {"available" if ytd is not None else "unavailable"}'
    return (_thousands(monthly) if monthly is not None else None,
            _thousands(ytd) if ytd is not None else None, note)


def merge_direct_cost_crosstabs(*, report_path, crosstab_path, closed_period):
    """Update report in place; keep unavailable cells empty and status REVIEW REQUIRED."""
    report_path, crosstab_path = Path(report_path), Path(crosstab_path)
    report = load_workbook(report_path)
    raw = load_workbook(crosstab_path, read_only=True, data_only=True)
    results = []
    try:
        for index, route in enumerate(INDIVIDUAL_DIRECT_COST_MAPPINGS, 1):
            company, line = route['company'], route['fs_line']
            name = f'{index:03d}_{company}_{line}'
            name = re.sub(r'[\\/*?:\[\]]', '_', name)[:31]
            sheet_name = f'{company} Direct Cost'
            if sheet_name not in report:
                raise ValueError(f'Missing validation worksheet: {sheet_name}')
            sheet = report[sheet_name]
            if _key(sheet['A7'].value) != 'FINANCIALSTATEMENTLINE':
                raise ValueError(f'Unexpected report format in {sheet_name}')
            matches = [r for r in range(8, sheet.max_row + 1) if _key(sheet.cell(r, 1).value) == _key(line)]
            if len(matches) != 1:
                results.append((company, line, 'SKIPPED', f'Expected one report row, found {len(matches)}'))
                continue
            row = matches[0]
            monthly = ytd = None
            if name in raw.sheetnames:
                monthly, ytd, note = _extract(raw[name], route, closed_period)
            else:
                note = 'Capture worksheet missing (failed or skipped Tableau capture)'
            # Report schema: D monthly individual, F monthly variance 2, G monthly status;
            # J YTD individual, L YTD variance 2, M YTD status.
            for value, individual_col, consolidated_col, variance_col, status_col in (
                (monthly, 4, 3, 6, 7), (ytd, 10, 9, 12, 13)
            ):
                sheet.cell(row, individual_col).value = value
                baseline = _number(sheet.cell(row, consolidated_col).value)
                sheet.cell(row, variance_col).value = (float(baseline - Decimal(value))
                                                       if value is not None and baseline is not None else None)
                # Do not mark PASS on an incomplete capture; existing FS vs group status
                # is preserved in the separate variance 1 columns.
                status = 'REVIEW REQUIRED'
                sheet.cell(row, status_col).value = status
                sheet.cell(row, status_col).fill = PatternFill('solid', fgColor='FFF2CC')
                sheet.cell(row, status_col).font = Font(color='9C5700', bold=True)
            results.append((company, line, 'YTD CAPTURED' if ytd is not None else 'REVIEW REQUIRED', note))
        report.save(report_path)
    finally:
        raw.close()
        report.close()
    return {'processed': len(results), 'ytd_populated': sum(x[2] == 'YTD CAPTURED' for x in results),
            'details': results, 'report_path': str(report_path)}
