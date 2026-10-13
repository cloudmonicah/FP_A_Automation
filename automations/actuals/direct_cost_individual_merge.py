"""Safely merge mapped individual Direct Cost crosstab actuals into compact report.

Monthly and YTD amounts are selected from explicit Tableau period/measure columns. Values in Tableau
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
    """Read Tableau's period row + measure row together (both Actual and Actual YTD)."""
    if len(rows) < 3:
        return None
    period_headers = [_key(x) for x in rows[0]]
    measure_headers = [_key(x) for x in rows[1]]
    category = next((i for i, h in enumerate(measure_headers)
                     if h in ('GLACCOUNTMAPPING', 'DIRECTCOSTCATEGORY', 'CATEGORY', 'ACCOUNTMAPPING')), None)
    company = next((i for i, h in enumerate(measure_headers) if h == 'COMPANY'), None)
    month_key = _key(closed_period.strftime('%B %Y'))
    def cols(period):
        return [i for i, (p, m) in enumerate(zip(period_headers, measure_headers))
                if p == period and m in ('ACTUAL', 'ACTUALYTD')]
    monthly = cols(month_key)
    ytd = cols('GRANDTOTAL')
    # If a view only exports a running YTD figure for the month, do not
    # mistake that figure for a monthly amount. Tableau's explicit Grand Total
    # in the supplied exports is the fiscal-to-date actual across July/August.
    return (1, category, company,
            monthly[0] if len(monthly) == 1 else None,
            ytd[0] if len(ytd) == 1 else None)


def _extract(ws, route, closed_period):
    rows = list(ws.values)
    found = _find_header(rows, closed_period)
    if not found:
        return None, None, 'Missing Tableau period/measure header rows'
    header, cat_col, company_col, month_col, ytd_col = found
    if month_col is None and ytd_col is None:
        return None, None, 'No unique period-specific Actual and Grand Total Actual columns'

    company = _key(route['company'])
    line = _key(route['fs_line'])
    aggregation = str(route.get('aggregation', '')).upper()
    categories = {_key(x) for x in route.get('tableau_categories', []) if x}
    # Correct the legacy ALIV reseller labels to the actual GL Account Mapping
    # values observed in both the consolidated and individual Tableau exports.
    if company == 'ALIV' and line == 'RESELLERCOMMISSIONS':
        categories = {'PUSHTOTALKCOMMISSION', 'TOPUPANDPLANCOMMISSION'}
    # Control total: sum distinct detail rows in the single ALIV source capture;
    # never add it on top of the component mapping values.
    if aggregation == 'TOTAL' and company == 'ALIV' and line == 'DIRECTCOST':
        categories = {'INTERCONNECTCOST', 'PUSHTOTALKCOMMISSION',
                      'TOPUPANDPLANCOMMISSION', 'ROAMINGCOSTS', 'SERVICEFEES'}
    if not categories and cat_col is not None:
        return None, None, 'No verified Tableau category mapping'

    # Some dedicated views (CBL Programming Fees) export just one total row
    # with no category or company dimension. Permit only a known dedicated
    # workbook, exactly one numeric detail row and the mapped CBL category.
    dimensionless = cat_col is None
    if dimensionless and not (company == 'CBL' and line == 'PROGRAMMINGFEES'):
        return None, None, 'No category column; cannot safely filter this worksheet'

    matched = []
    for row in rows[header + 1:]:
        if not any(v is not None and str(v).strip() for v in row):
            continue
        if dimensionless:
            matched.append(row)
            continue
        if cat_col >= len(row) or _key(row[cat_col]) not in categories:
            continue
        if company_col is not None and company_col < len(row):
            if _key(row[company_col]) != company:
                continue
        # Exclude preaggregated Tableau subtotal/total rows.
        descriptors = [_key(x) for x in row[:max(cat_col + 1, (company_col or 0) + 1)]]
        if any(x in ('GRANDTOTAL', 'SUBTOTAL', 'TOTAL') for x in descriptors):
            continue
        matched.append(row)
    if dimensionless and len(matched) != 1:
        return None, None, f'Dedicated view has {len(matched)} rows; expected exactly one'
    if not matched:
        return None, None, 'No exact matching Tableau category/company rows'

    def aggregate(column):
        if column is None:
            return None
        values = [_number(row[column] if column < len(row) else None) for row in matched]
        return sum(values, Decimal(0)) if all(v is not None for v in values) else None

    monthly, ytd = aggregate(month_col), aggregate(ytd_col)
    note = (f'Summed {len(matched)} matching detail row(s); '
            f'Monthly: {monthly if monthly is not None else "unavailable"}; '
            f'YTD: {ytd if ytd is not None else "unavailable"}')
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
