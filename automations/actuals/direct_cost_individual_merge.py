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
from automations.actuals.mapping import INDIVIDUAL_DIRECT_COST_MAPPINGS
from automations.actuals.report_builder import update_individual_values


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


def _extract_dedicated_total(rows, header, cat_col, company_col, month_col, ytd_col, company):
    """A worksheet that already holds this line's total: read its total row.

    Prefers the company's own total row, then the Grand Total row, then a lone
    data row. Anything else is reported instead of guessed.
    """
    data = [r for r in rows[header + 1:] if any(v is not None and str(v).strip() for v in r)]
    width = max(cat_col or 0, company_col or 0) + 1
    def labels(row):
        return [_key(x) for x in row[:max(width, 1)]]
    def is_total(row):
        return any(x in ('GRANDTOTAL', 'TOTAL', 'SUBTOTAL') for x in labels(row))
    own = [r for r in data if is_total(r) and company_col is not None and company_col < len(r)
           and _key(r[company_col]) == company]
    grand = [r for r in data if any(x == 'GRANDTOTAL' for x in labels(r))]
    pick = own[0] if len(own) == 1 else grand[0] if len(grand) == 1 else data[0] if len(data) == 1 else None
    if pick is None:
        return None, None, 'The Reseller Commissions total row could not be identified in the Tableau table'
    def value(column):
        if column is None or column >= len(pick):
            return None
        return _number(pick[column])
    monthly, ytd = value(month_col), value(ytd_col)
    note = (f'Read the worksheet total row; Monthly: {monthly if monthly is not None else "unavailable"}; '
            f'YTD: {ytd if ytd is not None else "unavailable"}')
    return (_thousands(monthly) if monthly is not None else None,
            _thousands(ytd) if ytd is not None else None, note)


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
    if route.get('dedicated_total'):
        return _extract_dedicated_total(rows, header, cat_col, company_col, month_col, ytd_col, company)
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
    """Write captured individual values into the final report and rebuild it.

    Lines whose capture failed keep an empty individual value, so they stay PENDING
    with the capture reason in their notes.
    """
    report_path, crosstab_path = Path(report_path), Path(crosstab_path)
    updates = {company: {} for company in ("ALIV", "CBL")}
    results = []
    raw = load_workbook(crosstab_path, read_only=True, data_only=True)
    try:
        for index, route in enumerate(INDIVIDUAL_DIRECT_COST_MAPPINGS, 1):
            company, line = route['company'], route['fs_line']
            name = re.sub(r'[\\/*?:\[\]]', '_', f'{index:03d}_{company}_{line}')[:31]
            monthly = ytd = None
            if name in raw.sheetnames:
                monthly, ytd, note = _extract(raw[name], route, closed_period)
            else:
                note = 'Capture worksheet missing (failed or skipped Tableau capture)'
            reason = None if monthly is not None and ytd is not None else f'Individual capture: {note}'
            updates[company][line] = (monthly, ytd, reason)
            results.append([company, line, 'YTD CAPTURED' if ytd is not None else 'REVIEW REQUIRED', note])
    finally:
        raw.close()

    for company, company_updates in updates.items():
        outcome = update_individual_values(report_path, company, company_updates)
        for result in results:
            if result[0] == company and outcome.get(result[1]) == 'missing':
                result[2:] = ['SKIPPED', 'Expected one report row for this line']

    results = [tuple(r) for r in results]
    return {'processed': len(results), 'ytd_populated': sum(x[2] == 'YTD CAPTURED' for x in results),
            'details': results, 'report_path': str(report_path)}
