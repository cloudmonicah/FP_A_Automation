"""Locate the monthly / YTD "Actual" columns of an ALIV statement by header text.

ALIV statements are not laid out the same way every month (July 2026 has a
12-month trend block on the left, August starts with the Actual column), so no
column letter or cell address may be assumed. Columns are found from the
"Actual" headers and the "YTD" heading above them, and the reporting month is
taken from the date cells above the table.
"""
from __future__ import annotations

import re
from datetime import date, datetime

_TEXT_MONTH_FORMATS = ("%B %Y", "%b %Y", "%b-%y", "%B-%y", "%Y-%m")
HEADER_SCAN_ROWS = 35
PERIOD_SCAN_ROWS = 8


def _norm(value) -> str:
    return " ".join(
        re.sub(r"^\d+\s+", "", str(value or "")).upper().replace("&", " AND ").split()
    )


def _year_month(value):
    if isinstance(value, (datetime, date)):
        return value.year, value.month
    if isinstance(value, str):
        text = value.strip()
        for fmt in _TEXT_MONTH_FORMATS:
            try:
                parsed = datetime.strptime(text, fmt)
                return parsed.year, parsed.month
            except ValueError:
                pass
    return None


def reporting_months(rows, scan_rows: int = PERIOD_SCAN_ROWS) -> list[tuple[int, int]]:
    """Months found in the cells above the table (the statement's reporting date)."""
    found = []
    for row in rows[:scan_rows]:
        for value in row:
            ym = _year_month(value)
            if ym and ym not in found:
                found.append(ym)
    return found


def locate_aliv_columns(rows, closed_period, name: str = "statement"):
    """Return (header_row, month_col, ytd_col), all 0-based, for the selected month.

    Raises a descriptive ValueError if the statement is for another month or the
    Actual / Actual YTD headers cannot be identified unambiguously.
    """
    wanted = (closed_period.year, closed_period.month)
    months = reporting_months(rows)
    if wanted not in months:
        shown = ", ".join(f"{y}-{m:02d}" for y, m in months) or "no month found"
        raise ValueError(
            f"ALIV statement {name} is for {shown} (read from the top rows), but the "
            f"selected close is {wanted[0]}-{wanted[1]:02d}. Select the matching statement file."
        )
    candidates = []
    for r, row in enumerate(rows[:HEADER_SCAN_ROWS]):
        for c, value in enumerate(row):
            if _norm(value) != "ACTUAL":
                continue
            nearby = " ".join(
                _norm(rows[rr][cc])
                for rr in range(max(0, r - 4), r)
                for cc in range(max(0, c - 1), min(len(rows[rr]), c + 2))
            )
            candidates.append((r, c, "YTD" in nearby))
    pairs = []
    for r, c, is_ytd in candidates:
        if is_ytd:
            continue
        later = [cc for rr, cc, flag in candidates if flag and rr == r and cc > c]
        if later:
            pairs.append((r, c, later[0]))
    if len(pairs) != 1:
        raise ValueError(
            f"ALIV statement {name}: could not identify one Actual / Actual YTD header "
            f"pair (found {len(pairs)}). Check the column headings above the table."
        )
    return pairs[0]


def column_letter(index0: int) -> str:
    from openpyxl.utils import get_column_letter

    return get_column_letter(index0 + 1)
