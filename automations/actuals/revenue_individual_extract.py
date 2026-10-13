"""Offline-first ALIV Consumer Revenue detail extractor.
Reads a copied Tableau TSV crosstab, not a live Tableau session.
Preserves raw signed dollars; no sign conversion or forced reconciliation.
"""

import csv
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import StringIO
from pathlib import Path
from .revenue_individual_mapping import ALIV_CONSUMER_ROUTES


def norm(x):
    return re.sub(r"\s+", " ", str(x or "").strip()).upper()


def amount(x):
    x = str(x or "").strip().replace(",", "").replace("$", "")
    if not x or x in ("-", "–"):
        return None
    if x.startswith("(") and x.endswith(")"):
        x = "-" + x[1:-1]
    try:
        return Decimal(x)
    except InvalidOperation:
        return None


def extract_aliv_consumer_tsv(text, closed_period, *, fiscal_start_month=7):
    rows = [
        r
        for r in csv.reader(
            StringIO(text.lstrip("\ufeff")),
            delimiter="\t",
        )
        if any(str(v).strip() for v in r)
    ]

    if len(rows) < 2:
        raise ValueError(
            "Revenue crosstab must contain a header and at least one detail row"
        )
    width = max(map(len, rows))
    rows = [r + [""] * (width - len(r)) for r in rows]
    # Recognize Tableau's two-row (month / measure) header.
    first, second = rows[0], rows[1]
    two = any(norm(v) in ("ACTUAL", "ACTUAL YTD") for v in second) and any(
        "20" in str(v) or "GRAND TOTAL" in norm(v) for v in first
    )
    headers = []
    current = ""
    for i in range(width):
        if two:
            if norm(first[i]):
                current = first[i].strip()
            headers.append(
                norm((current + " " + second[i]).strip())
                if norm(second[i])
                in ("ACTUAL", "ACTUAL YTD", "BUDGET", "VAR $", "VAR %")
                else norm(second[i] or first[i])
            )
        else:
            headers.append(norm(first[i]))
    data = rows[2:] if two else rows[1:]
    if not data:
        raise ValueError("No Revenue detail rows found in Tableau crosstab")

    def col(name):
        matches = [i for i, h in enumerate(headers) if h == norm(name)]
        if len(matches) != 1:
            raise ValueError(
                f"Expected one {name} column; found {len(matches)}. Headers: {headers}"
            )
        return matches[0]

    company_col = col("Company")
    type_col = col("GL FS Type")
    monthly_names = [
        closed_period.strftime("%B %Y ACTUAL").upper(),
        closed_period.strftime("%b-%y ACTUAL").upper(),
    ]
    monthly_cols = [i for i, h in enumerate(headers) if h in monthly_names]
    if len(monthly_cols) != 1:
        raise ValueError(
            f"Could not uniquely identify monthly Actual for {closed_period:%B %Y}: {headers}"
        )
    monthly_col = monthly_cols[0]
    ytd_cols = [
        i
        for i, h in enumerate(headers)
        if h in ("GRAND TOTAL ACTUAL", "GRAND TOTAL ACTUAL YTD", "ACTUAL YTD")
    ]
    ytd_col = ytd_cols[0] if len(ytd_cols) == 1 else None
    # If YTD is not explicit, sum available fiscal months, only when all month headers are present.
    if ytd_col is None:
        start_year = closed_period.year - (closed_period.month < fiscal_start_month)
        months = []
        y, m = start_year, fiscal_start_month
        while (y, m) <= (closed_period.year, closed_period.month):
            label = date(y, m, 1).strftime("%B %Y ACTUAL").upper()
            found = [i for i, h in enumerate(headers) if h == label]
            if len(found) != 1:
                raise ValueError(f"YTD unavailable: missing {label}")
            months.append(found[0])
            y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    output = {}
    for route in ALIV_CONSUMER_ROUTES:
        matched = []
        last_company = last_type = ""
        for row in data:
            if norm(row[company_col]):
                last_company = norm(row[company_col])
            if norm(row[type_col]):
                last_type = norm(row[type_col])
            if last_company != "ALIV" or last_type != norm(route["gl_fs_type"]):
                continue
            # Exclude explicit Tableau subtotal/grand-total rows to prevent double-counting.
            descriptors = [norm(v) for v in row[:monthly_col]]
            if any(
                v in ("TOTAL", "GRAND TOTAL") or v.startswith("TOTAL ")
                for v in descriptors
            ):
                continue
            mon = amount(row[monthly_col])
            if mon is None:
                continue
            if ytd_col is not None:
                ytd = amount(row[ytd_col])
            else:
                vals = [amount(row[i]) for i in months]
                ytd = sum(vals) if all(v is not None for v in vals) else None
            matched.append((mon, ytd))
        if not matched:
            output[route["fs_line"]] = {
                "status": "REVIEW REQUIRED",
                "reason": "No matching detail rows",
                "monthly": None,
                "ytd": None,
                "rows": 0,
            }
        else:
            monthly = sum(x[0] for x in matched)
            ytd = (
                sum(x[1] for x in matched)
                if all(x[1] is not None for x in matched)
                else None
            )
            output[route["fs_line"]] = {
                "status": (
                    "EXTRACTED - NOT VALIDATED"
                    if ytd is not None
                    else "REVIEW REQUIRED"
                ),
                "monthly": monthly,
                "ytd": ytd,
                "rows": len(matched),
            }
    return output


def extract_file(path, closed_period):
    return extract_aliv_consumer_tsv(
        Path(path).read_text(encoding="utf-8-sig"), closed_period
    )
