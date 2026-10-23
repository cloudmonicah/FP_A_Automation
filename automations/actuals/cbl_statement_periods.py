"""Read CBL monthly actuals from Data and reported YTD from Financials."""

from datetime import date, datetime
from pathlib import Path
from decimal import Decimal
import re
from openpyxl import load_workbook


def canonical(value):
    text = re.sub(r"^\d+\s*", "", str(value or "").strip())
    return re.sub(r"\s*-\s*", "-", " ".join(text.upper().replace("&", " AND ").split()))


# Names the CBL statement has used for the same Revenue line over time.
# Financials tab -> the name this report uses.
_REVENUE_STATEMENT_NAMES = {
    "GROSS REVENUES": "Gross Revenue",
    "SMART HOME": "Other",
}
# The report's name -> other names the Data tab uses for that line.
_REVENUE_DATA_NAMES = {
    "MSA-ALIV": ("INTERCOMPANY",),
    "OTHER": ("SMART HOME",),
}


def amount(value, label):
    if isinstance(value, bool) or not isinstance(value, (int, float, Decimal)):
        raise ValueError(f"Missing numeric CBL statement actual: {label}")
    if not Decimal(str(value)).is_finite():
        raise ValueError(f"Invalid CBL statement actual: {label}")
    return float(value)


def read_cbl_area(workbook_path, area, closed_period=None):
    """Return dictionaries in $000. No monthly amounts are added to derive YTD."""
    report = load_workbook(Path(workbook_path), read_only=True, data_only=True)
    try:
        financials = report["Financials"]
        period = financials["C2"].value
        if not isinstance(period, (date, datetime)):
            raise ValueError("CBL Financials C2 must contain the reporting date.")
        if closed_period is not None and (period.year, period.month) != (
            closed_period.year,
            closed_period.month,
        ):
            raise ValueError(
                "CBL statement period does not match the selected closed month."
            )
        closed_period = date(period.year, period.month, 1)
        earnings = report["Statement of Earnings"]
        earnings_period = earnings["D3"].value
        if not isinstance(earnings_period, (date, datetime)) or (
            earnings_period.year,
            earnings_period.month,
        ) != (period.year, period.month):
            raise ValueError(
                "CBL Statement of Earnings period does not match Financials."
            )
        if (
            canonical(earnings["D6"].value) != "CURRENT MONTH"
            or canonical(earnings["D8"].value) != "ACTUAL"
        ):
            raise ValueError(
                "CBL Statement of Earnings monthly actual headers do not match."
            )
        data = report["Data"]
        rows = list(data.iter_rows(values_only=True))
        candidates = [
            i
            for i, value in enumerate(rows[2])
            if isinstance(value, (date, datetime))
            and (value.year, value.month) == (period.year, period.month)
            and str(rows[1][i]).strip().upper() == "CY ACTUAL"
        ]
        if len(candidates) != 1:
            raise ValueError(
                "Expected one current-year actual month column on CBL Data."
            )
        month_column = candidates[0]
        sections = {
            "Revenue": ("Gross Revenue", "Cost of Goods Sold"),
            "Direct Cost": ("Cost of Goods Sold", "Gross Profit"),
            "OPEX": ("Operating Costs", "Total Operating Costs"),
        }
        start_label, end_label = sections[area]
        # September's statement says "Gross Revenues"; earlier months say "Gross Revenue".
        start_labels = {canonical(start_label)}
        if area == "Revenue":
            start_labels.add(canonical("Gross Revenues"))
        fs_rows = list(financials.iter_rows(values_only=True))
        starts = [
            i
            for i, row in enumerate(fs_rows)
            if canonical(row[1]) in start_labels
        ]
        if area in ("Revenue", "Direct Cost"):
            starts = [i for i in starts if fs_rows[i][2] is None]
        if len(starts) != 1:
            # Quarter-end statements can repeat a section (for example a quarter
            # block). Keep only the headings that sit under an "Actual YTD" header.
            def under_ytd(index):
                headers = [
                    str(row[2] or "").strip().upper()
                    for row in fs_rows[:index]
                    if str(row[2] or "").strip().upper()
                    in ("ACTUAL YTD", "CURRENT MONTH", "ACTUAL")
                ]
                return bool(headers) and headers[-1] == "ACTUAL YTD"

            narrowed = [i for i in starts if under_ytd(i)]
            if len(narrowed) == 1:
                starts = narrowed
        if len(starts) != 1:
            raise ValueError(
                f"Expected one CBL {area} section start; found {len(starts)} "
                f"heading(s) '{start_label}' on the Financials tab"
                + (f" (rows {', '.join(str(i + 1) for i in starts)})." if starts else ".")
            )
        start = starts[0]
        # Require the section's own header to state Actual YTD.
        prior_headers = [
            str(row[2] or "").strip().upper()
            for row in fs_rows[:start]
            if str(row[2] or "").strip().upper()
            in ("ACTUAL YTD", "CURRENT MONTH", "ACTUAL")
        ]
        if not prior_headers or prior_headers[-1] != "ACTUAL YTD":
            raise ValueError(f"CBL {area} Financials section is not marked Actual YTD.")
        ends = [
            i
            for i in range(start + 1, len(fs_rows))
            if canonical(fs_rows[i][1]) == canonical(end_label)
        ]
        if not ends:
            raise ValueError(f"CBL {area} section end not found.")
        end = ends[0]
        limit = end if area != "OPEX" else end + 1
        ytd = {}
        monthly = {}
        for row in fs_rows[start + 1 : limit]:
            raw_label = str(row[1] or "").strip()
            if not raw_label or not isinstance(row[2], (int, float)):
                continue
            label = re.sub(r"^\d+\s*", "", raw_label).strip()
            if area == "Revenue":
                if canonical(label).startswith("TOTAL "):
                    continue  # e.g. "Total Service Revenues" is a subtotal, not a line
                label = _REVENUE_STATEMENT_NAMES.get(canonical(label), label)
            if canonical(label) in {canonical(k) for k in ytd}:
                raise ValueError(f"Duplicate CBL {area} statement line: {label}")
            ytd[label] = amount(row[2], label)
            if area == "OPEX" and canonical(label) == "TOTAL OPERATING COSTS":
                earnings = report["Statement of Earnings"]
                matches = [
                    r
                    for r in earnings.iter_rows(values_only=True)
                    if canonical(r[1]) == "OPERATING COSTS"
                ]
                if len(matches) != 1:
                    raise ValueError("Expected one CBL monthly Operating costs total.")
                monthly[label] = amount(matches[0][3], label)
            elif area == "Direct Cost" and canonical(label) == "COST OF GOODS SOLD":
                matches = [
                    r
                    for r in report["Statement of Earnings"].iter_rows(values_only=True)
                    if canonical(r[1]) == "TOTAL COGS"
                ]
                if len(matches) != 1:
                    raise ValueError("Expected one CBL monthly Total COGS.")
                monthly[label] = amount(matches[0][3], label)
            elif area == "OPEX":
                code = "MD" if canonical(label) == "MD MEDIA" else raw_label.split()[0]
                matches = [
                    r
                    for r in rows
                    if str(r[1] or "").strip().casefold() == "opex"
                    and str(r[2]).strip().upper() == code.upper()
                ]
                if len(matches) != 1:
                    raise ValueError(
                        f"Expected one CBL Data OPEX code {code} for {label}."
                    )
                monthly[label] = amount(matches[0][month_column], label)
            else:
                data_start_label = (
                    "Revenue Breakdown for P&L Summary"
                    if area == "Revenue"
                    else "Cost of Goods Sold"
                )
                data_end_label = (
                    "Opex Breakdown for Income stmt mthly"
                    if area == "Revenue"
                    else "Gross Profit"
                )
                begins = [
                    i
                    for i, r in enumerate(rows)
                    if canonical(r[1]) == canonical(data_start_label)
                ]
                if len(begins) != 1:
                    raise ValueError(f"Expected one CBL Data {area} section.")
                stops = [
                    i
                    for i in range(begins[0] + 1, len(rows))
                    if canonical(rows[i][1]) == canonical(data_end_label)
                ]
                if not stops:
                    raise ValueError(f"CBL Data {area} section end is missing.")
                names = {canonical(label)} | set(
                    _REVENUE_DATA_NAMES.get(canonical(label), ())
                )
                matches = [
                    r
                    for r in rows[begins[0] + 1 : stops[0]]
                    if canonical(r[2]) in names
                ]
                if len(matches) != 1:
                    raise ValueError(f"Expected one CBL Data monthly line for {label}.")
                monthly[label] = amount(matches[0][month_column], label)
        if area == "Direct Cost":
            if (
                canonical(earnings["P6"].value) != "YEAR-TO-DATE"
                or canonical(earnings["P8"].value) != "ACTUAL"
            ):
                raise ValueError(
                    "CBL Statement of Earnings YTD actual headers do not match."
                )
            totals = [
                r
                for r in earnings.iter_rows(values_only=True)
                if canonical(r[1]) == "TOTAL COGS"
            ]
            if len(totals) != 1:
                raise ValueError("Expected one CBL reported Total COGS.")
            monthly["Cost of Goods Sold"] = amount(totals[0][3], "monthly Total COGS")
            ytd["Cost of Goods Sold"] = amount(totals[0][15], "YTD Total COGS")
        if area in ("Revenue", "Direct Cost"):
            # The Financials tab hides lines that are zero. The Statement of
            # Earnings still lists them, so a line that is 0 there reads as 0.
            from automations.actuals import mapping as _mapping

            wanted = {
                "Revenue": _mapping.REVENUE_MAPPINGS,
                "Direct Cost": _mapping.DIRECT_COST_MAPPINGS,
            }[area]
            have = {canonical(k) for k in ytd}
            for entry in wanted:
                if entry.get("company") != "CBL":
                    continue
                name = entry["fs_line"]
                if canonical(name) in have:
                    continue
                for r in earnings.iter_rows(values_only=True):
                    if canonical(r[1]) == canonical(name):
                        month_v = r[3] if len(r) > 3 else None
                        ytd_v = r[15] if len(r) > 15 else None
                        if (
                            isinstance(month_v, (int, float))
                            and isinstance(ytd_v, (int, float))
                            and month_v == 0
                            and ytd_v == 0
                        ):
                            ytd[name] = 0.0
                            monthly[name] = 0.0
                        break
        if not ytd:
            raise ValueError(f"No CBL {area} statement values found.")
        if area == "OPEX":
            # The statement names these lines differently from the report / Tableau.
            aliases = {
                "Contact Center": "Customer Experience & Contact Center",
                "MD Media": "Media",
                "Total Operating Costs": "Total Opex",
            }
            for source_name, report_name in aliases.items():
                for values in (monthly, ytd):
                    if source_name in values and report_name not in values:
                        values[report_name] = values[source_name]
        return {
            "MONTH": monthly,
            "YTD": ytd,
            "period": closed_period,
            "monthly_source": f"Data column {month_column+1}",
            "ytd_source": "Financials column C",
        }
    finally:
        report.close()
