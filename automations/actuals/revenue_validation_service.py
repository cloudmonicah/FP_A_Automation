"""Capture selected Revenue TWBs and validate monthly and YTD separately."""

from datetime import date, datetime
from pathlib import Path
from decimal import Decimal
from io import StringIO
from uuid import uuid4
import csv
import re
import time
import xml.etree.ElementTree as ET
from openpyxl import load_workbook

SOURCES = [
    (
        "consumer",
        "ALIV Consumer",
        "ALIV RESIDENTIAL REVENUE Detail - Monthly Actual v Budget YTD",
    ),
    ("consumer", "CBL Consumer", "CBL RESIDENTIAL REVENUE SUMMARY - Monthly Actuals"),
    (
        "corporate",
        "Corporate",
        "CBL & ALIV REVENUE SUMMARY - Month YTD Actual v Budget",
    ),
    ("group", "CBL Group", "CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY"),
    # Consolidated ALIV "Other" = OTHER REVENUE on the ALIV FS Rollup view.
    (
        "group",
        "ALIV Group Other",
        "ALIV FS Rollup View_Monthly Revenue Forecast_Summary",
    ),
    # Individual captures for CBL Government, Other and MSA-ALIV.
    (
        "corporate",
        "CBL Government",
        "CBL COMMERCIAL REVENUE by Sub Account Description - Monthly Actual v Budget",
    ),
    ("finance", "CBL Finance View", "CBL Monthly Revenue Forecast"),
    ("finance", "ALIV Finance View", "ALIV Monthly Revenue Forecast"),
    ("media", "CBL Media", "Monthly YTD Actual v Budget by Sub account"),
]

# Media's YTD comes from a different worksheet than its Monthly capture.
MEDIA_YTD_TAB = "Year to Date Summary"


def resolve_worksheet(path, wanted):
    """Exact worksheet name in the .twb: exact, then case-insensitive, then
    a name that ends with / contains ``wanted`` (e.g. 'MEDIA - Year to Date Summary')."""
    names = [
        w.get("name", "")
        for w in ET.parse(path).getroot().findall("./worksheets/worksheet")
    ]
    key = " ".join(wanted.lower().split())
    for test in (
        lambda n: n == wanted,
        lambda n: " ".join(n.lower().split()) == key,
        lambda n: " ".join(n.lower().split()).endswith(key),
        lambda n: key in " ".join(n.lower().split()),
    ):
        hits = [n for n in names if test(n)]
        if len(hits) == 1:
            return hits[0]
    raise ValueError(
        f"{Path(path).name} has no unique worksheet matching '{wanted}'. "
        f"Worksheets: {', '.join(names)}"
    )


# Sources that are read as the INDIVIDUAL value only (the same statement line
# is also read from the Group rollup as the consolidated value).
INDIVIDUAL_ONLY_LABELS = {
    "CBL Government",
    "CBL Finance View",
    "ALIV Finance View",
    "CBL Media",
}

# Sources that may be missing without stopping the whole run.
OPTIONAL_KEYS = {"finance", "media"}

# Companies a failed capture affects.
AFFECTED_COMPANIES = {
    "ALIV Consumer": ("ALIV",),
    "CBL Consumer": ("CBL",),
    "Corporate": ("ALIV", "CBL"),
    "CBL Group": ("CBL",),
    "ALIV Group Other": ("ALIV",),
    "CBL Government": ("CBL",),
    "CBL Finance View": ("CBL",),
    "ALIV Finance View": ("ALIV",),
    "CBL Media": ("CBL",),
}


def _lookup(company, fs_line):
    """The row-lookup dictionary configured in mapping.py for one line."""
    from automations.actuals.mapping import INDIVIDUAL_REVENUE_MAPPINGS

    for route in INDIVIDUAL_REVENUE_MAPPINGS:
        if route["company"] == company and route["fs_line"] == fs_line:
            if not route.get("lookup"):
                raise ValueError(f"No Tableau row lookup configured for {company} / {fs_line}.")
            return dict(route["lookup"])
    raise ValueError(f"No individual Revenue mapping for {company} / {fs_line}.")


def canonical(value):
    return " ".join(str(value or "").strip().upper().split())


def number(value):
    text = str(value or "").strip().replace(",", "").replace("$", "")
    if not text:
        raise ValueError(
            "A mapped Revenue actual is blank; it cannot be treated as zero."
        )
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]
    result = Decimal(text)
    if not result.is_finite():
        raise ValueError("Invalid Revenue actual.")
    return float(result)


_PERIOD_TEXT_FORMATS = (
    "%B %Y",
    "%b %Y",
    "%b-%y",
    "%B-%y",
    "%b-%Y",
    "%Y-%m-%d",
    "%Y-%m",
    "%m/%d/%Y",
    "%m/%d/%y",
    "%d-%b-%y",
    "%d-%b-%Y",
)


def statement_year_month(value):
    """(year, month) from a date, an Excel serial number or period text; else None."""
    if isinstance(value, (datetime, date)):
        return value.year, value.month
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 < value < 80000:  # looks like an unformatted Excel date serial
            from openpyxl.utils.datetime import from_excel

            found = from_excel(value)
            return found.year, found.month
        return None
    text = " ".join(str(value or "").split())
    for fmt in _PERIOD_TEXT_FORMATS:
        try:
            found = datetime.strptime(text, fmt)
        except ValueError:
            continue
        return found.year, found.month
    return None


def find_statement_period(sheet, preferred="D1", scan_rows=6, scan_cols=10):
    """Reporting month of a statement sheet as ((year, month) | None, cell, raw value).

    Uses the usual cell first; if it is empty or unreadable, falls back to the first
    readable month in the top rows. A mismatch is still an error for the caller, so
    picking the wrong month's file can never pass silently.
    """
    raw = sheet[preferred].value
    found = statement_year_month(raw)
    if found:
        return found, preferred, raw
    for row in sheet.iter_rows(min_row=1, max_row=scan_rows, max_col=scan_cols):
        for item in row:
            found = statement_year_month(item.value)
            if found:
                return found, item.coordinate, item.value
    return None, preferred, raw


def prepare_revenue_copy(source, destination, tabs, closed_period, basis):
    """Set the period filters and write the workbook to ``destination``.

    ``destination`` may be the source itself: the Revenue flow edits the
    source workbook in place and saves it, so no temporary copy is made.
    Returns the filter values that were set (recorded in the crosstab workbook).
    """
    root = ET.parse(source).getroot()
    start = date(
        closed_period.year + (closed_period.month == 12),
        closed_period.month % 12 + 1,
        1,
    )
    fiscal_start = date(closed_period.year - (closed_period.month < 7), 7, 1)
    lower = (
        fiscal_start
        if basis == "YTD"
        else date(closed_period.year, closed_period.month, 1)
    )
    actual_names = set()
    for tab in tabs:
        sheet = root.find(f'./worksheets/worksheet[@name="{tab}"]')
        if sheet is None:
            raise ValueError(f"Revenue worksheet not found: {tab}")
        date_filters = [
            f
            for f in sheet.iter("filter")
            if any(
                f.get("column", "").endswith(f"[{level}:GL Posting Date:ok]")
                for level in ("my", "tmn")
            )
        ]
        if not date_filters:
            raise ValueError(
                f"No mapped posting-date filter on Revenue worksheet: {tab}"
            )
        actual_names.update(
            c.get("name")
            for c in sheet.iter("column")
            if c.get("caption") == "Actual YTD"
        )
    parameters = [
        c
        for c in root.iter("column")
        if canonical(c.get("caption")) == "FORECAST START"
        and "param-domain-type" in c.attrib
    ]
    if not parameters or len({c.get("name") for c in parameters}) != 1:
        raise ValueError("Expected one Revenue Forecast Start parameter name.")
    for parameter in parameters:
        parameter.set("value", f"#{start:%Y-%m-%d}#")
        if parameter.find("calculation") is not None:
            parameter.find("calculation").set("formula", f"#{start:%Y-%m-%d}#")
    formula = f"IF [GL Posting Date] >= #{lower:%Y-%m-%d}# AND [GL Posting Date] < #{start:%Y-%m-%d}# THEN [Actual] ELSE 0 END"
    for c in root.iter("column"):
        if c.get("name") in actual_names and c.find("calculation") is not None:
            c.find("calculation").set("formula", formula)
    months = []
    current = lower
    while current < start:
        months.append(current)
        current = date(current.year + (current.month == 12), current.month % 12 + 1, 1)
    for f in root.iter("filter"):
        column = f.get("column", "")
        for level in ("my", "tmn"):
            if column.endswith(f"[{level}:GL Posting Date:ok]"):
                for child in list(f):
                    f.remove(child)
                union = ET.SubElement(f, "groupfilter", function="union")
                for month in months:
                    member = (
                        f"{month:%Y%m}"
                        if level == "my"
                        else f"#{month:%Y-%m-%d} 00:00:00#"
                    )
                    ET.SubElement(
                        union,
                        "groupfilter",
                        function="member",
                        level=f"[{level}:GL Posting Date:ok]",
                        member=member,
                    )
        if column.endswith("[none:Fiscal Year:ok]"):
            for child in list(f):
                f.remove(child)
            ET.SubElement(
                f,
                "groupfilter",
                function="member",
                level="[none:Fiscal Year:ok]",
                member=str(fiscal_start.year + 1),
            )
        if column.endswith("[none:Forecast Type:nk]"):
            for member in f.findall('.//groupfilter[@function="member"]'):
                if re.fullmatch(
                    r"FY\d{2} Forecast_\d{6}", member.get("member", "").strip('"'), re.I
                ):
                    member.set(
                        "member",
                        f'"FY{(fiscal_start.year+1)%100:02} Forecast_{start:%m%Y}"',
                    )
    ET.ElementTree(root).write(destination, encoding="utf-8", xml_declaration=True)
    return {
        "Basis": basis,
        "Forecast Start": f"{start:%m/%d/%Y}",
        "Forecast Type": f"FY{(fiscal_start.year + 1) % 100:02} Forecast_{start:%m%Y}",
        "GL Posting Date months": (
            f"{months[0]:%b %Y} to {months[-1]:%b %Y} ({len(months)} month(s))"
            if months
            else "none"
        ),
        "Fiscal Year": str(fiscal_start.year + 1),
        "Actual window": f"{lower:%m/%d/%Y} up to {start:%m/%d/%Y}",
    }


def _month_actual_columns(headers, closed_period, basis):
    """Actual columns, one per requested month, or None if any is missing."""
    current = (
        date(closed_period.year - (closed_period.month < 7), 7, 1)
        if basis == "YTD"
        else date(closed_period.year, closed_period.month, 1)
    )
    months = []
    while current <= closed_period:
        months.append(current)
        current = date(current.year + (current.month == 12), current.month % 12 + 1, 1)
    by_key = {canonical(h): h for h in headers}
    columns = []
    for month in months:
        found = None
        for label in (month.strftime("%b-%y"), month.strftime("%B %Y")):
            for key in (
                label.upper(),
                f"{label.upper()} ACTUAL",
                f"{label.upper()} ACTUAL YTD",
            ):
                if key in by_key:
                    found = by_key[key]
                    break
            if found:
                break
        if not found:
            return None
        columns.append(found)
    return columns


def parse_capture(text, basis, closed_period):
    """Flatten one or two header rows and restore suppressed hierarchy labels."""
    raw = [
        r
        for r in csv.reader(StringIO(text), delimiter="\t")
        if any(v.strip() for v in r)
    ]
    if len(raw) < 2:
        raise ValueError("Revenue capture is empty.")
    first = raw[0]
    two_headers = any(
        canonical(v)
        in (
            "ACTUAL YTD",
            "BUDGET YTD",
            "VAR YTD $",
            "VAR YTD %",
            "ACTUAL",
            "BUDGET",
            "VAR $",
            "VAR %",
        )
        for v in raw[1]
    )
    if two_headers:
        from automations.actuals.direct_cost_validation_service import (
            _parse_direct_cost_crosstab_text,
        )

        rows = _parse_direct_cost_crosstab_text(text)
    else:
        if len(first) != len(set(first)):
            raise ValueError("Revenue capture has duplicate headers.")
        rows = [dict(zip(first, r + [""] * (len(first) - len(r)))) for r in raw[1:]]
    headers = list(rows[0])
    normalized = {canonical(h): h for h in headers}
    target = "GRAND TOTAL ACTUAL YTD" if two_headers else "GRAND TOTAL"
    if target not in normalized:
        if "ACTUAL YTD" in normalized:
            target = "ACTUAL YTD"
        elif basis == "MONTH":
            labels = [
                closed_period.strftime("%b-%y").upper(),
                closed_period.strftime("%B %Y").upper(),
            ]
            target = next((label for label in labels if label in normalized), target)
    month_columns = None
    if target not in normalized:
        # Months-as-columns sheets (no Grand Total): the value is the sum of
        # the month columns for the period (MONTH = closed month; YTD = every
        # fiscal month from July through the closed month).
        month_columns = _month_actual_columns(headers, closed_period, basis)
        if not month_columns:
            raise ValueError(f"Revenue actual column is missing: {target}")
    # Period filters were set on the temporary TWB before capture. Check visible month headers too.
    visible = set()
    for h in headers:
        for pattern, fmt in [
            (r"^([A-Za-z]+ \d{4})(?: |$)", "%B %Y"),
            (r"^([A-Za-z]{3}-\d{2})(?: |$)", "%b-%y"),
        ]:
            match = re.match(pattern, str(h))
            if match:
                try:
                    visible.add(datetime.strptime(match[1], fmt).date())
                except ValueError:
                    pass
    expected = set()
    current = (
        date(closed_period.year - (closed_period.month < 7), 7, 1)
        if basis == "YTD"
        else date(closed_period.year, closed_period.month, 1)
    )
    while current <= closed_period:
        expected.add(current)
        current = date(current.year + (current.month == 12), current.month % 12 + 1, 1)
    if visible and visible != expected:
        raise ValueError(
            "Revenue capture month headers do not match the requested period."
        )
    dimensions = [h for h in headers if h.startswith("GL ") or h == "Company"]
    last = {h: "" for h in dimensions}
    for row in rows:
        for h in dimensions:
            value = str(row.get(h) or "").strip()
            if value:
                if value != last[h]:
                    for lower in dimensions[dimensions.index(h) + 1 :]:
                        last[lower] = ""
                last[h] = value
            else:
                row[h] = last[h]
        if month_columns:
            parts = [row.get(h) for h in month_columns]
            if all(not str(v or "").strip() for v in parts):
                row["_actual"] = ""
            else:
                # A month with no postings is an empty cell in Tableau; it adds nothing.
                row["_actual"] = sum(number(v) for v in parts if str(v or "").strip())
        else:
            row["_actual"] = row.get(normalized[target])
    return rows


def unique(rows, **dimensions):
    matches = [
        r
        for r in rows
        if all(canonical(r.get(k)) == canonical(v) for k, v in dimensions.items())
    ]
    if len(matches) != 1:
        raise ValueError(
            f"Expected one Revenue subtotal for {dimensions}; found {len(matches)}."
        )
    return number(matches[0]["_actual"])


def source_values(label, rows):
    if label == "ALIV Consumer":
        types = {
            "Consumer Prepaid": "PAY IN ADVANCE",
            "Consumer Postpaid": "POSTPAID",
            "Visitor Roaming": "VISITOR ROAMING",
            "Interconnect": "INTERCONNECT REVENUE",
        }
        return {
            "ALIV": {
                line: unique(
                    rows,
                    **{
                        "Company": "ALIV",
                        "GL FS Type": kind,
                        "GL FS IS Breakdown": "Total",
                    },
                )
                for line, kind in types.items()
            }
        }
    if label == "CBL Consumer":
        return {
            "CBL": {
                "Residential": unique(
                    rows,
                    **{
                        "Company": "CBL",
                        "GL Bus Code Name": "Total",
                        "GL Account Mapping": "Total",
                    },
                )
            }
        }
    if label == "Corporate":
        aliv = unique(rows, **{"Company": "ALIV", "GL Account Mapping": "Total"})
        total = unique(rows, **{"Company": "CBL", "GL Account Mapping": "Total"})
        oneoff = unique(
            rows, **{"Company": "CBL", "GL Account Mapping": "CORPORATE ONE OFF"}
        )
        return {
            "ALIV": {"Corporate": aliv},
            "CBL": {
                "Corporate - Recurring": total - oneoff,
                "Corporate - One off": oneoff,
            },
        }
    if label == "CBL Government":
        return {"CBL": {"Government": unique(rows, **_lookup("CBL", "Government"))}}
    if label == "CBL Finance View":
        return {
            "CBL": {
                line: unique(rows, **_lookup("CBL", line))
                for line in ("Other", "MSA-ALIV")
            }
        }
    if label == "ALIV Finance View":
        return {"ALIV": {"Other": unique(rows, **_lookup("ALIV", "Other"))}}
    if label == "CBL Media":
        dims = _lookup("CBL", "Media")
        detail = [
            r
            for r in rows
            if canonical(r.get("Company")) != "GRAND TOTAL"
            and all(canonical(r.get(k)) == canonical(v) for k, v in dims.items())
        ]
        if not detail:
            raise ValueError(f"No Media rows found for {dims}.")
        total = sum(number(r["_actual"]) for r in detail)
        grand = [r for r in rows if canonical(r.get("Company")) == "GRAND TOTAL"]
        # Each displayed cell is rounded to a whole dollar, so allow $1 per cell.
        if len(grand) == 1 and abs(number(grand[0]["_actual"]) - total) > max(
            1.0, len(detail) * 2.0
        ):
            raise ValueError(
                f"Media rows add to {total:,.2f} but Grand Total is "
                f"{number(grand[0]['_actual']):,.2f}; check the Media lookup."
            )
        return {"CBL": {"Media": total}}
    if label == "ALIV Group Other":
        return {
            "ALIV": {
                "Other": unique(
                    rows,
                    **{
                        "Company": "ALIV",
                        "GL FS Revenue by Type": "OTHER REVENUE",
                        "GL Account Mapping": "Total",
                    },
                )
            }
        }
    # CBL Group rollup (sheet "CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY").
    # This sheet has only two dimension columns, so a route is (Bus Code, FS Category).
    # It has no "GL Account Mapping" column (that exists on the Consumer/Corporate sheets).
    routes = {
        "Government": ("ALIV BUSINESS", "CORPORATE GOVERNMENT"),
        "Other": ("HFC FIXED NETWORK CONSUMER", "OTHER SUBSCRIBER"),
        "MSA-ALIV": ("INTERCOMPANY", "OTHER SUBSCRIBER"),
        "Media": ("MEDIA", "CORPORATE POST PAID RECURRING"),
        "Gross Revenue": ("Grand Total", "Total"),
    }
    return {
        "CBL": {
            line: unique(
                rows,
                **dict(
                    zip(
                        ["GL Bus Code Name", "GL FS Category", "GL Account Mapping"],
                        route,
                    )
                ),
            )
            for line, route in routes.items()
        }
    }


def aliv_statement_periods(path, closed_period):
    from automations.actuals.revenue_financial_statement_reader import (
        ALIV_REVENUE_LINES,
        _canonical_statement_line,
    )

    name = Path(path).name
    report = load_workbook(path, read_only=True, data_only=True)
    try:
        if "P & L_" not in report.sheetnames:
            raise ValueError(
                f"ALIV statement {name} has no 'P & L_' sheet; found {report.sheetnames}."
            )
        sheet = report["P & L_"]
        from automations.actuals.statement_columns import locate_aliv_columns

        rows = list(sheet.iter_rows(values_only=True))
        # Columns are found from the headers: the ALIV statement layout changes
        # between months, so no fixed cell or column letter is assumed.
        _header, month_col, ytd_col = locate_aliv_columns(rows, closed_period, name)
        result = {"MONTH": {}, "YTD": {}}
        for row in rows:
            key = _canonical_statement_line(row[1] if len(row) > 1 else None)
            month = row[month_col] if len(row) > month_col else None
            if key in ALIV_REVENUE_LINES and isinstance(month, (int, float)):
                label = ALIV_REVENUE_LINES[key]
                if label in result["MONTH"]:
                    raise ValueError(f"Duplicate ALIV Revenue statement label: {label}")
                result["MONTH"][label] = number(month)
                # July is month 1 of the fiscal year: YTD equals the monthly actual.
                ytd = month if closed_period.month == 7 else row[ytd_col]
                result["YTD"][label] = number(ytd)
                if canonical(label) == "SERVICE REVENUES":
                    break
        missing = set(ALIV_REVENUE_LINES.values()) - set(result["MONTH"])
        if missing:
            raise ValueError(f"ALIV Revenue statement lines missing: {sorted(missing)}")
        return result
    finally:
        report.close()


def _clear_open_menus():
    """Close any Tableau menu or dialog a failed attempt left open.

    A menu left open by one failed capture blocks every later capture, which
    is what turned a single failure into an all-PENDING report.
    """
    try:
        from pywinauto.keyboard import send_keys

        send_keys("{ESC}{ESC}")
        time.sleep(0.5)
    except Exception:
        pass


def run_live_revenue_validation(
    *,
    aliv_fs_path,
    cbl_fs_path,
    tableau_sources,
    output_paths,
    closed_period,
    capture=None,
    opener=None,
    closer=None,
    wait_seconds=8,
):
    from automations.actuals.cbl_statement_periods import read_cbl_area
    from automations.actuals.revenue_validation_service import compare_revenue_company
    from automations.actuals.actuals_report_layout import write_area_workbook
    from automations.actuals.actuals_report_layout import (
        add_company_columns,
        status_cell,
        refresh_summary,
        report_status,
    )
    from automations.actuals.crosstab_audit import append_crosstab

    if capture is None:
        from automations.actuals.tableau_desktop_export_service import (
            get_tableau_crosstab_text,
        )

        capture = get_tableau_crosstab_text
    if opener is None:
        from automations.forecast.tableau_service import open_tableau_workbook

        opener = open_tableau_workbook
    if closer is None:
        from automations.actuals.tableau_desktop_close_service import (
            close_tableau_forecast_workbook,
        )

        closer = close_tableau_forecast_workbook
    # Validate all selected files/tabs before opening Tableau. Optional sources
    # (Finance View) that are missing only leave their lines PENDING.
    notes = {basis: {"ALIV": [], "CBL": []} for basis in ("MONTH", "YTD")}
    unavailable = set()
    for key, label, tab in SOURCES:
        path = Path(tableau_sources.get(key) or "")
        problem = None
        if not path.is_file() or path.suffix.lower() != ".twb":
            problem = f"Select the Revenue {key} TWB."
        elif (
            ET.parse(path).getroot().find(f'./worksheets/worksheet[@name="{tab}"]')
            is None
        ):
            problem = f"{path.name} does not contain the mapped Revenue tab: {tab}"
        if not problem:
            continue
        if key not in OPTIONAL_KEYS and label not in INDIVIDUAL_ONLY_LABELS:
            raise ValueError(problem)
        unavailable.add(label)
        for basis in notes:
            for company in AFFECTED_COMPANIES[label]:
                notes[basis][company].append(f"{label}: {problem}")
    statements = {
        "ALIV": aliv_statement_periods(aliv_fs_path, closed_period),
        "CBL": read_cbl_area(cbl_fs_path, "Revenue", closed_period),
    }
    tag = closed_period.strftime("%Y-%m")
    folder = Path(output_paths["revenue"])
    folder.mkdir(parents=True, exist_ok=True)
    raw = (
        Path(output_paths["tableau_raw_copies"]) / f"Revenue_Actuals_{uuid4().hex[:8]}"
    )
    raw.mkdir(parents=True, exist_ok=True)
    audit = raw / "Revenue_Actuals_Crosstabs.xlsx"
    values = {basis: {"ALIV": {}, "CBL": {}} for basis in ("MONTH", "YTD")}
    individual_values = {basis: {"ALIV": {}, "CBL": {}} for basis in values}
    sequence = 0
    filters_used = {}
    for basis in ("MONTH", "YTD"):
        for key in ("consumer", "corporate", "group", "finance", "media"):
            source = Path(tableau_sources[key])
            targets = [
                (label, tab)
                for k, label, tab in SOURCES
                if k == key and label not in unavailable
            ]
            if not targets:
                continue
            if basis == "YTD":
                try:
                    targets = [
                        (label, resolve_worksheet(source, MEDIA_YTD_TAB))
                        if label == "CBL Media"
                        else (label, tab)
                        for label, tab in targets
                    ]
                except Exception as exc:
                    for label, _ in targets:
                        for company in AFFECTED_COMPANIES[label]:
                            notes[basis][company].append(
                                f"{label}: {type(exc).__name__}: {exc}"
                            )
                    continue
            opened = False
            try:
                # Set this month's filters directly in the source workbook and
                # save it; no temporary copy is made.
                try:
                    filters = prepare_revenue_copy(
                        source, source, [tab for _, tab in targets], closed_period, basis
                    )
                except Exception as exc:
                    if key not in OPTIONAL_KEYS:
                        raise
                    for label, _ in targets:
                        for company in AFFECTED_COMPANIES[label]:
                            notes[basis][company].append(
                                f"{label}: {type(exc).__name__}: {exc}"
                            )
                    continue
                filters_used[(basis, source.name)] = filters
                print(f"Revenue {basis}: saved {source.name} | {filters}", flush=True)
                opener(tableau_workbook_path=str(source))
                opened = True
                if wait_seconds:
                    time.sleep(wait_seconds)
                for label, tab in targets:
                    sequence += 1
                    print(f"Revenue {basis}: {label} | {tab}", flush=True)
                    try:
                        _clear_open_menus()
                        try:
                            text = capture(
                                worksheet_name=tab, tableau_workbook_path=str(source)
                            )
                        except RuntimeError as first_error:
                            # One retry: clear menus, give Tableau time to
                            # finish loading the filtered view, then try again.
                            print(
                                f"Revenue {basis}: {label} capture failed "
                                f"({first_error}); retrying once.",
                                flush=True,
                            )
                            _clear_open_menus()
                            time.sleep(max(wait_seconds, 5))
                            text = capture(
                                worksheet_name=tab, tableau_workbook_path=str(source)
                            )
                        rows = parse_capture(text, basis, closed_period)
                        company_tag = "+".join(AFFECTED_COMPANIES[label])
                        append_crosstab(
                            audit,
                            [
                                {k: v for k, v in record.items() if not k.startswith("_")}
                                for record in rows
                            ],
                            source=source,
                            tab=tab,
                            company=company_tag,
                            area="Revenue",
                            basis=basis,
                            sequence=sequence,
                            filters={**filters, "Capture": label},
                        )
                        parsed = source_values(label, rows)
                        for company, items in parsed.items():
                            if label not in INDIVIDUAL_ONLY_LABELS:
                                if set(items) & set(values[basis][company]):
                                    raise ValueError(
                                        "Duplicate Revenue mapping across sources."
                                    )
                                values[basis][company].update(items)
                            # Consumer, Corporate and Finance View are independent
                            # detail sources. Group rollup is NOT an individual source.
                            if key in ("consumer", "corporate", "finance", "media"):
                                individual_values[basis][company].update(items)
                    except Exception as exc:
                        _clear_open_menus()
                        for company in AFFECTED_COMPANIES[label]:
                            notes[basis][company].append(
                                f"{label}: {type(exc).__name__}: {exc}"
                            )
            finally:
                if opened:
                    closer(tableau_workbook_path=str(source))
    print("Revenue: all Tableau captures finished; preparing comparisons.", flush=True)
    for basis in ("MONTH", "YTD"):
        for company in ("ALIV", "CBL"):
            if notes[basis][company]:
                print(
                    f"Revenue {basis} {company} capture warnings: "
                    + " | ".join(notes[basis][company]),
                    flush=True,
                )
    compared = {
        basis: {
            company: compare_revenue_company(
                company=company,
                fs_values=statements[company][basis],
                tableau_values=values[basis][company],
            )
            for company in ("ALIV", "CBL")
        }
        for basis in values
    }
    for basis in values:
        for company in ("ALIV", "CBL"):
            for item in compared[basis][company]["details"]:
                if item["status"] == "ACTION REQUIRED":
                    item["note"] = (
                        "; ".join(notes[basis][company])
                        or "Revenue subtotal mapping is not confirmed for this statement line."
                    )
    path = folder / f"Revenue_Actuals_Validation_{tag}.xlsx"
    start = date(
        closed_period.year + (closed_period.month == 12),
        closed_period.month % 12 + 1,
        1,
    )
    print(f"Revenue: writing validation workbook: {path}", flush=True)
    write_area_workbook(
        area="Revenue",
        output_file=path,
        closed_period=closed_period,
        forecast_start=start,
        aliv_result=compared["MONTH"]["ALIV"],
        cbl_result=compared["MONTH"]["CBL"],
    )
    print(f"Revenue: base validation workbook saved: {path}", flush=True)
    report = load_workbook(path)
    try:
        for company in ("ALIV", "CBL"):
            sheet = report[f"{company} Revenue"]
            add_company_columns(sheet)
            lookup = {
                item["statement_line"]: item
                for item in compared["YTD"][company]["details"]
            }
            for row in range(7, sheet.max_row + 1):
                statement_line = sheet.cell(row, 1).value

                if statement_line not in lookup:
                    continue

                item = lookup[statement_line]

                for col, key in (
                    (8, "fs_rounded_000"),
                    (9, "tableau_rounded_000"),
                    (10, "variance_000"),
                ):
                    sheet.cell(row, col, item.get(key))

                status_cell(
                    sheet.cell(row, 11),
                    item["status"],
                )

                if item["status"] != "PASS":
                    from openpyxl.comments import Comment

                    sheet.cell(row, 11).comment = Comment(
                        item.get("note", "Revenue validation requires review."),
                        "Validation",
                    )
        # Merge independently captured detail values into the approved report.
        # Never substitute consolidated values for missing individual captures.
        from automations.actuals.actuals_report_layout import (
            canonical as report_canonical,
            rounded as report_rounded,
        )
        from automations.actuals.report_builder import PeriodValues
        from automations.actuals.report_builder import variance as report_variance
        from openpyxl.comments import Comment

        for company in ("ALIV", "CBL"):
            sheet = report[f"{company} Revenue"]
            # This is the capture (A:S) workbook; the final layout is built later.
            for row in range(7, sheet.max_row + 1):
                label = report_canonical(sheet.cell(row, 1).value)
                if not label:
                    continue
                for basis in ("MONTH", "YTD"):
                    matched = [
                        (name, value)
                        for name, value in individual_values[basis][company].items()
                        if report_canonical(name) == label
                    ]
                    if len(matched) != 1 or matched[0][1] is None:
                        continue
                    # Captured Tableau amounts are dollars; report is $000.
                    amount_000 = report_rounded(
                        Decimal(str(matched[0][1])) / Decimal("1000")
                    )
                    (
                        fs_col,
                        group_col,
                        individual_col,
                        fs_diff_col,
                        group_diff_col,
                        status_col,
                    ) = (
                        (2, 3, 12, 13, 14, 15)
                        if basis == "MONTH"
                        else (8, 9, 16, 17, 18, 19)
                    )
                    sheet.cell(row, individual_col, amount_000)
                    fs_value = sheet.cell(row, fs_col).value
                    group_value = sheet.cell(row, group_col).value
                    sheet.cell(row, fs_diff_col, report_variance(fs_value, amount_000))
                    sheet.cell(
                        row, group_diff_col, report_variance(group_value, amount_000)
                    )
                    status_cell(
                        sheet.cell(row, status_col),
                        PeriodValues(fs_value, group_value, amount_000).status,
                    )
                # Lines without independent detail remain NOT VALIDATED/PENDING.
        refresh_summary(report)
        report.save(path)
    finally:
        report.close()
    print(f"Revenue: validation workbook saved successfully: {path}", flush=True)
    status = "REVIEW REQUIRED"  # Revenue mappings are not yet fully verified.
    log = (
        Path(output_paths["logs"])
        / f"Revenue_Validation_{tag}_{datetime.now():%Y%m%d_%H%M%S}.log"
    )
    log.parent.mkdir(parents=True, exist_ok=True)
    log.write_text(
        "\n".join(
            [
                f"Revenue actuals: {tag}",
                f"Overall: {status}",
                f"Report: {path}",
                f"Crosstabs: {audit}",
            ]
            + [
                f"{basis} {company}: {note}"
                for basis in notes
                for company in notes[basis]
                for note in notes[basis][company]
            ]
        ),
        encoding="utf-8",
    )
    return {
        "area": "Revenue",
        "overall_status": status,
        "validation_workbook": str(path),
        "raw_tableau_workbook": str(audit),
        "log_file": str(log),
        "forecast_start": start,
        "forecast_type": next(iter(filters_used.values()), {}).get("Forecast Type"),
        "filters_used": [
            {"Basis": b, "Workbook": w, **f} for (b, w), f in filters_used.items()
        ],
        "capture_notes": [
            {"Basis": basis, "Company": company, "Note": note}
            for basis in notes
            for company in notes[basis]
            for note in notes[basis][company]
        ],
        "companies": compared["MONTH"],
        "ytd_companies": compared["YTD"],
    }


def compare_revenue_company(*, company, fs_values, tableau_values):
    """Conservative Revenue comparison using the shared report schema."""
    from automations.actuals.mapping import REVENUE_MAPPINGS
    from decimal import Decimal, ROUND_HALF_UP

    def as_000(value):
        if value is None:
            return None
        return int(
            Decimal(str(value)).quantize(
                Decimal("1"),
                rounding=ROUND_HALF_UP,
            )
        )

    details = []

    for mapping in REVENUE_MAPPINGS:
        if canonical(mapping.get("company")) != canonical(company):
            continue

        line = mapping["fs_line"]
        fs_raw = fs_values.get(line)
        tableau_raw = tableau_values.get(line)

        fs_rounded = as_000(fs_raw)

        tableau_rounded = (
            as_000(Decimal(str(tableau_raw)) / Decimal("1000"))
            if tableau_raw is not None
            else None
        )

        difference = (
            tableau_rounded - fs_rounded
            if fs_rounded is not None and tableau_rounded is not None
            else None
        )

        status = (
            "ACTION REQUIRED"
            if fs_raw is None or tableau_raw is None
            else "REVIEW REQUIRED"
        )

        details.append(
            {
                "company": company,
                "statement_line": line,
                "tableau_category": line,
                "fs_raw_000": fs_raw,
                "fs_rounded_000": fs_rounded,
                "tableau_raw_dollars": tableau_raw,
                "tableau_rounded_000": tableau_rounded,
                "variance_000": difference,
                "status": status,
                "note": (
                    "Missing Financial Statement or Tableau value."
                    if status == "ACTION REQUIRED"
                    else "Revenue mapping and period amounts require "
                    "confirmation; no automatic PASS."
                ),
            }
        )

    summary = {
        "total": len(details),
        "passed": sum(row["status"] == "PASS" for row in details),
        "failed": sum(row["status"] == "FAIL" for row in details),
        "action_required": sum(row["status"] == "ACTION REQUIRED" for row in details),
    }

    return {
        "company": company,
        "status": "REVIEW REQUIRED",
        "overall_status": "REVIEW REQUIRED",
        "summary": summary,
        "details": details,
    }


def validate_revenue_workbook_path(revenue_workbook_path):
    if not revenue_workbook_path:
        raise ValueError("Please select the consolidated Revenue Tableau workbook.")

    workbook = Path(revenue_workbook_path)

    if not workbook.is_file():
        raise FileNotFoundError(f"Revenue Tableau workbook not found: {workbook}")

    if workbook.suffix.lower() not in (".twb", ".twbx"):
        raise ValueError("Revenue Tableau workbook must be a .twb or .twbx file.")

    return str(workbook)
