"""Mapped ALIV OPEX monthly/YTD tests, separate from forecast runs."""

from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime
from uuid import uuid4
import time
import re
import xml.etree.ElementTree as ET

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

from automations.actuals.validation_config import (
    INDIVIDUAL_OPEX_ACTUALS_MAPPINGS,
    OPEX_MAPPINGS,
)


# Workbook and tab paths come from each corresponding OPEX mapping.
def _route_for(mapping):
    matches = [
        route
        for route in OPEX_MAPPINGS
        if all(
            route.get(key) == mapping[key]
            for key in (
                "company",
                "fs_line",
                "tableau_category",
            )
        )
    ]

    if len(matches) != 1:
        raise ValueError(
            f"Expected one OPEX mapping for "
            f"{mapping['company']} / {mapping['fs_line']}"
        )

    return matches[0]


MAPPING = INDIVIDUAL_OPEX_ACTUALS_MAPPINGS[0]
SOURCE = _route_for(MAPPING)["individual_tableau_workbook"]
TAB = _route_for(MAPPING)["individual_tableau_worksheet"]


def _number(value):
    if value is None or isinstance(value, bool) or not str(value).strip():
        raise ValueError("Actual amount is missing.")

    text = str(value).strip().replace(",", "").replace("$", "")

    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]

    result = Decimal(text)

    if not result.is_finite():
        raise ValueError("Actual amount is invalid.")

    return result


def _rounded(value):
    return int(
        value.quantize(
            Decimal("1"),
            rounding=ROUND_HALF_UP,
        )
    )


def _actual(rows, mapping=MAPPING):
    matches = []

    for row in rows:
        company = [
            value
            for key, value in row.items()
            if str(key).strip().casefold() == "company"
        ]

        if len(company) == 1 and str(company[0]).strip().upper() == mapping["company"]:
            fields = [
                value
                for key, value in row.items()
                if str(key).strip().casefold()
                == mapping["actual_measure"].strip().casefold()
            ]

            if len(fields) != 1:
                raise ValueError(
                    "Expected one Actual YTD column " "in the summary capture."
                )

            matches.append(_number(fields[0]))

    if len(matches) != 1:
        raise ValueError(
            "Expected one ALIV summary row; " "no subtotal summation is permitted."
        )

    return matches[0]


def _prepare(
    source,
    destination,
    closed_period,
    basis,
    mapping=MAPPING,
    tab=TAB,
    tableau_settings=None,
):
    """Prepare parameters and period on an isolated workbook copy."""

    tree = ET.parse(source)
    root = tree.getroot()

    sheet = root.find(f"./worksheets/worksheet[@name='{tab}']")

    if sheet is None:
        raise ValueError(f"Worksheet not found: {tab}")

    start = date(
        closed_period.year + (closed_period.month == 12),
        closed_period.month % 12 + 1,
        1,
    )

    tableau_settings = tableau_settings or {}

    selected_start = tableau_settings.get("forecast_start")
    if selected_start is not None:
        if isinstance(selected_start, str):
            selected_start = date.fromisoformat(selected_start)
        if not isinstance(selected_start, date):
            raise ValueError("Invalid GUI Forecast Start")
        start = selected_start

    fiscal_start_month = mapping.get("fiscal_start_month", 7)

    fiscal_start = date(
        closed_period.year - (closed_period.month < fiscal_start_month),
        fiscal_start_month,
        1,
    )

    lower = (
        fiscal_start
        if basis == "YTD"
        else date(
            closed_period.year,
            closed_period.month,
            1,
        )
    )

    fiscal_end = fiscal_start.year + 1

    actuals_end = date(
        closed_period.year + (closed_period.month == 12),
        closed_period.month % 12 + 1,
        1,
    )

    forecast_type = tableau_settings.get("forecast_type")
    if forecast_type is None:
        forecast_type = f"FY{fiscal_end % 100:02} Forecast_{start:%m%Y}"
    if not isinstance(forecast_type, str) or not forecast_type.strip():
        raise ValueError("Invalid GUI Forecast Type")
    forecast_type = forecast_type.strip()

    def key(value):
        return re.sub(
            r"[^a-z0-9]",
            "",
            str(value or "").casefold(),
        )

    names = {
        column.get("name")
        for column in sheet.iter("column")
        if key(column.get("caption")) == "forecaststart"
    }

    if len(names) != 1:
        raise ValueError(
            "Expected one Forecast Start dependency " "on the actuals tab."
        )

    for column in root.iter("column"):
        if column.get("name") in names and "param-domain-type" in column.attrib:
            column.set(
                "value",
                f"#{start:%Y-%m-%d}#",
            )

            calculation = column.find("calculation")

            if calculation is not None:
                calculation.set(
                    "formula",
                    f"#{start:%Y-%m-%d}#",
                )

    view = sheet.find(".//view")

    dependencies = [
        dependency
        for dependency in sheet.findall(".//datasource-dependencies")
        if any(
            column.get("caption") == mapping["actual_measure"]
            for column in dependency.findall("column")
        )
    ]

    if len(dependencies) != 1:
        raise ValueError("Expected one actuals datasource dependency.")

    dependency = dependencies[0]
    datasource = dependency.get("datasource")

    company_filters = [
        filter_element
        for filter_element in sheet.findall(".//filter")
        if filter_element.get("column", "").endswith("[none:Company:nk]")
    ]

    if not company_filters:
        ET.SubElement(
            dependency,
            "column-instance",
            {
                "column": "[Company]",
                "derivation": "None",
                "name": "[none:Company:nk]",
                "pivot": "key",
                "type": "nominal",
            },
        )

        new_filter = ET.Element(
            "filter",
            {
                "class": "categorical",
                "column": (f"[{datasource}].[none:Company:nk]"),
            },
        )

        allowed_before = {
            "datasources",
            "mapsources",
            "datasource-dependencies",
            "filter",
        }

        insertion = next(
            (
                index
                for index, child in enumerate(view)
                if child.tag not in allowed_before
            ),
            len(view),
        )

        view.insert(insertion, new_filter)
        company_filters = [new_filter]

    if len(company_filters) != 1:
        raise ValueError("Expected one Company filter.")

    for child in list(company_filters[0]):
        company_filters[0].remove(child)

    ET.SubElement(
        company_filters[0],
        "groupfilter",
        function="member",
        level="[none:Company:nk]",
        member=f'"{mapping["company"]}"',
    )

    months = []
    current = lower

    while current < actuals_end:
        months.append(current)

        current = date(
            current.year + (current.month == 12),
            current.month % 12 + 1,
            1,
        )

    date_filters = [
        filter_element
        for filter_element in sheet.findall(".//filter")
        if any(
            filter_element.get("column", "").endswith(f"[{level}:GL Posting Date:ok]")
            for level in ("tmn", "my")
        )
    ]

    for target in date_filters:
        level = (
            "[my:GL Posting Date:ok]"
            if target.get("column").endswith("[my:GL Posting Date:ok]")
            else "[tmn:GL Posting Date:ok]"
        )

        for child in list(target):
            target.remove(child)

        union = ET.SubElement(
            target,
            "groupfilter",
            function="union",
        )

        for month in months:
            member = (
                f"{month:%Y%m}"
                if level.startswith("[my:")
                else f"#{month:%Y-%m-%d} 00:00:00#"
            )

            ET.SubElement(
                union,
                "groupfilter",
                function="member",
                level=level,
                member=member,
            )

    # Explicit bounds also support HR, which has no date filter.
    actual_columns = [
        column
        for column in dependency.findall("column")
        if column.get("caption") == mapping["actual_measure"]
    ]

    if len(actual_columns) != 1:
        raise ValueError("Expected one Actual YTD measure.")

    actual_name = actual_columns[0].get("name")

    formula = (
        f"IF [GL Posting Date] >= #{lower:%Y-%m-%d}# "
        f"AND [GL Posting Date] < #{actuals_end:%Y-%m-%d}# "
        "THEN [Actual] ELSE 0 END"
    )

    for column in root.iter("column"):
        if column.get("name") == actual_name and column.find("calculation") is not None:
            column.find("calculation").set(
                "formula",
                formula,
            )

    for filter_element in root.iter("filter"):
        column = filter_element.get("column", "")

        if column.endswith("[none:Forecast Type:nk]"):
            for member in filter_element.findall('.//groupfilter[@function="member"]'):
                if re.fullmatch(
                    r"FY\d{2} Forecast_\d{6}",
                    member.get("member", "").strip('"'),
                    re.I,
                ):
                    member.set(
                        "member",
                        f'"{forecast_type}"',
                    )

        if column.endswith("[none:Fiscal Year:ok]"):
            for child in list(filter_element):
                filter_element.remove(child)

            ET.SubElement(
                filter_element,
                "groupfilter",
                function="member",
                level="[none:Fiscal Year:ok]",
                member=str(fiscal_end),
            )

    tree.write(
        destination,
        encoding="utf-8",
        xml_declaration=True,
    )

    return start, forecast_type


def _statement_values(path, closed_period, mapping=MAPPING):
    company = str(mapping['company']).upper()
    label = mapping['fs_line']
    if company == 'ALIV':
        from automations.actuals.adaptive_aliv_statement_reader import get_aliv_line
        return get_aliv_line(path, closed_period, label)
    if company == 'CBL':
        from automations.actuals.cbl_statement_periods import read_cbl_area
        from automations.actuals.actuals_report_layout import canonical
        data = read_cbl_area(path, 'OPEX', closed_period)
        def find(scope):
            matches=[v for k,v in data[scope].items() if canonical(k)==canonical(label)]
            if len(matches)!=1: raise ValueError(f'CBL {scope} statement line missing/ambiguous: {label}')
            return float(matches[0])
        return find('MONTH'),find('YTD')
    raise ValueError(f'Unsupported company: {company}')



def _save_complete_crosstab(audit_path, raw_file, source, opened_copy, tab, company, report_name, sequence):
    """Store EVERY Tableau column as captured, including numeric Actual YTD.

    Raw .txt is a temporary clipboard transport only; each capture is also
    preserved as an .xlsx file and in the combined audit workbook.
    """
    import csv
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import PatternFill, Font
    from openpyxl.utils import get_column_letter
    with Path(raw_file).open(encoding="utf-8-sig", newline="") as fh:
        table=list(csv.reader(fh,delimiter="\t"))
    if not table or len(table)<3:
        raise ValueError("Raw Tableau clipboard did not contain a complete table")
    audit_path=Path(audit_path)
    wb=load_workbook(audit_path) if audit_path.exists() else Workbook()
    if not audit_path.exists():
        wb.active.title="Capture Index"
        wb.active.append(["Sheet","Company","Source TWB","Worksheet","Rows","Raw XLSX"])
    title=f"{sequence:03d}_{company}_{Path(source).stem}"[:31]
    import re
    title=re.sub(r"[\\/*?:\[\]]","_",title)
    if title in wb.sheetnames: raise ValueError(f"Duplicate capture sheet: {title}")
    ws=wb.create_sheet(title)
    for row in table:
        values=[]
        for value in row:
            text=str(value).strip()
            if text and not text.endswith("%") and not text.startswith("="):
                try:
                    from decimal import Decimal
                    amount=Decimal(text.replace(",", "").replace("$", ""))
                    values.append(float(amount))
                    continue
                except Exception: pass
            values.append(text)
        ws.append(values)
    for row in ws.iter_rows(min_row=1,max_row=min(2,ws.max_row)):
        for cell in row:
            cell.fill=PatternFill("solid",fgColor="703070")
            cell.font=Font(bold=True,color="FFFFFF")
    ws.freeze_panes="E3"
    for i in range(1,ws.max_column+1):
        ws.column_dimensions[get_column_letter(i)].width=27 if i<=4 else 19
    ws.auto_filter.ref=f"A2:{get_column_letter(ws.max_column)}{ws.max_row}"
    raw_xlsx=Path(raw_file).with_suffix(".xlsx")
    single=Workbook(); only=single.active; only.title="Full Tableau Crosstab"
    for row in table:
        only.append([str(v).strip() for v in row])
    single.save(raw_xlsx); single.close()
    wb["Capture Index"].append([title,company,str(source),tab,len(table)-2,str(raw_xlsx)])
    wb.save(audit_path); wb.close()
    # Remove redundant TXT only after both Excel exports are successfully saved.
    Path(raw_file).unlink(missing_ok=True)
    return title

def _capture_mapping(
    *,
    report_path,
    aliv_fs_path,
    cbl_fs_path,
    output_paths,
    closed_period,
    mapping,
    report,
    raw,
    audit,
    sequence_offset=0,
    source=None,
    capture=None,
    opener=None,
    closer=None,
    wait_seconds=8,
    tableau_settings=None,
):
    from automations.actuals.tableau_standalone_forecast_service import (
        _append_crosstab_audit,
    )
    from automations.actuals.individual_opex_crosstab_parser import (
        extract_monthly_ytd,
    )

    if capture is None:
        from automations.actuals.tableau_desktop_export_service import (
            get_tableau_crosstab_rows,
        )

        capture = get_tableau_crosstab_rows

    if opener is None:
        from automations.forecast.tableau_service import (
            open_tableau_workbook,
        )

        opener = open_tableau_workbook

    if closer is None:
        from automations.actuals.tableau_desktop_close_service import (
            close_tableau_forecast_workbook,
        )

        closer = close_tableau_forecast_workbook

    route = _route_for(mapping)
    source = Path(source or route["individual_tableau_workbook"])
    tab = route["individual_tableau_worksheet"]

    # Consolidated comparisons are optional inputs, not a prerequisite for
    # capturing the individual Tableau workbook.
    consolidated_values = {}
    try:
        from automations.actuals.report_builder import META_SHEET

        company_sheet = report[mapping["output_company_sheet"]]
        # Final layout: YTD consolidated is column J; capture layout: column I.
        ytd_index = 9 if META_SHEET in report.sheetnames else 8
        for values in company_sheet.iter_rows(min_row=7, values_only=True):
            if str(values[0]).strip().casefold() == mapping["fs_line"].casefold():
                for basis, index in (("MONTH", 2), ("YTD", ytd_index)):
                    try:
                        consolidated_values[basis] = _number(values[index])
                    except (ValueError, IndexError, TypeError):
                        consolidated_values[basis] = None
                break
    except Exception as exc:
        print(f"[INDIVIDUAL OPEX] Consolidated comparison unavailable: {exc}", flush=True)

    try:
        company = mapping["company"].upper()

        if company == "ALIV":
            statement_path = aliv_fs_path
        elif company == "CBL":
            statement_path = cbl_fs_path
        else:
            raise ValueError(f"Unsupported company: {company}")

        if not statement_path:
            raise ValueError(
                f"Missing {company} financial statement path."
            )

        monthly_fs, ytd_fs = _statement_values(
            statement_path,
            closed_period,
            mapping,
        )
        fs_error = None
    except Exception as exc:
        monthly_fs = ytd_fs = None
        fs_error = str(exc)

    statements = {
        "MONTH": monthly_fs,
        "YTD": ytd_fs,
    }

    temporary = source.with_name(f"{source.stem}_Actuals_{uuid4().hex[:8]}.twb")

    opened = False
    actuals = {}
    audit_sheet = None
    start = None
    forecast_type = None
    capture_error = None
    close_warning = ""

    sequence = sequence_offset + 1

    try:
        # A missing statement comparison must not block a Tableau capture.
        # IMPORTANT:
        # YTD preparation retains all fiscal months.
        # The existing Forecast Start calculation remains unchanged.
        start, forecast_type = _prepare(
            source,
            temporary,
            closed_period,
            "YTD",
            mapping,
            tab,
            tableau_settings=tableau_settings,
        )

        print(
            f"[INDIVIDUAL OPEX] Prepared: {temporary}",
            flush=True,
        )
        print(
            f"[INDIVIDUAL OPEX] Worksheet: {tab}",
            flush=True,
        )
        print(
            f"[INDIVIDUAL OPEX] Reporting month: {closed_period:%B %Y}",
            flush=True,
        )

        # Snapshot existing raw exports so we can identify this capture.
        raw = Path(raw)
        raw.mkdir(parents=True, exist_ok=True)
        before = set(raw.glob("*.txt"))

        opener(tableau_workbook_path=str(temporary))
        opened = True

        if wait_seconds:
            time.sleep(wait_seconds)

        # ONE Tableau capture for MONTH and YTD.
        rows = capture(
            worksheet_name=tab,
            raw_output_folder=raw,
            sequence_number=sequence,
            tableau_workbook_path=str(temporary),
        )

        new_files = list(set(raw.glob("*.txt")) - before)

        if len(new_files) != 1:
            raise ValueError(
                "Expected exactly one new raw Tableau .txt export; "
                f"found {len(new_files)}. Check {raw}"
            )

        raw_file = new_files[0]

        print(
            f"[INDIVIDUAL OPEX] Raw crosstab: {raw_file}",
            flush=True,
        )

        # Extract both values from the same raw capture.
        actuals = extract_monthly_ytd(
            raw_file,
            mapping["company"],
            closed_period,
            mapping.get("fiscal_start_month", 7),
        )

        print(
            f"[INDIVIDUAL OPEX] {mapping['company']} "
            f"MONTH={actuals['MONTH']:,.0f} "
            f"YTD={actuals['YTD']:,.0f}",
            flush=True,
        )

        try:
            audit_sheet = _save_complete_crosstab(
                audit, raw_file, source, temporary, tab,
                mapping["company"], Path(report_path).name, sequence,
            )
        except Exception as audit_exc:
            # The parsed amounts are already in memory; preserve them even
            # if the optional audit workbook cannot be written.
            print(f"[INDIVIDUAL OPEX] Audit export warning: {audit_exc}", flush=True)

    except Exception as exc:
        capture_error = f"{type(exc).__name__}: {exc}"
        print(
            f"[INDIVIDUAL OPEX] ERROR: {capture_error}",
            flush=True,
        )
        import traceback

        traceback.print_exc()

    finally:
        if opened:
            try:
                closer(tableau_workbook_path=str(temporary))
            except Exception as exc:
                close_warning = f" Close warning: {exc}"

        # Do not leave modified temporary TWB files behind.
        # If closing fails, retain the file for investigation.
        if not close_warning:
            temporary.unlink(missing_ok=True)

    results = []

    for basis in ("MONTH", "YTD"):
        statement = statements[basis]
        consolidated = consolidated_values.get(basis)

        row = [
            basis,
            float(statement) if statement is not None else None,
            float(consolidated) if consolidated is not None else None,
            None,
            None,
            None,
            None,
            "ACTION REQUIRED",
            str(source),
            tab,
            audit_sheet,
            start.isoformat() if start is not None else None,
            forecast_type,
            "",
        ]

        if capture_error:
            row[13] = capture_error + close_warning
            results.append(row)
            continue

        # Capture is independent of reconciliation. Preserve Tableau actuals
        # even when a financial statement or consolidated comparison is missing.
        try:
            actual = actuals[basis]
            rounded = _rounded(actual / 1000)
            row[3] = float(actual)
            row[4] = rounded
        except Exception as exc:
            row[13] = f"Actual extraction error: {type(exc).__name__}: {exc}"
            results.append(row)
            continue

        try:
            if statement is None:
                raise ValueError(f"Missing financial statement {basis}.")

            if consolidated is None:
                raise ValueError(
                    f"Missing consolidated {basis} for {mapping['fs_line']}."
                )

            actual = actuals[basis]
            rounded = _rounded(actual / 1000)

            variance = rounded - _rounded(statement)
            consolidated_variance = rounded - _rounded(consolidated)

            row[3:7] = [
                float(actual),
                rounded,
                variance,
                consolidated_variance,
            ]

            row[7] = "PASS" if variance == 0 and consolidated_variance == 0 else "FAIL"

        except Exception as exc:
            row[13] = f"{type(exc).__name__}: {exc}"

        if close_warning:
            row[13] += close_warning

        results.append(row)

    return results


def run_individual_opex_pilot(
    *,
    report_path,
    aliv_fs_path,
    output_paths,
    closed_period,
    cbl_fs_path=None,
    source=None,
    capture=None,
    opener=None,
    closer=None,
    wait_seconds=8,
    tableau_settings=None,
):
    # Determine which company or companies the user selected.
    settings = tableau_settings or {}
    company_selection = str(
        settings.get("company", "Both")
    ).strip().upper()

    if company_selection == "BOTH":
        selected_companies = {"ALIV", "CBL"}
    elif company_selection in {"ALIV", "CBL"}:
        selected_companies = {company_selection}
    else:
        raise ValueError(
            f"Invalid company selection: {company_selection}"
        )

    report = load_workbook(report_path)
    raw = Path(output_paths["tableau_raw_copies"]) / f"Individual_Actuals_{uuid4().hex[:8]}"
    raw.mkdir(parents=True, exist_ok=True)
    audit = raw / "Individual_Actuals_Crosstabs.xlsx"
    groups = []
    from automations.actuals.actuals_report_layout import merge_individual_values
    from automations.actuals.report_builder import META_SHEET, update_individual_values

    # A finished report is rebuilt through the report builder; the capture
    # workbook is merged cell by cell.
    final_layout = META_SHEET in report.sheetnames

    def unavailable_pair(message):
        # The merger consumes fixed 14-field rows, one per basis.
        return [[basis, None, None, None, None, None, None,
                 "ACTION REQUIRED", None, None, None, None, None, message]
                for basis in ("MONTH", "YTD")]

    try:
        for mapping in INDIVIDUAL_OPEX_ACTUALS_MAPPINGS:
            if mapping["company"].upper() not in selected_companies:
                continue
            try:
                route = _route_for(mapping)
                if route.get("individual_actuals_enabled") is not True:
                    continue
                print(f"Individual actuals: {mapping['company']} / {mapping['fs_line']}"
                      f" | {route['individual_tableau_workbook']}", flush=True)
                rows = _capture_mapping(
                    report_path=report_path,
                    aliv_fs_path=aliv_fs_path,
                    cbl_fs_path=cbl_fs_path,
                    output_paths=output_paths,
                    closed_period=closed_period,
                    mapping=mapping,
                    report=report,
                    raw=raw,
                    audit=audit,
                    sequence_offset=len(groups),
                    source=None,
                    capture=capture,
                    opener=opener,
                    closer=closer,
                    wait_seconds=wait_seconds,
                    tableau_settings=tableau_settings,
                )
            except Exception as exc:
                import traceback
                print(f"[INDIVIDUAL OPEX] SKIPPED {mapping['company']} / "
                      f"{mapping['fs_line']}: {type(exc).__name__}: {exc}", flush=True)
                traceback.print_exc()
                rows = unavailable_pair(f"{type(exc).__name__}: {exc}")

            groups.append((mapping, rows))
            # Commit every mapping immediately. A later failure cannot erase
            # successful captures already written to the intermediate report.
            try:
                if final_layout:
                    month, ytd = rows
                    update_individual_values(
                        report_path,
                        mapping["company"].upper(),
                        {mapping["fs_line"]: (month[4], ytd[4], month[13] or ytd[13])},
                    )
                else:
                    merge_individual_values(report, [(mapping, rows)])
                    report.save(report_path)
                print(f"[INDIVIDUAL OPEX] Saved {mapping['company']} / "
                      f"{mapping['fs_line']} to validation workbook", flush=True)
            except Exception as exc:
                print(f"[INDIVIDUAL OPEX] MERGE ERROR {mapping['company']} / "
                      f"{mapping['fs_line']}: {type(exc).__name__}: {exc}", flush=True)
                # Try to preserve previously completed rows and continue.
                try:
                    if not final_layout:
                        report.save(report_path)
                except Exception as save_exc:
                    print(f"[INDIVIDUAL OPEX] SAVE ERROR: {save_exc}", flush=True)
    finally:
        report.close()

    statuses = [row[7] for _, rows in groups for row in rows]

    return {
        "status": (
            "PASS"
            if statuses and all(status == "PASS" for status in statuses)
            else "REVIEW REQUIRED"
        ),
        "audit_workbook": str(audit),
        "statuses": statuses,
    }
