"""Final actuals validation report (OPEX, Direct Cost and Revenue).

Every area produces the same side-by-side layout per company sheet:

    A      Financial Statement Line
    B:H    Monthly   FS | Consolidated | Individual | V1 | V2 | Status | Notes
    I:O    YTD       FS | Consolidated | Individual | V1 | V2 | Status | Notes
    P      (hidden)  source note carried over from the capture step

V1 = Consolidated Tableau - Financial Statement
V2 = Individual Tableau - Consolidated Tableau

The workbook also keeps a hidden "_meta" sheet so the report can be rebuilt
after later steps (for example the Direct Cost individual merge) change values.
"""

from __future__ import annotations

import getpass
import os
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

TOLERANCE_000 = 1.0  # $1,000 tolerance; all report values are in $000

STATUS_PASS = "PASS"
STATUS_REVIEW = "REVIEW REQUIRED"
STATUS_PENDING = "PENDING"

COMPANIES = ("ALIV", "CBL")
SUMMARY_SHEET = "Executive Summary"
META_SHEET = "_meta"
OVERALL_PREFIX = "OVERALL VALIDATION STATUS:"

# Lines rendered as totals (bold with a top rule).
TOTAL_LINES = {
    "OPERATING EXPENSES",
    "DIRECT COST",
    "COST OF GOODS SOLD",
    "GROSS REVENUE",
    "SERVICE REVENUES",
}

# --- Styling --------------------------------------------------------------
FONT_NAME = "Aptos"
NAVY = "243F5C"
MONTHLY_COLOR = "5B2C6F"
YTD_COLOR = "145A70"
MONTHLY_TINT = "F3EDF6"
YTD_TINT = "E8F1F4"
BAND_FILL = "F7F9FB"
BORDER_COLOR = "D0D7DE"
VARIANCE_FLAG_FILL = "FBE3E3"

STATUS_STYLES = {
    STATUS_PASS: ("E2F0D9", "1E6B1E"),
    STATUS_REVIEW: ("F4CCCC", "9C0006"),
    STATUS_PENDING: ("FFF2CC", "8A5A00"),
}

AMOUNT_FORMAT = "#,##0;[Red](#,##0);–"

THIN = Side(style="thin", color=BORDER_COLOR)
MEDIUM = Side(style="medium", color="7F8C8D")

BLOCK_HEADERS = (
    "Financial Statement",
    "Consolidated Tableau",
    "Individual Tableau",
    "V1\nConsolidated − FS",
    "V2\nIndividual − Consolidated",
    "Status",
    "Notes",
)
MONTHLY_COL = 2  # B
YTD_COL = 9  # I
SOURCE_NOTE_COL = 16  # P (hidden)
LAST_COL = 15  # O
HEADER_ROW = 6
FIRST_DATA_ROW = 7
TOLERANCE_CELL = "$M$4"


def canonical(value) -> str:
    return " ".join(str(value or "").upper().replace("&", " AND ").split())


def is_number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def variance(base, compare):
    """compare - base, or None when either side is missing."""
    if not is_number(base) or not is_number(compare):
        return None
    return round(float(compare) - float(base), 6)


# --- Data model -----------------------------------------------------------


@dataclass
class PeriodValues:
    fs: float | None = None
    consolidated: float | None = None
    individual: float | None = None

    @property
    def v1(self):
        return variance(self.fs, self.consolidated)

    @property
    def v2(self):
        return variance(self.consolidated, self.individual)

    @property
    def status(self) -> str:
        v1, v2 = self.v1, self.v2
        if v1 is None or v2 is None:
            return STATUS_PENDING
        if abs(v1) <= TOLERANCE_000 and abs(v2) <= TOLERANCE_000:
            return STATUS_PASS
        return STATUS_REVIEW

    def note(self) -> str:
        parts = []
        v1, v2 = self.v1, self.v2
        if self.fs is None:
            parts.append("Financial statement value missing")
        if self.consolidated is None:
            parts.append("Consolidated Tableau value missing")
        elif v1 is not None and abs(v1) > TOLERANCE_000:
            parts.append(f"Consolidated differs from FS by {v1:,.1f}")
        if self.individual is None:
            parts.append("Individual Tableau not captured")
        elif v2 is not None and abs(v2) > TOLERANCE_000:
            parts.append(f"Individual differs from consolidated by {v2:,.1f}")
        return "; ".join(parts)


@dataclass
class ReportLine:
    name: str
    monthly: PeriodValues = field(default_factory=PeriodValues)
    ytd: PeriodValues = field(default_factory=PeriodValues)
    source_note: str = ""

    @property
    def is_total(self) -> bool:
        key = canonical(self.name)
        return key.startswith("TOTAL") or key in {canonical(x) for x in TOTAL_LINES}


@dataclass
class CompanySheet:
    company: str
    area: str
    lines: list[ReportLine]

    @property
    def title(self) -> str:
        return f"{self.company} {self.area}"

    @property
    def statuses(self) -> list[str]:
        return [s for line in self.lines for s in (line.monthly.status, line.ytd.status)]

    @property
    def status(self) -> str:
        return overall_status(self.statuses)


@dataclass
class ReportMeta:
    area: str
    closed_period: str = ""
    forecast_start: str = ""
    forecast_type: str = ""
    run_at: str = ""
    prepared_by: str = ""


def overall_status(statuses: Iterable[str]) -> str:
    statuses = list(statuses)
    if any(s == STATUS_REVIEW for s in statuses):
        return STATUS_REVIEW
    if not statuses or any(s != STATUS_PASS for s in statuses):
        return STATUS_PENDING
    return STATUS_PASS


def total_components(lines: list[ReportLine]) -> dict[int, tuple[int, int]]:
    """For each total line, the (first, last) list indexes of the lines it sums.

    A total sums the contiguous non-total lines directly above it.
    """
    result = {}
    start = 0
    for index, line in enumerate(lines):
        if line.is_total:
            if index > start:
                result[index] = (start, index - 1)
            start = index + 1
    return result


def _note_formula(status_ref: str, text: str):
    """Notes stay empty while the line's Status is PASS (live, like the Status)."""
    if not text:
        return None
    safe = text.replace('"', '""')
    # Excel limits a quoted text constant inside a formula to 255 characters.
    chunks = [safe[i : i + 240] for i in range(0, len(safe), 240)]
    joined = "&".join(f'"{c}"' for c in chunks)
    return f'=IF({status_ref}="{STATUS_PASS}","",{joined})'


def apply_zero_rule(sheet: CompanySheet) -> None:
    """A line the statement reports as 0 shows 0 (not a blank) in the Tableau
    columns too: Tableau lists no rows for a category with no postings."""
    for line in sheet.lines:
        if line.is_total:
            continue
        for period in (line.monthly, line.ytd):
            if period.fs == 0:
                if period.consolidated is None:
                    period.consolidated = 0.0
                if period.individual is None:
                    period.individual = 0.0


def apply_total_rules(sheet: CompanySheet) -> None:
    """Total lines show the sum of the lines above in the Consolidated and
    Individual columns (the Financial Statement column is never changed).
    The workbook carries the same rule as an Excel SUM formula."""
    for index, (first, last) in total_components(sheet.lines).items():
        total = sheet.lines[index]
        for period in ("monthly", "ytd"):
            for attr in ("consolidated", "individual"):
                parts = [
                    getattr(getattr(sheet.lines[i], period), attr)
                    for i in range(first, last + 1)
                ]
                numbers = [p for p in parts if is_number(p)]
                setattr(
                    getattr(total, period),
                    attr,
                    float(sum(numbers)) if numbers else None,
                )


# --- Reading existing workbooks -------------------------------------------


def _company_sheets(workbook) -> dict[str, Worksheet]:
    found = {}
    for sheet in workbook.worksheets:
        for company in COMPANIES:
            if sheet.title.startswith(company + " "):
                if company in found:
                    raise ValueError(
                        f"Expected one {company} sheet; found {found[company].title!r} "
                        f"and {sheet.title!r}"
                    )
                found[company] = sheet
    missing = [c for c in COMPANIES if c not in found]
    if missing:
        raise ValueError(f"Report has no sheet for: {', '.join(missing)}")
    return found


def _num(value):
    return float(value) if is_number(value) else None


def _text(value) -> str:
    return "" if value is None else str(value).strip()


def detect_layout(workbook) -> str:
    """Return 'final', 'compact_v1' or 'capture' (the raw A:S capture layout)."""
    sheet = next(iter(_company_sheets(workbook).values()))
    if META_SHEET in workbook.sheetnames and _text(sheet.cell(HEADER_ROW, 1).value) == (
        "Financial Statement Line"
    ):
        return "final"
    if _text(sheet["A7"].value) == "Financial Statement Line" and _text(
        sheet["M7"].value
    ) == "YTD Status":
        return "compact_v1"
    if _text(sheet["A6"].value) == "Financial Statement Line":
        return "capture"
    raise ValueError(f"{sheet.title}: unrecognised report layout")


def _read_rows(sheet, first_row, columns, note_col=None) -> list[ReportLine]:
    name_c, mfs, mcons, mind, yfs, ycons, yind = columns
    lines = []
    for row in range(first_row, sheet.max_row + 1):
        name = _text(sheet.cell(row, name_c).value)
        if not name:
            continue
        cell = lambda c: _num(sheet.cell(row, c).value)  # noqa: E731
        note = _text(sheet.cell(row, note_col).value) if note_col else ""
        if note.lower().rstrip(".") == "reconciled":
            note = ""
        lines.append(
            ReportLine(
                name=name,
                monthly=PeriodValues(cell(mfs), cell(mcons), cell(mind)),
                ytd=PeriodValues(cell(yfs), cell(ycons), cell(yind)),
                source_note=note,
            )
        )
    return lines


def read_report(path) -> tuple[ReportMeta, list[CompanySheet]]:
    """Read any known report layout into the data model."""
    workbook = load_workbook(path, data_only=True)
    try:
        layout = detect_layout(workbook)
        sheets = _company_sheets(workbook)
        area = sheets["ALIV"].title.split(" ", 1)[1]
        meta = ReportMeta(area=area)

        if layout == "final":
            meta_ws = workbook[META_SHEET]
            values = {
                _text(r[0]): _text(r[1])
                for r in meta_ws.iter_rows(values_only=True)
                if r and r[0]
            }
            for key in ("closed_period", "forecast_start", "forecast_type", "run_at", "prepared_by"):
                setattr(meta, key, values.get(key, ""))
            columns, first, note_col = (1, 2, 3, 4, 9, 10, 11), FIRST_DATA_ROW, SOURCE_NOTE_COL
        elif layout == "compact_v1":
            meta.closed_period = _text(sheets["ALIV"]["C2"].value)
            columns, first, note_col = (1, 2, 3, 4, 8, 9, 10), 8, None
        else:
            meta.closed_period = _text(sheets["ALIV"]["C2"].value)
            meta.forecast_start = _text(sheets["ALIV"]["C3"].value)
            columns, first, note_col = (1, 2, 3, 12, 8, 9, 16), 7, 6

        companies = [
            CompanySheet(company, area, _read_rows(sheets[company], first, columns, note_col))
            for company in COMPANIES
        ]
        return meta, companies
    finally:
        workbook.close()


# --- Writing --------------------------------------------------------------


def _fill(color):
    return PatternFill(fill_type="solid", fgColor=color)


def _font(**kwargs):
    return Font(name=FONT_NAME, **kwargs)


def _style_status(cell, status):
    fill, color = STATUS_STYLES.get(status, STATUS_STYLES[STATUS_PENDING])
    cell.value = status
    cell.fill = _fill(fill)
    cell.font = _font(bold=True, color=color, size=10)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _subtitle(meta: ReportMeta) -> str:
    forecast = " ".join(x for x in (meta.forecast_type, f"from {meta.forecast_start}" if meta.forecast_start else "") if x)
    parts = [
        f"Closed period: {meta.closed_period or 'not supplied'}",
        f"Forecast: {forecast or 'not supplied'}",
        f"Run: {meta.run_at}",
    ]
    if meta.prepared_by:
        parts.append(f"By: {meta.prepared_by}")
    return "   |   ".join(parts)


def _setup_print(ws, last_col, repeat_rows):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_area = f"A1:{get_column_letter(last_col)}{max(ws.max_row, 1)}"
    if repeat_rows:
        ws.print_title_rows = repeat_rows
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Page &P of &N"


# Column widths on company sheets: line name, status and notes columns are wider.
COLUMN_WIDTHS = {1: 32, 7: 18, 8: 46, 14: 18, 15: 46}


def _row_height(ws, row) -> float:
    """Excel does not resize rows for wrapped text written by openpyxl, so
    estimate the lines each wrapped cell needs and size the row to fit."""
    lines = 1
    for col in (1, MONTHLY_COL + 5, MONTHLY_COL + 6, YTD_COL + 5, YTD_COL + 6):
        text = ws.cell(row, col).value
        if not text:
            continue
        width = ws.column_dimensions[get_column_letter(col)].width or 10
        chars_per_line = max(int(width * 1.15), 1)  # notes use a 9pt font
        needed = 0
        for part in str(text).split("\n"):
            needed += 1
            length = 0
            for word in part.split():
                if length and length + 1 + len(word) > chars_per_line:
                    needed += 1
                    length = len(word)
                else:
                    length += (1 if length else 0) + len(word)
        lines = max(lines, needed)
    return max(20, 13 * lines + 7)


def _add_conditional_formats(ws, line_count: int) -> None:
    """Colour variances and statuses from their live values, so the colours
    follow the formulas when someone edits a number."""
    if not line_count:
        return
    last_row = FIRST_DATA_ROW + line_count - 1
    flag_fill = _fill(VARIANCE_FLAG_FILL)
    flag_font = Font(name=FONT_NAME, bold=True, color="9C0006")
    for start in (MONTHLY_COL, YTD_COL):
        for offset in (3, 4):
            letter = get_column_letter(start + offset)
            rng = f"{letter}{FIRST_DATA_ROW}:{letter}{last_row}"
            ws.conditional_formatting.add(
                rng,
                FormulaRule(
                    formula=[
                        f"AND(ISNUMBER({letter}{FIRST_DATA_ROW}),"
                        f"ABS({letter}{FIRST_DATA_ROW})>{TOLERANCE_CELL})"
                    ],
                    fill=flag_fill,
                    font=flag_font,
                ),
            )
        letter = get_column_letter(start + 5)
        targets = [f"{letter}{FIRST_DATA_ROW}:{letter}{last_row}"]
        if start == MONTHLY_COL:
            targets.append("M2")
        for rng in targets:
            for label, (fill, color) in STATUS_STYLES.items():
                ws.conditional_formatting.add(
                    rng,
                    CellIsRule(
                        operator="equal",
                        formula=[f'"{label}"'],
                        fill=_fill(fill),
                        font=Font(name=FONT_NAME, bold=True, color=color),
                    ),
                )


def _write_company_sheet(ws, sheet: CompanySheet, meta: ReportMeta):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90

    # Title, subtitle and sheet status
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=LAST_COL)
    title = ws.cell(1, 1, f"{sheet.title.upper()} — ACTUALS VALIDATION")
    title.font = _font(size=16, bold=True, color="FFFFFF")
    title.fill = _fill(NAVY)
    title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 30

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=11)
    sub = ws.cell(2, 1, _subtitle(meta))
    sub.font = _font(size=10, color="44546A")
    sub.alignment = Alignment(vertical="center", indent=1)
    ws.cell(2, 12, "Sheet status").font = _font(size=10, bold=True, color="44546A")
    ws.cell(2, 12).alignment = Alignment(horizontal="right", vertical="center")
    ws.merge_cells(start_row=2, start_column=13, end_row=2, end_column=LAST_COL)
    _style_status(ws.cell(2, 13), sheet.status)
    last_row = FIRST_DATA_ROW + len(sheet.lines) - 1
    status_ranges = [
        f"{get_column_letter(c + 5)}{FIRST_DATA_ROW}:{get_column_letter(c + 5)}{last_row}"
        for c in (MONTHLY_COL, YTD_COL)
    ]

    def _count(label):
        return "+".join(f'COUNTIF({r},"{label}")' for r in status_ranges)

    ws.cell(2, 13).value = (
        f'=IF({_count(STATUS_REVIEW)}>0,"{STATUS_REVIEW}",'
        f'IF({_count(STATUS_PASS)}={len(sheet.lines) * 2},"{STATUS_PASS}","{STATUS_PENDING}"))'
    )
    ws.row_dimensions[2].height = 22

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=LAST_COL)
    legend = ws.cell(
        3,
        1,
        "All amounts in $000.  V1 = Consolidated Tableau − Financial Statement.  "
        "V2 = Individual Tableau − Consolidated Tableau.  "
        "V1, V2, Status and total lines are live Excel formulas.  "
        "PASS when both variances are within the tolerance in M4.  "
        "Shaded variances are outside tolerance.",
    )
    legend.font = _font(size=9, italic=True, color="6B7785")
    legend.alignment = Alignment(vertical="center", indent=1)

    ws.cell(4, 12, "Tolerance (± $000)").font = _font(size=10, bold=True, color="44546A")
    ws.cell(4, 12).alignment = Alignment(horizontal="right", vertical="center")
    tol = ws.cell(4, 13, TOLERANCE_000)
    tol.font = _font(size=10, bold=True, color="1F3FBF")
    tol.fill = _fill("FFF9DB")
    tol.number_format = "0.0"
    tol.alignment = Alignment(horizontal="center", vertical="center")

    # Block banners
    period = meta.closed_period or ""
    for start, label, color in (
        (MONTHLY_COL, f"MONTHLY  —  {period}".strip(" —"), MONTHLY_COLOR),
        (YTD_COL, f"YEAR TO DATE  —  {period}".strip(" —"), YTD_COLOR),
    ):
        ws.merge_cells(start_row=5, start_column=start, end_row=5, end_column=start + 6)
        banner = ws.cell(5, start, label)
        banner.font = _font(size=11, bold=True, color="FFFFFF")
        banner.fill = _fill(color)
        banner.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 22

    # Column headers
    header = ws.cell(HEADER_ROW, 1, "Financial Statement Line")
    header.font = _font(bold=True, color="FFFFFF", size=10)
    header.fill = _fill(NAVY)
    header.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    for start, color in ((MONTHLY_COL, MONTHLY_COLOR), (YTD_COL, YTD_COLOR)):
        for offset, label in enumerate(BLOCK_HEADERS):
            cell = ws.cell(HEADER_ROW, start + offset, label)
            cell.font = _font(bold=True, color="FFFFFF", size=10)
            cell.fill = _fill(color)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.cell(HEADER_ROW, SOURCE_NOTE_COL, "Source note")
    ws.row_dimensions[HEADER_ROW].height = 58
    for col in range(1, LAST_COL + 1):
        ws.column_dimensions[get_column_letter(col)].width = COLUMN_WIDTHS.get(col, 15)

    # Data rows
    row = FIRST_DATA_ROW
    for index, line in enumerate(sheet.lines):
        band = _fill(BAND_FILL) if index % 2 else None
        name = ws.cell(row, 1, line.name)
        name.font = _font(size=10, bold=line.is_total)
        name.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
        if band:
            name.fill = band

        components = total_components(sheet.lines).get(index)
        for start, values in ((MONTHLY_COL, line.monthly), (YTD_COL, line.ytd)):
            fs_c, cons_c, ind_c, v1_c, v2_c = (get_column_letter(start + i) for i in range(5))
            amounts = (values.fs, values.consolidated, values.individual)
            for offset, amount in enumerate(amounts):
                if offset > 0 and components:
                    # Total line: sum of the lines above in this column.
                    col_letter = get_column_letter(start + offset)
                    rng = (
                        f"{col_letter}{FIRST_DATA_ROW + components[0]}:"
                        f"{col_letter}{FIRST_DATA_ROW + components[1]}"
                    )
                    amount = f'=IF(COUNT({rng})=0,"",SUM({rng}))'
                cell = ws.cell(row, start + offset, amount)
                cell.number_format = AMOUNT_FORMAT
                cell.font = _font(size=10, bold=line.is_total)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if band:
                    cell.fill = band
            # V1, V2 and Status are formulas so they show in the formula bar.
            ws.cell(row, start + 3).value = (
                f'=IF(OR({fs_c}{row}="",{cons_c}{row}=""),"",ROUND({cons_c}{row}-{fs_c}{row},6))'
            )
            ws.cell(row, start + 4).value = (
                f'=IF(OR({cons_c}{row}="",{ind_c}{row}=""),"",ROUND({ind_c}{row}-{cons_c}{row},6))'
            )
            for offset in (3, 4):
                cell = ws.cell(row, start + offset)
                cell.number_format = AMOUNT_FORMAT
                cell.font = _font(size=10, bold=line.is_total)
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if band:
                    cell.fill = band
            status = ws.cell(row, start + 5)
            status.value = (
                f'=IF(OR({v1_c}{row}="",{v2_c}{row}=""),"{STATUS_PENDING}",'
                f'IF(AND(ABS({v1_c}{row})<={TOLERANCE_CELL},ABS({v2_c}{row})<={TOLERANCE_CELL}),'
                f'"{STATUS_PASS}","{STATUS_REVIEW}"))'
            )
            status.font = _font(bold=True, size=10)
            status.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            note_parts = [values.note()]
            if start == MONTHLY_COL and line.source_note:
                note_parts.insert(0, line.source_note)
            note_text = "; ".join(p.strip().rstrip(".") for p in note_parts if p and p.strip())
            status_ref = f"{get_column_letter(start + 5)}{row}"
            note = ws.cell(row, start + 6, _note_formula(status_ref, note_text))
            note.font = _font(size=9, color="44546A")
            note.alignment = Alignment(vertical="center", wrap_text=True)
            if band:
                note.fill = band

        ws.cell(row, SOURCE_NOTE_COL, line.source_note or None)
        ws.row_dimensions[row].height = _row_height(ws, row)

        for col in range(1, LAST_COL + 1):
            cell = ws.cell(row, col)
            cell.border = Border(
                bottom=THIN,
                top=MEDIUM if line.is_total else None,
                left=MEDIUM if col == YTD_COL else None,
            )
        row += 1

    _add_conditional_formats(ws, len(sheet.lines))

    for col in range(1, LAST_COL + 1):
        ws.cell(HEADER_ROW, col).border = Border(left=MEDIUM if col == YTD_COL else None)

    ws.column_dimensions[get_column_letter(SOURCE_NOTE_COL)].hidden = True

    # No filter buttons: on a sheet this size they only cover the headers.
    ws.freeze_panes = ws.cell(FIRST_DATA_ROW, 2)
    _setup_print(ws, LAST_COL, f"{5}:{HEADER_ROW}")


def _write_summary(ws, meta: ReportMeta, sheets: list[CompanySheet], area_sheet_titles):
    ws.sheet_view.showGridLines = False
    overall = overall_status(s.status for s in sheets)

    ws.merge_cells("A1:I1")
    title = ws["A1"]
    title.value = f"{meta.area.upper()} ACTUALS VALIDATION — EXECUTIVE SUMMARY"
    title.font = _font(size=16, bold=True, color="FFFFFF")
    title.fill = _fill(NAVY)
    title.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 32

    info = (
        ("Closed period", meta.closed_period or "Not supplied"),
        ("Forecast start", meta.forecast_start or "Not supplied"),
        ("Forecast type", meta.forecast_type or "Not supplied"),
        ("Run date", meta.run_at),
        ("Prepared by", meta.prepared_by or ""),
    )
    for offset, (label, value) in enumerate(info, 3):
        ws.cell(offset, 1, label).font = _font(bold=True, color="44546A")
        ws.cell(offset, 2, value).font = _font()

    ws.merge_cells("E3:I4")
    badge = ws["E3"]
    _style_status(badge, overall)
    badge.value = f"{OVERALL_PREFIX} {overall}"
    badge.font = _font(size=14, bold=True, color=STATUS_STYLES[overall][1])

    # Company table
    header_row = 9
    headers = (
        "Company",
        "Lines",
        "Monthly PASS",
        "Monthly REVIEW",
        "Monthly PENDING",
        "YTD PASS",
        "YTD REVIEW",
        "YTD PENDING",
        "Status",
    )
    for col, label in enumerate(headers, 1):
        cell = ws.cell(header_row, col, label)
        cell.font = _font(bold=True, color="FFFFFF", size=10)
        cell.fill = _fill(YTD_COLOR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row, sheet in enumerate(sheets, header_row + 1):
        monthly = [line.monthly.status for line in sheet.lines]
        ytd = [line.ytd.status for line in sheet.lines]
        values = (
            sheet.company,
            len(sheet.lines),
            monthly.count(STATUS_PASS),
            monthly.count(STATUS_REVIEW),
            monthly.count(STATUS_PENDING),
            ytd.count(STATUS_PASS),
            ytd.count(STATUS_REVIEW),
            ytd.count(STATUS_PENDING),
        )
        for col, value in enumerate(values, 1):
            cell = ws.cell(row, col, value)
            cell.font = _font(bold=col == 1)
            cell.alignment = Alignment(horizontal="left" if col == 1 else "center")
            cell.border = Border(bottom=THIN)
        link = ws.cell(row, 1)
        link.hyperlink = f"#'{area_sheet_titles[sheet.company]}'!A1"
        link.font = _font(bold=True, color="1F5FA8", underline="single")
        _style_status(ws.cell(row, 9), sheet.status)
        ws.cell(row, 9).border = Border(bottom=THIN)

    # Exceptions
    row = header_row + len(sheets) + 3
    ws.cell(row, 1, "ITEMS REQUIRING ATTENTION").font = _font(size=12, bold=True, color=NAVY)
    row += 1
    exc_headers = ("Company", "Statement line", "Period", "Status", "V1 ($000)", "V2 ($000)", "Notes")
    for col, label in enumerate(exc_headers, 1):
        cell = ws.cell(row, col, label)
        cell.font = _font(bold=True, color="FFFFFF", size=10)
        cell.fill = _fill(NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=9)
    row += 1
    first_exception_row = row

    for sheet in sheets:
        for index, line in enumerate(sheet.lines):
            for period, values, start in (("Monthly", line.monthly, MONTHLY_COL), ("YTD", line.ytd, YTD_COL)):
                if values.status == STATUS_PASS:
                    continue
                cells = (sheet.company, line.name, period, None, values.v1, values.v2, values.note())
                for col, value in enumerate(cells, 1):
                    cell = ws.cell(row, col, value)
                    cell.font = _font(size=10)
                    cell.border = Border(bottom=THIN)
                    if col in (5, 6):
                        cell.number_format = AMOUNT_FORMAT
                    cell.alignment = Alignment(vertical="center", wrap_text=col in (2, 7))
                _style_status(ws.cell(row, 4), values.status)
                ws.cell(row, 2).hyperlink = (
                    f"#'{area_sheet_titles[sheet.company]}'!"
                    f"{get_column_letter(start)}{FIRST_DATA_ROW + index}"
                )
                ws.cell(row, 2).font = _font(size=10, color="1F5FA8", underline="single")
                ws.merge_cells(start_row=row, start_column=7, end_row=row, end_column=9)
                row += 1
    if row == first_exception_row:
        ws.cell(row, 1, "No exceptions — every line passed for both periods.").font = _font(
            italic=True, color="1E6B1E"
        )
        row += 1

    # Sign-off
    row += 2
    ws.cell(row, 1, "REVIEW SIGN-OFF").font = _font(size=12, bold=True, color=NAVY)
    for offset, label in enumerate(("Prepared by", "Reviewed by", "Date reviewed"), 1):
        ws.cell(row + offset, 1, label).font = _font(bold=True, color="44546A")
        for col in (2, 3, 4):
            ws.cell(row + offset, col).border = Border(bottom=THIN)
    ws.cell(row + 1, 2, meta.prepared_by or None)

    widths = {1: 18, 2: 34, 3: 12, 4: 18, 5: 13, 6: 13, 7: 16, 8: 16, 9: 18}
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A2"
    _setup_print(ws, 9, None)
    ws.page_setup.orientation = "portrait"


def build_report_workbook(meta: ReportMeta, sheets: list[CompanySheet]) -> Workbook:
    workbook = Workbook()
    summary = workbook.active
    summary.title = SUMMARY_SHEET
    summary.sheet_properties.tabColor = NAVY

    titles = {}
    for sheet in sheets:
        apply_total_rules(sheet)
    for sheet in sheets:
        ws = workbook.create_sheet(sheet.title[:31])
        ws.sheet_properties.tabColor = MONTHLY_COLOR if sheet.company == "ALIV" else YTD_COLOR
        _write_company_sheet(ws, sheet, meta)
        titles[sheet.company] = ws.title

    _write_summary(summary, meta, sheets, titles)

    meta_ws = workbook.create_sheet(META_SHEET)
    for key, value in vars(meta).items():
        meta_ws.append([key, value])
    meta_ws.sheet_state = "hidden"

    workbook.active = 0
    return workbook


def save_report(meta: ReportMeta, sheets: list[CompanySheet], path) -> Path:
    """Write atomically so a crash never leaves a half-written report."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    staged = path.with_name(path.stem + ".__tmp__.xlsx")
    for sheet in sheets:
        apply_zero_rule(sheet)
    workbook = build_report_workbook(meta, sheets)
    try:
        workbook.save(staged)
        try:
            os.replace(staged, path)
        except PermissionError as exc:
            raise PermissionError(
                f"Could not save the report because it is open in another program "
                f"(usually Excel). Close it and run again.\n\n{path}"
            ) from exc
    finally:
        workbook.close()
        staged.unlink(missing_ok=True)
    return path


def _current_user() -> str:
    try:
        return getpass.getuser()
    except Exception:
        return ""


def _format_period(value) -> str:
    if isinstance(value, (date, datetime)):
        return value.strftime("%B %Y")
    return _text(value)


def finalize_report(path, *, forecast_type=None, forecast_start=None, run_at=None) -> str:
    """Rebuild any report layout as the final layout, in place. Safe to repeat."""
    meta, sheets = read_report(path)
    if forecast_type:
        meta.forecast_type = _text(forecast_type)
    if forecast_start:
        meta.forecast_start = (
            forecast_start.strftime("%m/%d/%Y")
            if isinstance(forecast_start, (date, datetime))
            else _text(forecast_start)
        )
    meta.closed_period = _format_period(meta.closed_period)
    meta.run_at = (run_at or datetime.now()).strftime("%Y-%m-%d %H:%M")
    meta.prepared_by = meta.prepared_by or _current_user()
    save_report(meta, sheets, path)
    return str(path)


def update_individual_values(path, company: str, updates: dict[str, tuple]) -> dict:
    """Set individual Tableau values on a final report and rebuild it.

    ``updates`` maps statement line -> (monthly, ytd, note). ``None`` amounts
    leave the cell empty, so that period stays PENDING.
    Returns {line: "updated" | "missing"}.
    """
    meta, sheets = read_report(path)
    sheet = next(s for s in sheets if s.company == company)
    by_key = {}
    for line in sheet.lines:
        by_key.setdefault(canonical(line.name), []).append(line)

    outcome = {}
    for name, (monthly, ytd, note) in updates.items():
        matches = by_key.get(canonical(name), [])
        if len(matches) != 1:
            outcome[name] = "missing"
            continue
        line = matches[0]
        line.monthly.individual = monthly
        line.ytd.individual = ytd
        if note:
            line.source_note = note
        outcome[name] = "updated"

    save_report(meta, sheets, path)
    return outcome


def report_status(path) -> str:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        for row in workbook[SUMMARY_SHEET].iter_rows(values_only=True):
            for value in row:
                if isinstance(value, str) and value.startswith(OVERALL_PREFIX):
                    return value.split(":", 1)[1].strip()
        return STATUS_PENDING
    finally:
        workbook.close()


def add_run_notes_sheet(path, *, filters=(), notes=()) -> None:
    """Add a 'Run Notes' tab: filters set in each Tableau workbook and capture problems.

    Formulas in the report are left untouched (the file is opened normally,
    not value-only).
    """
    from openpyxl.styles import Alignment, Font, PatternFill

    workbook = load_workbook(path)
    try:
        if "Run Notes" in workbook.sheetnames:
            del workbook["Run Notes"]
        ws = workbook.create_sheet("Run Notes", 1)
        ws.sheet_properties.tabColor = "243F5C"

        def block(row, title, rows):
            ws.cell(row, 1, title).font = Font(bold=True, size=12, color="243F5C")
            row += 1
            if not rows:
                ws.cell(row, 1, "None")
                return row + 2
            headers = list(dict.fromkeys(k for r in rows for k in r))
            for c, h in enumerate(headers, 1):
                cell = ws.cell(row, c, h)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="703070")
                cell.alignment = Alignment(horizontal="center")
            for r in rows:
                row += 1
                for c, h in enumerate(headers, 1):
                    ws.cell(row, c, str(r.get(h, "")))
            return row + 2

        nxt = block(1, "Filters set in the source Tableau workbooks (saved before each capture)", list(filters))
        block(nxt, "Capture problems (a missing value means its capture failed; see the note)", list(notes))
        for col, width in zip("ABCDEFGH", (14, 46, 18, 22, 22, 30, 40, 40)):
            ws.column_dimensions[col].width = width
        workbook.save(path)
    finally:
        workbook.close()
