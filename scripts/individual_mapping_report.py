"""Build the Step 3 (Validate) individual Tableau mapping review PDF.

    python scripts/individual_mapping_report.py ["Step 3 Individual Mapping Review.pdf"]

One table row per individual Tableau tab, with every forecast sheet it checks.
Needs reportlab (requirements-dev.txt). Run again after changing the mapping.
"""
import re
import sys
from collections import defaultdict
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from automations.forecast.mapping import FORECAST_CATEGORIES, FORECAST_ROUTES  # noqa: E402

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUT = sys.argv[1] if len(sys.argv) > 1 else "Step 3 Individual Mapping Review.pdf"
INK, MUTED, LINE, HEAD, WARN = (colors.HexColor(c) for c in ("#1f2933", "#616e7c", "#cbd2d9", "#243b53", "#8a4b00"))
WARNBG, INFOBG = colors.HexColor("#fff4e0"), colors.HexColor("#eaf2fb")
base = getSampleStyleSheet()
cell = ParagraphStyle("cell", parent=base["BodyText"], fontName="Helvetica", fontSize=7.4, leading=9, textColor=INK, alignment=TA_LEFT)
cellb = ParagraphStyle("cellb", parent=cell, fontName="Helvetica-Bold")
small = ParagraphStyle("small", parent=cell, fontSize=7, textColor=MUTED)
hdr = ParagraphStyle("hdr", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
h1 = ParagraphStyle("h1", parent=base["Title"], fontName="Helvetica-Bold", fontSize=18, alignment=TA_LEFT, textColor=HEAD, spaceAfter=4)
h2 = ParagraphStyle("h2", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, textColor=HEAD, spaceBefore=4, spaceAfter=6)
body = ParagraphStyle("body", parent=base["BodyText"], fontName="Helvetica", fontSize=9, leading=12, textColor=INK)
MARK = '<font color="#c05600"><b>[space]</b></font>'


def esc(text):
    return str(text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def shown(text):
    text = str(text or "")
    if not text:
        return '<font color="#9aa5b1">(none)</font>'
    lead, trail = len(text) - len(text.lstrip(" ")), len(text) - len(text.rstrip(" "))
    core = re.sub(r" {2,}", lambda m: " " + MARK * (len(m.group()) - 1) + " ", esc(text.strip(" ")))
    return MARK * lead + core + MARK * trail


def october(name):
    return re.sub(r"_[A-Za-z]{3}\s+Forecast\.xlsx$", "", name)


tabs = defaultdict(list)
for route in FORECAST_ROUTES:
    tabs[(route["individual_tableau_workbook"], route["individual_tableau_worksheet"])].append(route)


def companies(routes):
    return {str(r.get("company_filter") or "").upper() for r in routes} & {"ALIV", "CBL"}


def _values(values):
    values = [str(v).strip() for v in values]
    return ", ".join(values) if len(values) <= 3 else f"{len(values)} values ({', '.join(values[:2])}, ...)"


def selection(route):
    """The lines of the tab that make up this sheet, in plain words."""
    parts = [f"{esc(column.replace('GL ', ''))} = {esc(_values(values))}"
             for column, values in (route.get("individual_include") or {}).items()]
    parts += [f"not {esc(column.replace('GL ', ''))} = {esc(_values(values))}"
              for column, values in (route.get("individual_exclude") or {}).items()]
    return "sums " + "; ".join(parts) if parts else "whole tab total"


def deco(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(0.5 * inch, 0.35 * inch, "Step 3 individual mapping review - generated from "
                      f"automations/forecast/mapping.py on {date.today():%B %d, %Y}")
    canvas.drawRightString(10.5 * inch, 0.35 * inch, f"Page {doc.page}")
    canvas.restoreState()


doc = SimpleDocTemplate(OUT, pagesize=landscape(letter), leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                        topMargin=0.5 * inch, bottomMargin=0.55 * inch, title="Step 3 Individual Mapping Review")
checked = [r for r in FORECAST_ROUTES if r.get("validation_status") == "confirmed"]
need = [r for r in FORECAST_ROUTES if r.get("validation_status") != "confirmed"]
shared = [k for k, rs in tabs.items() if len(companies(rs)) > 1 and any(r.get("validation_status") == "confirmed" for r in rs)]
spaced = [k for k in tabs if k[1] != k[1].strip() or "  " in k[1]]
story = [
    Paragraph("Step 3 Individual Mapping Review", h1),
    Paragraph(f"The {len(tabs)} individual Tableau tabs Step 3 (Validate) compares the {len(FORECAST_ROUTES)} "
              "forecast sheets against. For each sheet, Step 3 opens a temporary copy of the tab, sets Forecast "
              "Start and the Tableau Forecast Type from the page, copies the table and compares its Grand Total "
              "to the sheet's Grand Total (PASS within $1,000).", body),
    Spacer(1, 10),
]
summary = [[Paragraph(x, hdr) for x in ("", "Count")]] + [
    [Paragraph(a, cellb), Paragraph(str(b), cell)] for a, b in (
        ("Individual tabs", len(tabs)),
        ("Forecast sheets", len(FORECAST_ROUTES)),
        ("Checked automatically", len(checked)),
        ("Not checked yet - need a department/line filter", len(need)),
        ("Group tabs shared by ALIV and CBL (each line sums its own company)", len(shared)),
        ("Tab names with an extra space", len(spaced)),
    )]
t = Table(summary, colWidths=[4.2 * inch, 1 * inch], hAlign="LEFT")
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), HEAD), ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                       ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
story += [t, Spacer(1, 12), Paragraph("Please confirm for each tab", h2)]
for item in (
    "the workbook and tab name are right (an orange <b>[space]</b> is a space the real tab name must have too);",
    "the tab shows only what the listed sheets hold, or - for a group tab - only ALIV and CBL split by Company;",
    'the tab has <b>Show Column Grand Totals</b> on (Analysis &gt; Totals), so the app can read its total;',
    "for lines marked <b>Needs a filter</b>: which filter (department, subdepartment or line) makes the tab "
    "match the sheet - or a different tab that already does.",
):
    story += [Paragraph("•&nbsp;&nbsp;" + item, body), Spacer(1, 2)]
story += [Spacer(1, 6), Paragraph("Row colours: <font color='#8a4b00'><b>orange</b></font> = needs a filter or "
                                  "has an extra space; <font color='#2f6fb3'><b>blue</b></font> = group tab, "
                                  "each line sums its own company's rows.", body)]

widths = [2.2, 0.4, 2.9, 0.6, 3.5, 0.4]
for category in FORECAST_CATEGORIES:
    keys = sorted((k for k, rs in tabs.items() if rs[0]["category"] == category), key=lambda k: (k[0].lower(), k[1].lower()))
    lines = sum(len(tabs[k]) for k in keys)
    story += [PageBreak(), Paragraph(f"{category} - {len(keys)} individual tabs, {lines} sheets", h2)]
    data = [[Paragraph(x, hdr) for x in ("Individual Tableau workbook / tab", "Sheets",
                                         "Forecast sheet(s) it checks (company)", "Company filter",
                                         "Step 3 - lines compared", "OK")]]
    style = [("BACKGROUND", (0, 0), (-1, 0), HEAD), ("GRID", (0, 0), (-1, -1), 0.35, LINE),
             ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 3),
             ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]
    for i, key in enumerate(keys, start=1):
        rs = tabs[key]
        folder = str(rs[0].get("individual_tableau_folder") or "").replace("\\", "/")
        group = len(companies(rs)) > 1
        sheets = "<br/>".join(f"{esc(october(r['workbook']))} / <b>{esc(r['excel_sheet'])}</b> "
                              f"<font color='#616e7c'>({esc(r.get('company_filter') or 'no company')})</font>" for r in rs)
        status = []
        for r in rs:
            if r.get("validation_status") == "confirmed":
                status.append(f"<b>{esc(r['excel_sheet'])}</b>: {selection(r)}")
            else:
                status.append(f"<font color='#8a4b00'><b>{esc(r['excel_sheet'])}</b>: needs a filter - "
                              f"{esc(r.get('validation_filter_note') or 'not named')}</font>")
        data.append([
            Paragraph(f"<b>{shown(key[0])}</b><br/><font color='#616e7c'>tab:</font> {shown(key[1])}<br/>"
                      f"<font color='#9aa5b1' size='6.3'>{esc(folder)}</font>", cell),
            Paragraph(str(len(rs)), cell),
            Paragraph(sheets, cell),
            Paragraph("No - kept as saved", cell),
            Paragraph("<br/>".join(status), cell),
            Paragraph("", cell),
        ])
        if any(r.get("validation_status") != "confirmed" for r in rs) or key in spaced:
            style.append(("BACKGROUND", (0, i), (-1, i), WARNBG))
        elif group:
            style.append(("BACKGROUND", (0, i), (-1, i), INFOBG))
    table = Table(data, colWidths=[w * inch for w in widths], repeatRows=1)
    table.setStyle(TableStyle(style))
    story.append(table)

doc.build(story, onFirstPage=deco, onLaterPages=deco)
print("written", OUT)
