"""Build the Monthly Forecast mapping review PDF from automations/forecast/mapping.py.

    python scripts/mapping_report.py ["Monthly Forecast Mapping Review.pdf"]

Needs reportlab (in requirements-dev.txt). Run it again after changing the mapping.
"""
import re
import sys
from collections import Counter
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from automations.forecast.mapping import FORECAST_CATEGORIES, FORECAST_ROUTES  # noqa: E402

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                TableStyle)

OUT = sys.argv[1] if len(sys.argv) > 1 else "Monthly Forecast Mapping Review.pdf"
INK, MUTED, LINE, HEAD, WARN, WARNBG = (colors.HexColor(c) for c in
                                         ("#1f2933", "#616e7c", "#cbd2d9", "#243b53", "#8a4b00", "#fff4e0"))
base = getSampleStyleSheet()
cell = ParagraphStyle("cell", parent=base["BodyText"], fontName="Helvetica", fontSize=7.4, leading=9,
                      textColor=INK, alignment=TA_LEFT)
cellb = ParagraphStyle("cellb", parent=cell, fontName="Helvetica-Bold")
small = ParagraphStyle("small", parent=cell, fontSize=7, textColor=MUTED)
hdr = ParagraphStyle("hdr", parent=cell, fontName="Helvetica-Bold", textColor=colors.white)
h1 = ParagraphStyle("h1", parent=base["Title"], fontName="Helvetica-Bold", fontSize=18, alignment=TA_LEFT,
                    textColor=HEAD, spaceAfter=4)
h2 = ParagraphStyle("h2", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=12.5, textColor=HEAD,
                    spaceBefore=4, spaceAfter=6)
body = ParagraphStyle("body", parent=base["BodyText"], fontName="Helvetica", fontSize=9, leading=12,
                      textColor=INK)


def esc(text):
    return (str(text or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def shown(text):
    """Escape and make leading/trailing/double spaces visible with a middle dot."""
    text = str(text or "")
    if not text:
        return '<font color="#9aa5b1">(none)</font>'
    lead = len(text) - len(text.lstrip(" "))
    trail = len(text) - len(text.rstrip(" "))
    core = esc(text.strip(" "))
    mark = '<font color="#c05600"><b>[space]</b></font>'
    core = re.sub(r" {2,}", lambda m: " " + mark * (len(m.group()) - 1) + " ", core)
    dot = lambda n: mark * n if n else ""
    return dot(lead) + core + dot(trail)


def october(name):
    return re.sub(r"_[A-Za-z]{3}\s+Forecast\.xlsx$", "_Oct Forecast.xlsx", name)


consolidated_use = Counter(r["tableau_worksheet"].strip().lower() for r in FORECAST_ROUTES)


def flags(route):
    out = []
    for field, label in (("tableau_worksheet", "consolidated tab"), ("excel_sheet", "Excel tab"),
                         ("individual_tableau_worksheet", "individual tab")):
        value = route.get(field) or ""
        if value != value.strip() or "  " in value:
            out.append(f"Extra space in {label} name")
    if consolidated_use[route["tableau_worksheet"].strip().lower()] > 1:
        out.append("Consolidated tab also feeds another sheet")
    if route.get("validation_status") == "needs_filter":
        note = route.get("validation_filter_note")
        out.append("Step 3 needs an extra filter" + (f": {note}" if note else ""))
    return out


def page_decor(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7.5)
    canvas.setFillColor(MUTED)
    canvas.drawString(0.5 * inch, 0.35 * inch,
                      "Monthly Forecast mapping review - generated from automations/forecast/mapping.py "
                      f"on {date.today():%B %d, %Y}")
    canvas.drawRightString(10.5 * inch, 0.35 * inch, f"Page {doc.page}")
    canvas.restoreState()


doc = SimpleDocTemplate(OUT, pagesize=landscape(letter), leftMargin=0.5 * inch, rightMargin=0.5 * inch,
                        topMargin=0.5 * inch, bottomMargin=0.55 * inch,
                        title="Monthly Forecast Mapping Review", author="FP&A Automation")
story = []

# ---- Summary page ----------------------------------------------------------
counts = Counter(r["category"] for r in FORECAST_ROUTES)
status = Counter(r.get("validation_status") for r in FORECAST_ROUTES)
flagged = [r for r in FORECAST_ROUTES if any("space" in f or "another sheet" in f for f in flags(r))]
story += [
    Paragraph("Monthly Forecast Mapping Review", h1),
    Paragraph(f"All {len(FORECAST_ROUTES)} mappings the app uses for Step 2 (Update from Tableau) and "
              "Step 3 (Validate). Each line links one tab in the consolidated FY Forecast Tableau workbook "
              "to one tab in a forecast Excel workbook, and to the individual Tableau workbook and tab that "
              "Step 3 checks it against. Workbook names are shown for October.", body),
    Spacer(1, 10),
]
summary = [[Paragraph(x, hdr) for x in ("Category", "Mappings", "Step 3 ready", "Step 3 needs a filter")]]
for cat in FORECAST_CATEGORIES:
    rows = [r for r in FORECAST_ROUTES if r["category"] == cat]
    summary.append([Paragraph(cat, cellb), Paragraph(str(len(rows)), cell),
                    Paragraph(str(sum(r.get("validation_status") == "confirmed" for r in rows)), cell),
                    Paragraph(str(sum(r.get("validation_status") == "needs_filter" for r in rows)), cell)])
summary.append([Paragraph("Total", cellb), Paragraph(str(len(FORECAST_ROUTES)), cellb),
                Paragraph(str(status["confirmed"]), cellb), Paragraph(str(status["needs_filter"]), cellb)])
t = Table(summary, colWidths=[2.2 * inch, 1.2 * inch, 1.4 * inch, 1.8 * inch], hAlign="LEFT")
t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), HEAD), ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                       ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#eef2f6")),
                       ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 4),
                       ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
story += [t, Spacer(1, 14), Paragraph("How to read the tables", h2),
          Paragraph("<b>Consolidated tab</b>: the tab in the FY Forecast .twb that Step 2 copies. "
                    "<b>Forecast workbook / Excel tab</b>: where Step 2 pastes it. "
                    "<b>Individual workbook / tab</b>: what Step 3 compares the Excel total to, filtered to "
                    "<b>Company</b>. A tick box is left for you to mark each line as checked.", body),
          Spacer(1, 4),
          Paragraph('Names are shown exactly as the app looks for them. An orange <font color="#c05600">'
                    '<b>[space]</b></font> marks a space at the start or end of a name, or an extra space in '
                    "the middle: the real tab name must have it too, so please confirm those.", body),
          Spacer(1, 14), Paragraph("Worth a look first", h2)]
items = [
    "Two consolidated tabs feed two different sheets each: <b>ALIV ALIV Business</b> (Consumer Marketing "
    "&amp; Product Development / Marketing - B2B (ALIV), and ALIV Business / ALIV) and <b>ALIV Direct Cost "
    "(FINANCE) - COMMISSIONS_SERVICE FEES Forecast</b> (two Commissions Service Fees workbooks, with and "
    "without \"(FINANCE)\" in the name). Confirm both copies are intended.",
    f"{sum(1 for r in FORECAST_ROUTES if any('space' in f for f in flags(r)))} mappings have a name with "
    "an extra space (marked in orange in the tables).",
    f"{status['needs_filter']} mappings need an extra department or line filter before Step 3 can "
    "mark them as passed; the filter is named in the Notes column.",
    "Two mappings are switched off in the file and not listed here: a second Media (MARKETING) line "
    "(the same sheet is already filled by the active one) and ALIV SAC COMMISSIONS Monthly Forecast.",
]
for item in items:
    story.append(Paragraph(f"•&nbsp;&nbsp;{item}", body))
    story.append(Spacer(1, 3))

# ---- One table per category ------------------------------------------------
widths = [0.3, 1.85, 2.05, 1.05, 2.25, 0.65, 1.45, 0.4]
for cat in FORECAST_CATEGORIES:
    rows = [r for r in FORECAST_ROUTES if r["category"] == cat]
    story.append(PageBreak())
    story.append(Paragraph(f"{cat} - {len(rows)} mappings", h2))
    data = [[Paragraph(x, hdr) for x in ("#", "Consolidated Tableau tab", "Forecast workbook",
                                         "Excel tab", "Individual Tableau workbook / tab", "Company",
                                         "Notes", "OK")]]
    style = [("BACKGROUND", (0, 0), (-1, 0), HEAD), ("GRID", (0, 0), (-1, -1), 0.35, LINE),
             ("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 3),
             ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 3),
             ("RIGHTPADDING", (0, 0), (-1, -1), 3)]
    for i, r in enumerate(rows, start=1):
        notes = flags(r)
        folder = str(r.get("individual_tableau_folder") or "").replace("\\", "/")
        data.append([
            Paragraph(str(i), small),
            Paragraph(shown(r["tableau_worksheet"]), cellb),
            Paragraph(esc(october(r["workbook"])), cell),
            Paragraph(shown(r["excel_sheet"]), cellb),
            Paragraph(f'{shown(r.get("individual_tableau_workbook"))}<br/>'
                      f'<font color="#616e7c">tab:</font> {shown(r.get("individual_tableau_worksheet"))}<br/>'
                      f'<font color="#9aa5b1" size="6.3">{esc(folder)}</font>', cell),
            Paragraph(esc(r.get("company_filter")), cell),
            Paragraph("<br/>".join(esc(n) for n in notes) or '<font color="#9aa5b1">-</font>',
                      ParagraphStyle("n", parent=cell, textColor=WARN if notes else MUTED)),
            Paragraph('', cell),
        ])
        if any("space" in n or "another sheet" in n for n in notes):
            style.append(("BACKGROUND", (0, i), (-1, i), WARNBG))
        elif i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f7f9fb")))
    table = Table(data, colWidths=[w * inch for w in widths], repeatRows=1)
    table.setStyle(TableStyle(style))
    story.append(table)

doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)
print("written", OUT)
