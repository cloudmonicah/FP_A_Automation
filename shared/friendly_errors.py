"""Turn technical errors into plain language for the people who run the app.

Every message shown in a report, a note or a dialog goes through ``friendly``.
The technical text is never shown to the user; ``explain`` prints it to the
console / log only, so support can still find the cause.
"""

from __future__ import annotations

import re

_RULES: list[tuple[re.Pattern, object]] = []


class PlainLanguageError(Exception):
    """An error whose message was written for the user; shown as is."""


def _rule(pattern, message):
    _RULES.append((re.compile(pattern, re.I | re.S), message))


def _quoted(text):
    found = re.search(r'"([^"]+)"', text)
    return found.group(1) if found else None


# --- Tableau ---------------------------------------------------------------
_rule(
    r'Tableau "Copy" menu',
    "Tableau did not respond when the table was copied. Make sure Tableau is "
    "open, shows the report and has no pop-up window, then run again.",
)
_rule(
    r'worksheet "(?P<tab>.+?)" could not be selected|Worksheet "(?P<tab2>.+?)" was not found among',
    lambda m, t: (
        f'Tableau could not open the tab "{m.group("tab") or m.group("tab2")}". '
        "The tab may have been renamed, or Tableau is still busy. "
        "Close any open Tableau menu or window and run again."
    ),
)
_rule(
    r"Worksheet not found: (?P<tab>.+)",
    lambda m, t: (
        f'The Tableau workbook does not contain a tab called "{m.group("tab").strip()}". '
        "Check that the workbook and tab name in the mapping are correct."
    ),
)
_rule(
    r"Tableau window|Tableau workbook .* not ready|not ready yet",
    "Tableau did not open in time. Close Tableau, then run again.",
)
_rule(
    r"Expected exactly one new raw Tableau",
    "The table copied from Tableau could not be saved. Run again; if it keeps "
    "happening, check that the output folder is not full or locked.",
)
_rule(
    r"does not contain the mapped Revenue tab",
    lambda m, t: "A Tableau workbook is missing a tab the report needs "
    f"({t.split(' does not contain')[0]}). Open the workbook and check the tab name.",
)
_rule(
    r"Select the Revenue \w+ TWB",
    "A Tableau workbook for Revenue was not found. Check that it is in the "
    "Workbooks folder and has the expected name.",
)
_rule(
    r"No mapped posting-date filter",
    "The Tableau tab has no posting-date filter, so the month cannot be set. "
    "Check that the tab is the right one.",
)
_rule(
    r"has no unique worksheet matching '(?P<tab>[^']+)'",
    lambda m, t: "The Tableau workbook does not have the expected tab "
    f'"{m.group("tab")}". Check the tab name in the workbook.',
)

# --- Tableau table contents --------------------------------------------------
_rule(
    r"Expected one Revenue subtotal for (?P<what>\{.*?\}); found (?P<n>\d+)",
    lambda m, t: (
        "The row needed for this line was not found in the Tableau table "
        "(the table layout or this line's mapping may have changed)."
        if m.group("n") == "0"
        else "More than one matching row was found in the Tableau table, so the "
        "right one cannot be picked. The mapping for this line needs review."
    ),
)
_rule(
    r"No Media rows found",
    "No Media rows were found in the Tableau table. Check the Media tab and filters.",
)
_rule(
    r"Media rows add to (?P<a>[\d,.]+) but Grand Total is (?P<b>[\d,.]+)",
    lambda m, t: f"The Media rows add up to {m.group('a')} but Tableau's Grand Total "
    f"shows {m.group('b')}. The Media rows picked may be incomplete.",
)
_rule(
    r"month headers|Missing Actual YTD reporting months|unexpected month",
    "The Tableau table shows different months than the ones selected. Check the "
    "month filter in the workbook.",
)
_rule(
    r"YTD does not reconcile",
    "The monthly figures in the Tableau table do not add up to its year-to-date "
    "total. Check the month filter in the workbook.",
)
_rule(
    r"A mapped Revenue actual is blank|Missing numeric|Actual extraction",
    "A number expected in the Tableau table was blank.",
)
_rule(
    r"Grand Total includes other companies|unambiguous company total|Expected one Grand Total",
    "The company total could not be found in the Tableau table. The table may "
    "include other companies or a different layout than expected.",
)
_rule(
    r"Incomplete Tableau crosstab|Misaligned Tableau|crosstab was empty|No Tableau .* rows",
    "The table copied from Tableau was empty or incomplete. Make sure the report "
    "has data for the selected month and run again.",
)

# --- Financial statements ----------------------------------------------------
_rule(
    r"statement period does not match|Statement of Earnings period does not match",
    "The financial statement is for a different month than the one selected. "
    "Select the matching statement file or change the closed month.",
)
_rule(
    r"Expected one current-year actual month column",
    "The financial statement does not have a column for the selected month.",
)
_rule(
    r"statement line missing/ambiguous|Duplicate CBL|Duplicate .* statement line",
    lambda m, t: "A line in the financial statement is missing or appears more than "
    f'once{(" (" + t.split(":")[-1].strip() + ")") if ":" in t else ""}.',
)
_rule(
    r"Missing financial statement|Missing numeric CBL statement",
    "The financial statement has no figure for this line.",
)
_rule(
    r"Missing consolidated",
    "The consolidated figure for this line is missing.",
)

# --- Files -------------------------------------------------------------------
_rule(
    r"PermissionError|open in another program|being used by another process",
    "A file could not be saved because it is open in another program (for "
    "example Excel). Close it and run again.",
)
_rule(
    r"FileNotFoundError|No such file|was not found:",
    "A file or folder could not be found. It may have been moved, renamed or "
    "deleted. Check the file selections and run again.",
)
_rule(
    r"UnicodeDecodeError|BadZipFile|not a zip file|InvalidFileException",
    "A file could not be read. It may be damaged or not the right type of file.",
)

GENERIC = (
    "Something unexpected went wrong. The details were saved in the log file; "
    "please send it to support."
)


def friendly(error, context: str = "") -> str:
    """Plain-language explanation for an exception or an error string."""
    text = str(error or "").strip()
    if isinstance(error, PlainLanguageError):
        return f"{context}: {text}" if context else text
    name = type(error).__name__ if isinstance(error, BaseException) else ""
    haystack = f"{name} {text}"
    for pattern, message in _RULES:
        match = pattern.search(haystack)
        if match:
            result = message(match, text) if callable(message) else message
            break
    else:
        result = GENERIC
    return f"{context}: {result}" if context else result


def explain(error, context: str = "") -> str:
    """Friendly text for the user; the technical text goes to the console/log."""
    text = friendly(error, context)
    name = type(error).__name__ if isinstance(error, BaseException) else "Error"
    print(f"[details] {context + ' | ' if context else ''}{name}: {error}", flush=True)
    return text
