"""Plain-language result messages for the Monthly Forecast page (no Tk, so testable)."""

from __future__ import annotations

import re


def step2_message(category: str, result: dict) -> tuple[bool, str]:
    """(needs attention, text) for the Step 2 pop-up, in plain words.

    Only things the user has to act on are shown as problems (with the
    warning icon). Housekeeping - months Tableau had no figures for,
    amounts that showed as dates - is listed as information. Merged cells
    that were unmerged are only written to the run log.
    """
    month_suffix = re.compile(r"_[A-Za-z]{3} Forecast\.xlsx$")

    def name(item):
        return month_suffix.sub("", item["workbook"]) + " / " + item["worksheet"]

    passed = result.get("passed", [])
    failed = result.get("failed", [])
    retried = result.get("check_manually") or []
    dropped = [(item, entry) for item in passed for entry in item.get("comments_dropped") or []]
    total = len(passed) + len(failed)

    if failed:
        not_updated = ("1 sheet was not updated and is unchanged." if len(failed) == 1
                       else f"{len(failed)} sheets were not updated and are unchanged.")
        lines = [f"{category}: {len(passed)} of {total} sheets updated. {not_updated}"]
    else:
        lines = [f"{category}: all {len(passed)} sheets updated in "
                 f"{result.get('workbooks_updated', 0)} workbooks. Formulas were kept."]

    attention = []
    for number, item in enumerate(failed[:10], start=1):
        attention += [f"{number}. {name(item)}", f"    {item['error']}"]
    if len(failed) > 10:
        attention.append(f"...and {len(failed) - 10} more (see the run log).")
    if retried:
        attention.append("Check these by hand - they were interrupted once and worked on the second try:")
        attention += [f"    - {name(item)}" for item in retried[:10]]
    if dropped:
        attention.append("These lines are no longer in Tableau, so their comments were removed with them:")
        attention += [f"    - {name(item)}: {entry['account']} - \"{entry['text'][:80]}\""
                      for item, entry in dropped[:10]]
    if attention:
        lines += ["", "Needs your attention", *attention]

    info = []
    blank = [item for item in passed if item.get("months_blank_in_tableau")]
    if blank:
        info.append("Tableau had no figures for these months, so they are empty in Excel:")
        info += [f"    - {name(item)}: {', '.join(item['months_blank_in_tableau'])}" for item in blank[:10]]
        if len(blank) > 10:
            info.append(f"    ...and {len(blank) - 10} more (see the run log).")
    fixed = sum(item.get("date_formats_fixed", 0) for item in passed)
    if fixed:
        info.append(f"{fixed} amount(s) were showing as dates and now show as numbers.")
    if info:
        lines += ["", "For your information", *info]

    if result.get("validation_folder"):
        lines += ["", f"Tableau extract and run log: {result['validation_folder']}"]
    if not failed:
        lines += ["", f"Next: click Validate under {category} (step 3)."]
    return bool(attention), "\n".join(lines)


def step3_message(category: str, result: dict) -> tuple[bool, str]:
    """(needs attention, text) for the Step 3 pop-up, in plain words."""
    month_suffix = re.compile(r"_[A-Za-z]{3} Forecast\.xlsx$")

    def name(item):
        return month_suffix.sub("", str(item.get("forecast_workbook") or "")) + " / " + str(item.get("worksheet") or "")

    def money(value):
        return "-" if value is None else f"{value:,.0f}"

    rows = [item for item in result.get("results", []) if item.get("category", category) == category]
    status = lambda wanted: [item for item in rows if item.get("status") == wanted]
    passed, failed, action, skipped = (status("PASS"), status("FAIL"),
                                       status("ACTION REQUIRED"), status("NOT VALIDATED"))

    if failed or action:
        lines = [f"{category}: {len(passed)} of {len(passed) + len(failed) + len(action)} "
                 "checked sheets match Tableau (within $1,000)."]
    elif passed:
        lines = [f"{category}: all {len(passed)} checked sheets match Tableau (within $1,000)."]
    else:
        lines = [f"{category}: no sheets could be checked automatically."]

    attention = []
    number = 0
    for item in failed[:10]:
        number += 1
        attention += [f"{number}. {name(item)}",
                      f"    Excel {money(item.get('excel_grand_total'))} vs Tableau "
                      f"{money(item.get('tableau_grand_total'))} - off by {money(abs(item.get('variance') or 0))}"]
    for item in action[:10]:
        number += 1
        attention += [f"{number}. {name(item)}", f"    Not checked: {item.get('note') or 'see the report'}"]
    if len(failed) + len(action) > 20:
        attention.append("...and more (see the report).")
    if attention:
        lines += ["", "Needs your attention", *attention]

    info = []
    if skipped:
        info.append(f"{len(skipped)} sheet(s) are not checked automatically yet - their Tableau tab "
                    "needs a department filter. They are listed in the report.")
    if result.get("conflict_copies"):
        info.append("OneDrive also has another copy of the report (" + ", ".join(result["conflict_copies"])
                    + "), saved while the report was open somewhere else. Teams may be showing that "
                    "older copy. Delete it once the report shows this run's results.")
    if info:
        lines += ["", "For your information", *info]

    if result.get("validation_report"):
        lines += ["", f"Validation report: {result['validation_report']}",
                  "Its Summary shows Last Updated; use Open Validation Report on this page to open it."]
    return bool(attention), "\n".join(lines)
