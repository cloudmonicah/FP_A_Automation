from __future__ import annotations
from shared.friendly_errors import explain

from pathlib import Path
import time

# ============================================================
# TABLEAU DESKTOP CLOSE SERVICE
# ============================================================
#
# PURPOSE
# -------
# Safely close the Tableau Desktop workbook used by the
# Forecast Refresh process after all Tableau extraction and
# Excel workbook updates have completed.
#
# IMPORTANT BUSINESS RULES
# ------------------------
# 1. Do NOT kill every Tableau process.
# 2. Only target the Tableau window associated with the
#    forecast workbook selected by the user.
# 3. Allow Tableau to close normally first.
# 4. Handle a possible save confirmation dialog.
# 5. Do NOT save changes during shutdown because the Forecast
#    Start parameter was already written to the .twb before
#    Tableau Desktop was opened.
# 6. Tableau cleanup should not cause a successful forecast
#    update to be reported as failed.
# ============================================================


# ============================================================
# NORMALIZE TEXT
# ============================================================


def _normalize_text(
    value,
) -> str:
    """
    Normalize text for safe window-title comparisons.
    """

    return " ".join(str(value or "").strip().lower().split())


# ============================================================
# BUILD EXPECTED WORKBOOK TITLE
# ============================================================


def _get_tableau_workbook_name(
    tableau_workbook_path: str | Path,
) -> str:
    """
    Return the Tableau workbook filename without its extension.

    Example:

        C:\\Forecast\\Tableau - FY27 FORECAST WORKBOOKS.twb

    becomes:

        Tableau - FY27 FORECAST WORKBOOKS
    """

    workbook_path = Path(tableau_workbook_path)

    return workbook_path.stem


# ============================================================
# FIND TARGET TABLEAU WINDOW
# ============================================================


def _find_target_tableau_window(
    tableau_workbook_path: str | Path,
):
    """
    Locate the Tableau Desktop window associated with the
    selected forecast workbook.

    The pywinauto import is intentionally lazy so importing this
    service does not initialize UI Automation during normal GUI
    startup.
    """

    # --------------------------------------------------------
    # LAZY IMPORT - DO NOT MOVE TO TOP OF FILE.
    # --------------------------------------------------------

    from pywinauto import Desktop

    workbook_name = _get_tableau_workbook_name(tableau_workbook_path)

    wanted = _normalize_text(workbook_name)

    # --------------------------------------------------------
    # Search visible top-level Windows using UI Automation.
    # --------------------------------------------------------

    desktop = Desktop(backend="uia")

    candidates = []

    for window in desktop.windows():

        try:

            title = window.window_text()

        except Exception:

            continue

        normalized_title = _normalize_text(title)

        if not normalized_title:
            continue

        # ----------------------------------------------------
        # Tableau Desktop titles normally contain the workbook
        # name.
        #
        # Example:
        #
        # Tableau - FY27 FORECAST WORKBOOKS - Tableau
        #
        # We intentionally search by the selected workbook name
        # rather than simply looking for "Tableau".
        # ----------------------------------------------------

        if wanted in normalized_title:

            candidates.append(window)

    if not candidates:

        return None

    # --------------------------------------------------------
    # If more than one matching window exists, prefer a window
    # whose title also contains "Tableau".
    # --------------------------------------------------------

    for window in candidates:

        try:

            title = _normalize_text(window.window_text())

        except Exception:

            continue

        if "tableau" in title:
            return window

    return candidates[0]


# ============================================================
# CHECK WHETHER TARGET WINDOW STILL EXISTS
# ============================================================


def _tableau_window_exists(
    tableau_workbook_path: str | Path,
) -> bool:
    """
    Return True if the selected Tableau workbook still appears
    to have an open Tableau Desktop window.
    """

    try:

        return _find_target_tableau_window(tableau_workbook_path) is not None

    except Exception:

        return False


# ============================================================
# HANDLE TABLEAU SAVE DIALOG
# ============================================================


def _handle_save_dialog() -> bool:
    """
    Look for a Tableau/Windows save confirmation dialog and
    choose the option that closes WITHOUT saving.

    Button captions vary between Tableau/Windows versions, so
    several common captions are supported.

    Returns:
        True  -> a dialog/button was handled
        False -> no supported dialog was found
    """

    # --------------------------------------------------------
    # LAZY IMPORT - DO NOT MOVE TO TOP OF FILE.
    # --------------------------------------------------------

    from pywinauto import Desktop

    desktop = Desktop(backend="uia")

    # --------------------------------------------------------
    # Preferred captions.
    #
    # We intentionally choose a NON-SAVE action.
    # --------------------------------------------------------

    discard_button_names = [
        "Don't Save",
        "Don’t Save",
        "Discard",
        "No",
    ]

    for window in desktop.windows():

        try:

            title = _normalize_text(window.window_text())

        except Exception:

            continue

        # ----------------------------------------------------
        # Only inspect likely Tableau/save confirmation windows.
        # ----------------------------------------------------

        if "tableau" not in title and "save" not in title:
            continue

        for button_name in discard_button_names:

            try:

                button = window.child_window(
                    title=button_name,
                    control_type="Button",
                )

                if button.exists(timeout=0.2):

                    print("Tableau close dialog: " f'clicking "{button_name}".')

                    button.click_input()

                    return True

            except Exception:

                continue

    return False


# ============================================================
# CLOSE TABLEAU FORECAST WORKBOOK
# ============================================================


def close_tableau_forecast_workbook(
    tableau_workbook_path: str | Path,
    wait_seconds: float = 8.0,
) -> dict:
    """
    Close the Tableau Desktop workbook used by Forecast Refresh.

    PARAMETERS
    ----------
    tableau_workbook_path:
        Exact .twb path selected by the user in the GUI.

    wait_seconds:
        Maximum time to wait for Tableau to close normally.

    RETURNS
    -------
    Dictionary containing:

        found
        close_requested
        dialog_handled
        closed
        warning

    IMPORTANT
    ---------
    Failure to close Tableau is returned as a warning rather
    than raising an exception.

    The Excel forecast run may already have completed
    successfully, so a cleanup issue should not incorrectly
    convert that successful run into an "Update Failed" result.
    """

    workbook_path = Path(tableau_workbook_path)

    result = {
        "found": False,
        "close_requested": False,
        "dialog_handled": False,
        "closed": False,
        "warning": None,
    }

    print()
    print("=" * 70)
    print("TABLEAU CLEANUP")
    print("=" * 70)

    print(f"Workbook: {workbook_path}")

    try:

        # ====================================================
        # FIND THE SPECIFIC TABLEAU WINDOW
        # ====================================================

        window = _find_target_tableau_window(workbook_path)

        if window is None:

            # ------------------------------------------------
            # Tableau may already have been closed manually.
            # This is not considered a failure.
            # ------------------------------------------------

            print("Selected Tableau workbook is not open.")

            print("No Tableau cleanup is required.")

            print("=" * 70)

            result["closed"] = True

            return result

        result["found"] = True

        try:

            window_title = window.window_text()

        except Exception:

            window_title = workbook_path.stem

        print(f"Tableau window found: {window_title}")

        # ====================================================
        # REQUEST NORMAL WINDOW CLOSE
        # ====================================================

        print("Requesting Tableau workbook close...")

        window.close()

        result["close_requested"] = True

        # ----------------------------------------------------
        # Give Tableau a moment to either close or display a
        # save confirmation dialog.
        # ----------------------------------------------------

        time.sleep(1.0)

        # ====================================================
        # HANDLE POSSIBLE SAVE CONFIRMATION
        # ====================================================

        dialog_handled = _handle_save_dialog()

        result["dialog_handled"] = dialog_handled

        # ====================================================
        # WAIT FOR TABLEAU TO CLOSE
        # ====================================================

        deadline = time.time() + wait_seconds

        while time.time() < deadline:

            if not _tableau_window_exists(workbook_path):

                result["closed"] = True

                print("Tableau workbook closed successfully.")

                print("=" * 70)

                return result

            # ------------------------------------------------
            # Tableau can sometimes show its confirmation
            # dialog slightly later.
            # ------------------------------------------------

            if not dialog_handled:

                if _handle_save_dialog():

                    dialog_handled = True

                    result["dialog_handled"] = True

            time.sleep(0.5)

        # ====================================================
        # WINDOW STILL OPEN
        # ====================================================

        warning = (
            "Forecast processing completed, but the selected "
            "Tableau workbook did not close automatically. "
            "It may require manual closing."
        )

        result["warning"] = warning

        print("[WARNING] " + warning)

        print("=" * 70)

        return result

    except Exception as exc:

        # ====================================================
        # CLEANUP FAILURE MUST NOT DESTROY SUCCESSFUL RESULTS
        # ====================================================

        warning = (
            "Processing completed, but Tableau could not be closed "
            "automatically. Please close it yourself. "
        )
        explain(exc)

        result["warning"] = warning

        print("[WARNING] " + warning)

        print("=" * 70)

        return result


# ============================================================
# CLOSE EVERY OPEN TABLEAU WORKBOOK
# ============================================================


def _is_tableau_process(process_id) -> bool:
    try:
        from pywinauto.application import process_module

        return Path(process_module(process_id)).name.lower() == "tableau.exe"
    except Exception:
        # If the process can't be inspected, rely on the window title alone.
        return True


def list_open_tableau_workbooks() -> list[str]:
    """Titles of the open Tableau Desktop workbook windows ("Tableau - <name>")."""
    from pywinauto import Desktop

    titles = []
    for window in Desktop(backend="uia").windows():
        try:
            title = window.window_text()
            if not _normalize_text(title).startswith("tableau - "):
                continue
            if not _is_tableau_process(window.process_id()):
                continue
        except Exception:
            continue
        titles.append(title)
    return titles


def close_all_tableau_workbooks(wait_seconds: float = 15.0) -> dict:
    """Close every open Tableau workbook WITHOUT saving.

    Called before each workbook is opened, so only one Tableau workbook is
    ever open during a run. Problems are returned as a warning, not raised.
    """
    from pywinauto import Desktop

    result = {"closed": [], "warning": None}
    try:
        titles = list_open_tableau_workbooks()
        if not titles:
            return result

        print(f"Closing {len(titles)} open Tableau workbook(s) before opening the next one...")
        for window in Desktop(backend="uia").windows():
            try:
                title = window.window_text()
                if title in titles:
                    window.close()
                    time.sleep(0.8)
                    _handle_save_dialog()
            except Exception:
                continue

        deadline = time.time() + wait_seconds
        remaining = titles
        while time.time() < deadline:
            remaining = list_open_tableau_workbooks()
            if not remaining:
                break
            _handle_save_dialog()
            time.sleep(0.5)

        result["closed"] = [t for t in titles if t not in remaining]
        if remaining:
            result["warning"] = "These Tableau workbooks did not close: " + ", ".join(remaining)
            print("[WARNING] " + result["warning"])
        else:
            print("All Tableau workbooks closed.")
    except Exception as exc:
        result["warning"] = "Tableau could not be closed automatically. Please close it yourself."
        print("[WARNING] " + result["warning"])
    return result
