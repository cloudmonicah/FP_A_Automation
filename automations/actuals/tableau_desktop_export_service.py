from io import StringIO
import csv
import time
import tkinter as tk

from pywinauto import Desktop

from datetime import datetime
from pathlib import Path
import re


def _get_tableau_window(
    max_attempts: int = 3,
    retry_wait_seconds: float = 1.0,
    target_workbook_name: str | None = None,
    tableau_workbook_path: str | Path | None = None,
):

    if tableau_workbook_path:

        target_workbook_name = Path(tableau_workbook_path).stem

    if not target_workbook_name:

        target_workbook_name = "FY27 FORECAST WORKBOOKS"

    # --------------------------------------------------------
    # Normalize the target workbook name once.
    # --------------------------------------------------------

    normalized_target = " ".join(target_workbook_name.strip().lower().split())

    last_titles: list[str] = []

    for attempt in range(
        1,
        max_attempts + 1,
    ):

        # ----------------------------------------------------
        # Refresh the UI Automation desktop tree on every
        # attempt because Tableau windows can change while the
        # workbook is loading.
        # ----------------------------------------------------

        desktop = Desktop(backend="uia")

        tableau_windows = desktop.windows(
            title_re=".*Tableau.*",
            visible_only=True,
        )

        current_titles: list[str] = []

        target_candidates = []

        # ----------------------------------------------------
        # Inspect every visible Tableau window.
        # ----------------------------------------------------

        for window in tableau_windows:

            try:

                title = window.window_text().strip()

            except Exception:

                title = ""

            current_titles.append(title or "[no title]")

            print("Tableau window detected: " f"{title or '[no title]'}")

            # ------------------------------------------------
            # Ignore generic Tableau startup windows.
            # ------------------------------------------------

            if not title:

                continue

            normalized_title = " ".join(title.lower().split())

            if normalized_title == "tableau":

                continue

            # ------------------------------------------------
            # CRITICAL SAFETY CHECK
            #
            # Only consider windows belonging to the selected
            # forecast workbook.
            # ------------------------------------------------

            if normalized_target not in normalized_title:

                continue

            # ------------------------------------------------
            # Confirm the workbook UI is fully loaded.
            #
            # A usable Tableau workbook must expose:
            #
            #     Worksheet
            #     Window
            # ------------------------------------------------

            menu_names = set()

            try:

                menu_controls = window.descendants(control_type="MenuItem")

            except Exception:

                menu_controls = []

            for control in menu_controls:

                try:

                    if not control.is_visible():

                        continue

                    menu_text = control.window_text().strip().lower()

                    if menu_text:

                        menu_names.add(menu_text)

                except Exception:

                    continue

            has_worksheet_menu = "worksheet" in menu_names

            has_window_menu = "window" in menu_names

            if has_worksheet_menu and has_window_menu:

                target_candidates.append(window)

        # ====================================================
        # TARGET WORKBOOK FOUND
        # ====================================================

        if target_candidates:

            # ------------------------------------------------
            # Tableau can occasionally expose more than one UIA
            # representation of the same workbook.
            #
            # Prefer the candidate that is currently visible and
            # enabled. If more than one remains, prefer the
            # foreground window. Otherwise use the first valid
            # target candidate.
            #
            # We NEVER fall back to another workbook.
            # ------------------------------------------------

            usable_candidates = []

            for window in target_candidates:

                try:

                    if window.is_visible() and window.is_enabled():

                        usable_candidates.append(window)

                except Exception:

                    continue

            if not usable_candidates:

                usable_candidates = target_candidates

            workbook_window = None

            # ------------------------------------------------
            # Prefer the target workbook if Windows already
            # considers it the foreground window.
            # ------------------------------------------------

            try:

                import ctypes

                foreground_handle = ctypes.windll.user32.GetForegroundWindow()

            except Exception:

                foreground_handle = None

            if foreground_handle:

                for window in usable_candidates:

                    try:

                        if window.handle == foreground_handle:

                            workbook_window = window

                            break

                    except Exception:

                        continue

            # ------------------------------------------------
            # Otherwise use the first valid representation of
            # the CORRECT workbook.
            # ------------------------------------------------

            if workbook_window is None:

                workbook_window = usable_candidates[0]

            # ------------------------------------------------
            # Bring the selected forecast workbook forward.
            # ------------------------------------------------

            try:

                workbook_window.set_focus()

            except Exception:

                pass

            time.sleep(0.8)

            try:

                selected_title = workbook_window.window_text().strip()

            except Exception:

                selected_title = "Tableau"

            # ------------------------------------------------
            # FINAL SAFETY CHECK
            #
            # Never return a window unless the requested
            # forecast workbook name is still present in the
            # selected title.
            # ------------------------------------------------

            normalized_selected_title = " ".join(selected_title.lower().split())

            if normalized_target not in normalized_selected_title:

                raise RuntimeError(
                    "Tableau workbook safety check "
                    "failed. Expected workbook "
                    f'"{target_workbook_name}", '
                    "but Windows selected "
                    f'"{selected_title}".'
                )

            print(
                "Connected to TARGET Tableau " "workbook window: " f"{selected_title}"
            )

            return workbook_window

        # ====================================================
        # TARGET NOT READY YET
        # ====================================================

        last_titles = current_titles

        if attempt < max_attempts:

            print(
                "Target Tableau workbook "
                f'"{target_workbook_name}" '
                "is not ready yet. "
                f"Waiting... "
                f"({attempt}/{max_attempts})"
            )

            time.sleep(retry_wait_seconds)

    # ========================================================
    # TARGET NEVER FOUND
    # ========================================================

    if not last_titles:

        raise RuntimeError(
            "Tableau Desktop did not become "
            "available after "
            f"{max_attempts} attempts."
        )

    raise RuntimeError(
        "The required Tableau workbook "
        f'"{target_workbook_name}" '
        "could not be found with a ready "
        "Worksheet and Window menu after "
        f"{max_attempts} attempts. "
        "Tableau windows detected: " + "; ".join(last_titles)
    )


def _clear_clipboard() -> None:

    root = tk.Tk()
    root.withdraw()

    try:
        root.clipboard_clear()
        root.update()

    finally:
        root.destroy()


# ============================================================
# READ WINDOWS CLIPBOARD
# ============================================================


def _get_clipboard_text() -> str:
    """
    Read text from the Windows clipboard using the native
    Windows clipboard API.

    Tableau may place copied crosstab data on the clipboard as
    Unicode text, which Tkinter does not always read correctly.
    """

    import ctypes
    from ctypes import wintypes

    # --------------------------------------------------------
    # Windows clipboard format for Unicode text.
    # --------------------------------------------------------
    CF_UNICODETEXT = 13

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    # --------------------------------------------------------
    # Configure Win32 function return types.
    # --------------------------------------------------------
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL

    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL

    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL

    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE

    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID

    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL

    # --------------------------------------------------------
    # Give Tableau a few chances to finish writing to the
    # clipboard before failing.
    # --------------------------------------------------------
    for attempt in range(10):

        # ----------------------------------------------------
        # Check whether Unicode text is available.
        # ----------------------------------------------------
        if user32.IsClipboardFormatAvailable(CF_UNICODETEXT):

            if not user32.OpenClipboard(None):
                time.sleep(0.5)
                continue

            try:
                handle = user32.GetClipboardData(CF_UNICODETEXT)

                if not handle:
                    time.sleep(0.5)
                    continue

                pointer = kernel32.GlobalLock(handle)

                if not pointer:
                    time.sleep(0.5)
                    continue

                try:
                    text = ctypes.wstring_at(pointer)

                finally:
                    kernel32.GlobalUnlock(handle)

                if text.strip():
                    return text

            finally:
                user32.CloseClipboard()

        time.sleep(0.5)

    raise RuntimeError(
        "Tableau executed Copy Crosstab, "
        "but no Unicode text was available "
        "on the Windows clipboard."
    )


# ============================================================
# FIND MENU ITEMS
# ============================================================


def _find_menu_item_in_tableau(
    tableau_window,
    menu_text: str,
):
    """
    Find a menu item that belongs directly to the main
    Tableau window.

    Used for:
        Window
        Worksheet
    """

    target = menu_text.strip().lower()

    for control in tableau_window.descendants(control_type="MenuItem"):

        try:
            text = control.window_text().strip().lower()

            if text == target and control.is_visible():
                return control

        except Exception:
            continue

    return None


def _find_visible_popup_menu_item(
    menu_text: str,
    process_id: int | None = None,
):
    """
    Search all visible desktop windows for a popup menu item.

    Tableau popup menus are often separate UI windows, which
    is why searching only the main Tableau window can fail.
    """

    target = menu_text.strip().lower()

    desktop = Desktop(backend="uia")

    # --------------------------------------------------------
    # Search every current top-level window.
    # --------------------------------------------------------
    for window in desktop.windows():

        if process_id is not None:
            try:
                if window.process_id() != process_id:
                    continue
            except Exception:
                continue

        # ----------------------------------------------------
        # Check the top-level window itself.
        # ----------------------------------------------------
        try:
            if (
                window.element_info.control_type == "MenuItem"
                and window.window_text().strip().lower() == target
                and window.is_visible()
            ):
                return window

        except Exception:
            pass

        # ----------------------------------------------------
        # Search controls contained inside the window.
        # ----------------------------------------------------
        try:
            controls = window.descendants(control_type="MenuItem")

        except Exception:
            continue

        for control in controls:

            try:
                text = control.window_text().strip().lower()

                if text == target and control.is_visible():
                    return control

            except Exception:
                continue

    return None


# ============================================================
# FIND VISIBLE POPUP MENU ITEM CONTAINING TEXT
# ============================================================


def _find_visible_popup_menu_item_contains(
    menu_text: str,
):
    """
    Find a visible Tableau popup-menu item whose text CONTAINS
    the requested text.

    This is useful for Tableau submenu items that may appear as:

        Crosstab to Excel
        Crosstab to Excel...
        Crosstab

    instead of relying on one exact caption.
    """

    target = menu_text.strip().lower()

    desktop = Desktop(backend="uia")

    for window in desktop.windows():

        try:
            controls = window.descendants(control_type="MenuItem")

        except Exception:
            continue

        for control in controls:

            try:
                text = control.window_text().strip().lower()

                if target in text and control.is_visible():
                    return control

            except Exception:
                continue

    return None


def _force_tableau_to_foreground(tableau_window) -> None:
    """
    Force the Tableau Desktop workbook window to become the active
    foreground Windows application.

    pywinauto set_focus() alone is not always enough when this
    automation is launched from the Tkinter application/background
    worker. The Win32 foreground call makes the focus change explicit.
    """

    import ctypes
    import time

    # Get the native Windows handle for the Tableau workbook.
    hwnd = tableau_window.handle

    user32 = ctypes.windll.user32

    # ------------------------------------------------------------
    # Restore Tableau if Windows currently considers it minimized.
    # SW_RESTORE = 9
    # ------------------------------------------------------------
    try:
        user32.ShowWindow(hwnd, 9)
    except Exception:
        pass

    # ------------------------------------------------------------
    # Ask Windows to bring Tableau to the foreground.
    # ------------------------------------------------------------
    try:
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass

    # ------------------------------------------------------------
    # Also use pywinauto's focus method.
    # ------------------------------------------------------------
    try:
        tableau_window.set_focus()
    except Exception:
        pass

    time.sleep(0.8)

    # ------------------------------------------------------------
    # VERIFY which native window actually owns the foreground.
    #
    # This is important. We no longer want to assume set_focus()
    # succeeded.
    # ------------------------------------------------------------
    foreground_hwnd = user32.GetForegroundWindow()

    if foreground_hwnd != hwnd:
        print(
            "WARNING: Tableau did not become the native "
            "foreground window on the first attempt."
        )
    else:
        print("Tableau confirmed as foreground window.")


# ============================================================
# FIND TABLEAU WORKSHEET IN WINDOW POPUP
# ============================================================


def _find_tableau_window_worksheet_control(
    tableau_window,
    worksheet_name: str,
):
    """
    Find a worksheet entry in Tableau's open Window menu.

    IMPORTANT:
    Tableau popup menus may be separate top-level Windows UIA
    elements rather than descendants of the main workbook window.

    The old implementation searched the ENTIRE Windows desktop.
    That allowed unrelated applications such as Teams, OneDrive,
    VS Code, Excel, etc. to enter the search.

    This implementation scopes the search to the SAME Windows
    process that owns the selected Tableau workbook.

    Matching remains exact after whitespace/case normalization.
    """

    import ctypes

    # --------------------------------------------------------
    # Normalize the requested worksheet name.
    # --------------------------------------------------------

    target_name = " ".join(worksheet_name.strip().lower().split())

    # --------------------------------------------------------
    # Get the process ID that owns the selected Tableau window.
    #
    # We derive this from the native Tableau window handle
    # instead of trusting a desktop-wide UI Automation search.
    # --------------------------------------------------------

    tableau_hwnd = tableau_window.handle

    process_id = ctypes.c_ulong()

    ctypes.windll.user32.GetWindowThreadProcessId(
        tableau_hwnd,
        ctypes.byref(process_id),
    )

    tableau_process_id = process_id.value

    if not tableau_process_id:
        raise RuntimeError(
            "Could not determine the Windows process ID "
            "for the selected Tableau workbook."
        )

    print(
        "Restricting worksheet search to Tableau " f"process ID: {tableau_process_id}"
    )

    desktop = Desktop(backend="uia")

    matches = []

    # --------------------------------------------------------
    # Tableau's Window popup may be represented as its own
    # top-level UIA window.
    #
    # We still inspect top-level windows, BUT ONLY windows whose
    # controls belong to the same Tableau process.
    # --------------------------------------------------------

    for candidate_window in desktop.windows(visible_only=True):

        controls = [candidate_window]

        try:
            controls.extend(candidate_window.descendants())
        except Exception:
            pass

        for control in controls:

            try:
                # --------------------------------------------
                # Skip anything not currently visible.
                # --------------------------------------------

                if not control.is_visible():
                    continue

                # --------------------------------------------
                # CRITICAL PROCESS SAFETY CHECK
                #
                # Only accept controls created by the SAME
                # Windows process as the selected Tableau
                # workbook.
                # --------------------------------------------

                control_process_id = control.element_info.process_id

                if control_process_id != tableau_process_id:
                    continue

                item_text = control.window_text().strip()

            except Exception:
                continue

            if not item_text:
                continue

            normalized_item = " ".join(item_text.lower().split())

            # --------------------------------------------
            # Require an EXACT worksheet-name match.
            # --------------------------------------------

            if normalized_item != target_name:
                continue

            try:
                control_type = control.element_info.control_type
            except Exception:
                control_type = "Unknown"

            matches.append(
                (
                    control,
                    control_type,
                    item_text,
                )
            )

    # --------------------------------------------------------
    # No exact Tableau-process match was found.
    # --------------------------------------------------------

    if not matches:
        return None

    print("Exact Tableau-process worksheet " f"matches found: {len(matches)}")

    # --------------------------------------------------------
    # Prefer the control types Tableau normally exposes for
    # Window-menu worksheet entries.
    # --------------------------------------------------------

    preferred_types = (
        "MenuItem",
        "ListItem",
        "Button",
        "Text",
    )

    for preferred_type in preferred_types:

        for (
            control,
            control_type,
            item_text,
        ) in matches:

            if control_type == preferred_type:

                print(
                    "Matched Tableau worksheet control: "
                    f"{item_text} "
                    f"[{control_type}]"
                )

                return control

    # --------------------------------------------------------
    # Tableau may expose another UIA type depending on its
    # current state/version. Exact text + process ownership is
    # enough for the safe fallback.
    # --------------------------------------------------------

    (
        control,
        control_type,
        item_text,
    ) = matches[0]

    print("Matched Tableau worksheet control: " f"{item_text} " f"[{control_type}]")

    return control


# ============================================================
# SELECT TABLEAU WORKSHEET
# ============================================================


def select_tableau_worksheet(
    worksheet_name: str,
    max_attempts: int = 3,
    tableau_workbook_path: str | Path | None = None,
) -> None:
    """
    Select one worksheet in the requested Tableau workbook.

    SAFETY:
    The worksheet popup search is restricted to controls owned
    by the SAME Windows process as the selected Tableau workbook.

    This prevents unrelated desktop applications from entering
    the worksheet search during the full GUI automation.
    """

    print(f"Selecting Tableau worksheet: " f"{worksheet_name}")

    last_error = None

    # ========================================================
    # RETRY THE COMPLETE WORKSHEET-SELECTION PROCESS
    # ========================================================

    for attempt in range(
        1,
        max_attempts + 1,
    ):

        print(f"Worksheet search attempt " f"{attempt}/{max_attempts}...")

        tableau_window = None

        try:

            # =================================================
            # STEP 1
            # GET THE EXACT FORECAST TABLEAU WORKBOOK
            # =================================================

            tableau_window = _get_tableau_window(
                tableau_workbook_path=(tableau_workbook_path),
            )

            # =================================================
            # STEP 2
            # FORCE THE CORRECT TABLEAU WINDOW FORWARD
            # =================================================

            _force_tableau_to_foreground(tableau_window)

            # -------------------------------------------------
            # Close popup menus left over from a previous
            # attempt.
            # -------------------------------------------------

            try:
                tableau_window.type_keys("{ESC}{ESC}")
            except Exception:
                pass

            time.sleep(0.5)

            # ESC can affect focus, so explicitly restore
            # Tableau again.
            _force_tableau_to_foreground(tableau_window)

            # =================================================
            # STEP 3
            # FIND TABLEAU'S WINDOW MENU
            # =================================================

            window_menu = _find_menu_item_in_tableau(
                tableau_window=tableau_window,
                menu_text="Window",
            )

            if window_menu is None:
                raise RuntimeError('Tableau "Window" menu control ' "was not found.")

            print('Opening Tableau "Window" menu...')

            # =================================================
            # STEP 4
            # OPEN THE WINDOW MENU
            # =================================================

            window_menu.click_input()

            # Give Tableau time to create its popup menu.
            time.sleep(1.0)

            print("Searching Tableau-process controls for: " f'"{worksheet_name}"')

            matching_item = None

            # =================================================
            # STEP 5
            # SEARCH ONLY TABLEAU-PROCESS CONTROLS
            # =================================================

            for popup_attempt in range(
                1,
                4,
            ):

                matching_item = _find_tableau_window_worksheet_control(
                    tableau_window=tableau_window,
                    worksheet_name=worksheet_name,
                )

                if matching_item is not None:
                    break

                print(
                    "Waiting for Tableau Window menu "
                    f"controls ({popup_attempt}/3)..."
                )

                time.sleep(0.4)

            # =================================================
            # STEP 6
            # TARGET STILL NOT FOUND
            # =================================================

            if matching_item is None:

                raise RuntimeError(
                    f'Worksheet "{worksheet_name}" '
                    "was not found among controls owned "
                    "by the selected Tableau workbook "
                    "process."
                )

            # =================================================
            # STEP 7
            # CLICK THE WORKSHEET
            # =================================================

            clicked = False

            # -------------------------------------------------
            # Strategy 1:
            # Physical input has been the most reliable method
            # with Tableau's Window menu.
            # -------------------------------------------------

            try:

                matching_item.click_input()

                clicked = True

            except Exception as click_error:

                print(
                    "Physical worksheet click failed; "
                    "trying UIA invoke. "
                    f"{click_error}"
                )

            # -------------------------------------------------
            # Strategy 2:
            # UI Automation Invoke.
            # -------------------------------------------------

            if not clicked:

                try:

                    matching_item.invoke()

                    clicked = True

                except Exception:
                    pass

            # -------------------------------------------------
            # Strategy 3:
            # Standard UIA click.
            # -------------------------------------------------

            if not clicked:

                try:

                    matching_item.click()

                    clicked = True

                except Exception:
                    pass

            if not clicked:

                raise RuntimeError(
                    f'Worksheet "{worksheet_name}" '
                    "was found in Tableau, but Tableau "
                    "would not accept the selection."
                )

            # -------------------------------------------------
            # Allow Tableau to finish rendering the selected
            # worksheet before Crosstab extraction begins.
            # -------------------------------------------------

            time.sleep(2.5)

            print("Tableau worksheet selected: " f"{worksheet_name}")

            return

        except Exception as exc:

            last_error = exc

            print("Worksheet selection attempt " f"{attempt} failed: {exc}")

            # =================================================
            # RESET TABLEAU BEFORE THE NEXT ATTEMPT
            # =================================================

            if tableau_window is not None:

                try:

                    _force_tableau_to_foreground(tableau_window)

                    tableau_window.type_keys("{ESC}{ESC}")

                except Exception:
                    pass

            time.sleep(1.0)

    # ========================================================
    # ALL ATTEMPTS FAILED
    # ========================================================

    raise RuntimeError(
        f'Tableau worksheet "{worksheet_name}" '
        f"could not be selected after "
        f"{max_attempts} attempts. "
        f"Last error: {last_error}"
    )


# ============================================================
# PARSE CROSSTAB
# ============================================================


def _parse_crosstab_text(
    clipboard_text: str,
) -> list[dict]:
    """
    Convert Tableau's tab-separated Crosstab output into
    Python dictionaries.
    """

    if not clipboard_text.strip():
        raise RuntimeError("Tableau copied an empty crosstab.")

    reader = csv.DictReader(
        StringIO(clipboard_text),
        delimiter="\t",
    )

    rows = [dict(row) for row in reader]

    if not rows:
        raise RuntimeError("No data rows were found in the Tableau crosstab.")

    return rows


# ============================================================
# COPY CROSSTAB
# ============================================================


def get_tableau_crosstab_rows(
    worksheet_name: str,
    max_menu_attempts: int = 3,
    raw_output_folder: str | Path | None = None,
    sequence_number: int | None = None,
    tableau_workbook_path: str | Path | None = None,
) -> list[dict]:

    # ========================================================
    # STEP 1 - SELECT REQUESTED TABLEAU WORKSHEET
    # ========================================================

    select_tableau_worksheet(
        worksheet_name=worksheet_name,
        tableau_workbook_path=tableau_workbook_path,
    )

    # --------------------------------------------------------
    # Large summary worksheets need time to render.
    # --------------------------------------------------------

    time.sleep(3)

    # ========================================================
    # STEP 2 - CONNECT TO TABLEAU WORKBOOK WINDOW
    # ========================================================

    # Always reconnect to the exact Tableau workbook selected
    # by the user. Never fall back to another open workbook.
    tableau_window = _get_tableau_window(
        tableau_workbook_path=tableau_workbook_path,
    )

    try:
        tableau_window.set_focus()

    except Exception:
        pass

    time.sleep(1)

    # ========================================================
    # STEP 3 - CLEAR OLD CLIPBOARD
    # ========================================================

    _clear_clipboard()

    print("Copying crosstab from: " f"{worksheet_name}")

    # ========================================================
    # STEP 4 - RETRY COMPLETE MENU CHAIN
    # ========================================================

    crosstab_item = None

    for attempt in range(
        1,
        max_menu_attempts + 1,
    ):

        print(
            "Finding Worksheet -> Copy -> Crosstab " f"{attempt}/{max_menu_attempts}..."
        )

        # ----------------------------------------------------
        # Refresh Tableau window reference every attempt.
        # ----------------------------------------------------

        tableau_window = _get_tableau_window(
            tableau_workbook_path=tableau_workbook_path,
        )

        try:
            tableau_window.set_focus()

        except Exception:
            pass

        # ----------------------------------------------------
        # Close popup menus from any prior attempt.
        # ----------------------------------------------------

        try:
            tableau_window.type_keys("{ESC}{ESC}{ESC}")

        except Exception:
            pass

        time.sleep(0.7)

        # ====================================================
        # OPEN WORKSHEET MENU
        # ====================================================

        worksheet_menu = _find_menu_item_in_tableau(
            tableau_window=(tableau_window),
            menu_text="Worksheet",
        )

        if worksheet_menu is None:

            print('Tableau "Worksheet" menu ' "was not available yet.")

            time.sleep(1)

            continue

        try:

            worksheet_menu.click_input()

        except Exception as exc:

            print("Could not open Worksheet menu: " f"{exc}")

            time.sleep(1)

            continue

        time.sleep(0.8)

        # ====================================================
        # FIND COPY
        # ====================================================

        copy_item = _find_visible_popup_menu_item(
            "Copy", process_id=tableau_window.process_id()
        )

        if copy_item is None:

            print('Tableau "Copy" menu ' "was not visible yet.")

            time.sleep(1)

            continue

        print('Tableau "Copy" menu found.')

        # ====================================================
        # TRY HOVER TO OPEN COPY SUBMENU
        # ====================================================

        try:

            copy_item.move_mouse_input()

            time.sleep(0.8)

        except Exception:
            pass

        crosstab_item = _find_visible_popup_menu_item(
            "Crosstab", process_id=tableau_window.process_id()
        )

        if crosstab_item is not None:

            print('Tableau "Crosstab" menu found.')

            break

        # ====================================================
        # TRY CLICKING COPY
        # ====================================================

        try:

            copy_item.click_input()

            time.sleep(0.8)

        except Exception:
            pass

        crosstab_item = _find_visible_popup_menu_item(
            "Crosstab", process_id=tableau_window.process_id()
        )

        if crosstab_item is not None:

            print('Tableau "Crosstab" menu found.')

            break

        # ====================================================
        # TRY RIGHT ARROW ON COPY
        # ====================================================

        try:

            copy_item.set_focus()

            copy_item.type_keys("{RIGHT}")

            time.sleep(0.8)

        except Exception:
            pass

        crosstab_item = _find_visible_popup_menu_item(
            "Crosstab", process_id=tableau_window.process_id()
        )

        if crosstab_item is not None:

            print('Tableau "Crosstab" menu found.')

            break

        print('Tableau "Crosstab" submenu ' "was not visible yet.")

        time.sleep(1)

    # ========================================================
    # STEP 5 - VALIDATE CROSSTAB MENU WAS FOUND
    # ========================================================

    if crosstab_item is None:

        raise RuntimeError(
            'Tableau "Crosstab" menu item '
            "could not be found after "
            f"{max_menu_attempts} complete "
            "menu attempts."
        )

    # ========================================================
    # STEP 6 - COPY CROSSTAB
    # ========================================================

    crosstab_item.click_input()

    print("Tableau Crosstab copy command executed.")

    time.sleep(4)

    # ========================================================
    # STEP 7 - READ CLIPBOARD
    # ========================================================

    clipboard_text = _get_clipboard_text()

    if raw_output_folder is not None:

        raw_output_path = Path(raw_output_folder)

        raw_output_path.mkdir(
            parents=True,
            exist_ok=True,
        )

        safe_worksheet_name = re.sub(
            r"[^A-Za-z0-9_-]+",
            "_",
            worksheet_name.strip(),
        ).strip("_")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

        if sequence_number is not None:
            raw_filename = (
                f"{sequence_number:03d}_" f"{safe_worksheet_name}_" f"{timestamp}.txt"
            )
        else:
            raw_filename = f"{safe_worksheet_name}_" f"{timestamp}.txt"

        # Shorten long Tableau worksheet names: Windows/OneDrive paths can
        # exceed the classic 260-character path limit.
        import hashlib
        worksheet_id = hashlib.sha1(worksheet_name.encode("utf-8")).hexdigest()[:8]
        short_prefix = f"{sequence_number:03d}_" if sequence_number is not None else ""
        raw_filename = f"{short_prefix}{safe_worksheet_name[:20]}_{worksheet_id}_{timestamp[:15]}.txt"
        raw_file = raw_output_path / raw_filename

        raw_file.write_text(
            clipboard_text,
            encoding="utf-8",
        )

        print("Raw Tableau clipboard saved: " f"{raw_file}")

    # ========================================================
    # STEP 8 - PARSE TABLEAU CROSSTAB
    # ========================================================

    rows = _parse_crosstab_text(clipboard_text)

    print("Tableau rows copied: " f"{len(rows)}")

    # ========================================================
    # STEP 9 - DISPLAY RETURNED COLUMNS
    # ========================================================

    if rows:

        print()

        print("=" * 70)

        print("TABLEAU CROSSTAB COLUMNS")

        print("=" * 70)

        for column in rows[0].keys():

            print(f" - {column}")

        print("=" * 70)

    return rows


# ============================================================
# DIAGNOSTIC: SHOW TABLEAU WORKSHEET MENU ITEMS
# ============================================================


def inspect_tableau_worksheet_menu() -> None:

    # ========================================================
    # STEP 1
    # CONNECT TO TABLEAU
    # ========================================================

    tableau_window = _get_tableau_window()

    # ========================================================
    # STEP 2
    # FIND THE WORKSHEET MENU
    # ========================================================

    worksheet_menu = _find_menu_item_in_tableau(
        tableau_window=tableau_window,
        menu_text="Worksheet",
    )

    if worksheet_menu is None:
        raise RuntimeError('Tableau "Worksheet" menu could not be found.')

    # ========================================================
    # STEP 3
    # OPEN THE WORKSHEET MENU
    # ========================================================

    worksheet_menu.click_input()

    time.sleep(1)

    # ========================================================
    # STEP 4
    # PRINT EVERY VISIBLE MENU ITEM WINDOWS CAN SEE
    # ========================================================

    print()
    print("=" * 70)
    print("VISIBLE TABLEAU MENU ITEMS")
    print("=" * 70)

    desktop = Desktop(backend="uia")

    item_number = 0

    for window in desktop.windows():

        try:
            controls = window.descendants(control_type="MenuItem")
        except Exception:
            continue

        for control in controls:

            try:

                if not control.is_visible():
                    continue

                text = control.window_text()

                item_number += 1

                print(f"{item_number:03d}: " f"[{text}]")

            except Exception:
                continue

    print("=" * 70)

    if item_number == 0:
        print("No visible MenuItem controls were detected.")


# ============================================================
# COPY RAW CROSSTAB TEXT FROM ACTIVE TABLEAU WORKSHEET
# ============================================================


def copy_active_tableau_crosstab_text() -> str:

    # ========================================================
    # STEP 1 - CONNECT TO OPEN TABLEAU WORKBOOK
    # ========================================================

    tableau_window = _get_tableau_window()

    # ========================================================
    # STEP 2 - CLEAR THE CLIPBOARD
    # ========================================================

    _clear_clipboard()

    print("Copying crosstab from active Tableau worksheet...")

    # ========================================================
    # STEP 3 - OPEN WORKSHEET MENU
    # ========================================================

    worksheet_menu = _find_menu_item_in_tableau(
        tableau_window=tableau_window,
        menu_text="Worksheet",
    )

    if worksheet_menu is None:
        raise RuntimeError('Tableau "Worksheet" menu could not be found.')

    worksheet_menu.click_input()

    time.sleep(1)

    # ========================================================
    # STEP 4 - FIND COPY
    # ========================================================

    copy_item = _find_visible_popup_menu_item("Copy")

    if copy_item is None:
        raise RuntimeError('Tableau "Copy" menu item could not be found.')

    copy_item.click_input()

    time.sleep(1)

    # ========================================================
    # STEP 5 - FIND CROSSTAB
    # ========================================================

    crosstab_item = _find_visible_popup_menu_item("Crosstab")

    if crosstab_item is None:
        raise RuntimeError('Tableau "Crosstab" menu item could not be found.')

    crosstab_item.click_input()

    print("Tableau Crosstab copy command executed.")

    # Give Tableau time to populate the clipboard.
    time.sleep(3)

    # ========================================================
    # STEP 6 - RETURN RAW CLIPBOARD TEXT
    # ========================================================

    clipboard_text = _get_clipboard_text()

    if not clipboard_text.strip():
        raise RuntimeError("Tableau copied an empty Crosstab.")

    print("Raw Tableau Crosstab copied successfully.")

    return clipboard_text


# ============================================================
# DIAGNOSTIC: SHOW ITEMS UNDER WORKSHEET -> COPY
# ============================================================


def inspect_tableau_copy_submenu() -> None:
    """
    Open Tableau's Worksheet -> Copy submenu and print every
    visible MenuItem that Windows UI Automation can currently see.

    This does NOT copy data or modify any workbook.
    """

    # ========================================================
    # STEP 1 - CONNECT TO TABLEAU
    # ========================================================

    tableau_window = _get_tableau_window()

    # ========================================================
    # STEP 2 - OPEN WORKSHEET MENU
    # ========================================================

    worksheet_menu = _find_menu_item_in_tableau(
        tableau_window=tableau_window,
        menu_text="Worksheet",
    )

    if worksheet_menu is None:
        raise RuntimeError('Tableau "Worksheet" menu could not be found.')

    worksheet_menu.click_input()

    time.sleep(1)

    # ========================================================
    # STEP 3 - CLICK COPY
    # ========================================================

    copy_item = _find_visible_popup_menu_item("Copy")

    if copy_item is None:
        raise RuntimeError('Tableau "Copy" menu item could not be found.')

    copy_item.click_input()

    time.sleep(1)

    # ========================================================
    # STEP 4 - PRINT ALL VISIBLE MENU ITEMS
    # ========================================================

    print()
    print("=" * 70)
    print("VISIBLE ITEMS AFTER WORKSHEET -> COPY")
    print("=" * 70)

    desktop = Desktop(backend="uia")

    item_number = 0

    for window in desktop.windows():

        try:
            controls = window.descendants(control_type="MenuItem")
        except Exception:
            continue

        for control in controls:

            try:

                if not control.is_visible():
                    continue

                text = control.window_text()

                item_number += 1

                print(f"{item_number:03d}: [{text}]")

            except Exception:
                continue

    print("=" * 70)


# ============================================================
# EXPORT ACTIVE TABLEAU CROSSTAB TO EXCEL
# ============================================================


def export_active_tableau_crosstab_to_excel() -> None:

    # ========================================================
    # CONNECT TO TABLEAU
    # ========================================================

    tableau_window = _get_tableau_window()

    print("Exporting Crosstab from currently active " "Tableau worksheet...")

    # ========================================================
    # TRY THE EXPORT MENU SEQUENCE
    # ========================================================

    crosstab_item = None

    for menu_attempt in range(1, 4):

        # Always bring Tableau back to the foreground before
        # another menu attempt.
        tableau_window.set_focus()
        time.sleep(0.75)

        print(f"Tableau export attempt {menu_attempt}/3...")

        # ----------------------------------------------------
        # Open Worksheet menu.
        # ----------------------------------------------------

        worksheet_menu = _find_menu_item_in_tableau(
            tableau_window=tableau_window,
            menu_text="Worksheet",
        )

        if worksheet_menu is None:
            time.sleep(0.75)
            continue

        worksheet_menu.click_input()
        time.sleep(1)

        # ----------------------------------------------------
        # Find Export.
        # ----------------------------------------------------

        export_item = _find_visible_popup_menu_item("Export")

        if export_item is None:
            try:
                tableau_window.type_keys("{ESC}")
            except Exception:
                pass

            time.sleep(0.75)
            continue

        print('Tableau "Export" menu found.')

        # ----------------------------------------------------
        # STRATEGY 1:
        # Hover over Export so Tableau opens the submenu.
        # ----------------------------------------------------

        try:
            export_item.move_mouse_input()
        except Exception:
            pass

        time.sleep(1.5)

        crosstab_item = _find_visible_popup_menu_item("Crosstab to Excel")

        if crosstab_item is not None:
            break

        # ----------------------------------------------------
        # STRATEGY 2:
        # Use RIGHT ARROW while Export has focus.
        # ----------------------------------------------------

        try:
            export_item.set_focus()
        except Exception:
            pass

        try:
            export_item.type_keys("{RIGHT}")
        except Exception:
            pass

        time.sleep(1.5)

        crosstab_item = _find_visible_popup_menu_item("Crosstab to Excel")

        if crosstab_item is not None:
            break

        # ----------------------------------------------------
        # STRATEGY 3:
        # Click Export. This worked in earlier test runs on
        # this Tableau installation, so keep it as a fallback.
        # ----------------------------------------------------

        try:
            export_item.click_input()
        except Exception:
            pass

        time.sleep(1.5)

        crosstab_item = _find_visible_popup_menu_item("Crosstab to Excel")

        if crosstab_item is not None:
            break

        # ----------------------------------------------------
        # Close any open popup menus before retrying.
        # ----------------------------------------------------

        try:
            tableau_window.type_keys("{ESC}{ESC}")
        except Exception:
            pass

        time.sleep(1)

    # ========================================================
    # VALIDATE THE EXACT MENU ITEM
    # ========================================================

    if crosstab_item is None:
        raise RuntimeError(
            "Tableau reached Worksheet -> Export, but "
            '"Crosstab to Excel" could not be opened after '
            "3 retry attempts."
        )

    actual_menu_text = crosstab_item.window_text().strip()

    print("Tableau Crosstab export option found: " f"{actual_menu_text}")

    # Never allow "Duplicate as Crosstab" or another similar
    # command to be executed accidentally.
    if actual_menu_text.lower() != "crosstab to excel":
        raise RuntimeError(
            "Incorrect Tableau export option detected: " f"{actual_menu_text}"
        )

    # ========================================================
    # EXECUTE EXPORT
    # ========================================================

    crosstab_item.click_input()

    print('Tableau "Crosstab to Excel" command executed.')

    print("Waiting for Tableau to open the exported " "Crosstab in Excel...")

    time.sleep(8)

    print()
    print("=" * 70)
    print("TABLEAU CROSSTAB EXPORT COMPLETE")
    print("=" * 70)
    print("Inspect the Excel workbook opened by Tableau.")
    print("No forecast workbook was modified.")
    print("=" * 70)


# ============================================================
# COPY TABLEAU CROSSTAB USING NATIVE WINDOWS MENU PATH
# ============================================================


def get_tableau_crosstab_rows_native(
    worksheet_name: str,
) -> list[dict]:
    """
    Follow the same steps a user performs manually:

        Window
        -> worksheet name

        Worksheet
        -> Copy
        -> Crosstab

    The Tableau workbook window is first identified using the
    existing UIA logic that is already working.

    Then the same window handle is controlled through the native
    Win32 menu system to avoid unreliable popup-menu searching.
    """

    # ========================================================
    # STEP 1 - FIND THE CORRECT TABLEAU WORKBOOK
    # ========================================================

    # This existing helper has already proven reliable at
    # identifying:
    #
    #     Tableau - FY27 FORECAST WORKBOOKS
    #
    tableau_uia_window = _get_tableau_window()

    # Capture the exact Windows handle for that workbook.
    tableau_handle = tableau_uia_window.handle

    print(f"Using Tableau window handle: {tableau_handle}")

    # ========================================================
    # STEP 2 - ATTACH TO THE SAME WINDOW USING WIN32
    # ========================================================

    tableau_window = Desktop(backend="win32").window(handle=tableau_handle)

    tableau_window.set_focus()

    time.sleep(1)

    # ========================================================
    # STEP 3 - SELECT THE TABLEAU WORKSHEET
    # ========================================================
    #
    # Manual equivalent:
    #
    #     Window
    #       -> ALIV Legal
    #
    # ========================================================

    print(f"Selecting Tableau worksheet: {worksheet_name}")

    tableau_window.menu_select(f"Window->{worksheet_name}")

    time.sleep(2)

    print(f"Tableau worksheet selected: {worksheet_name}")

    # ========================================================
    # STEP 4 - CLEAR OLD CLIPBOARD CONTENT
    # ========================================================

    _clear_clipboard()

    # ========================================================
    # STEP 5 - COPY THE CROSSTAB
    # ========================================================
    #
    # Manual equivalent:
    #
    #     Worksheet
    #       -> Copy
    #           -> Crosstab
    #
    # ========================================================

    print(f"Copying Crosstab from: {worksheet_name}")

    tableau_window.menu_select("Worksheet->Copy->Crosstab")

    print("Tableau Crosstab copy command executed.")

    # Give Tableau time to populate the clipboard.
    time.sleep(3)

    # ========================================================
    # STEP 6 - READ THE WINDOWS CLIPBOARD
    # ========================================================

    clipboard_text = _get_clipboard_text()

    # ========================================================
    # STEP 7 - PARSE THE CROSSTAB
    # ========================================================

    rows = _parse_crosstab_text(clipboard_text)

    print(f"Tableau rows copied: {len(rows)}")

    # ========================================================
    # STEP 8 - DISPLAY RETURNED COLUMNS
    # ========================================================

    if rows:

        print()
        print("=" * 70)
        print("TABLEAU CROSSTAB COLUMNS")
        print("=" * 70)

        for column in rows[0].keys():
            print(f" - {column}")

        print("=" * 70)

    return rows


def save_raw_tableau_clipboard(
    raw_text: str,
    worksheet_name: str,
    output_folder: str | Path,
    sequence_number: int | None = None,
) -> Path:

    # --------------------------------------------------------
    # Create a dedicated diagnostic folder inside the selected
    # forecast folder.
    # --------------------------------------------------------
    raw_folder = Path(output_folder) / "Tableau_Raw_Copies"

    raw_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    safe_worksheet_name = re.sub(
        r"[^A-Za-z0-9_-]+",
        "_",
        worksheet_name.strip(),
    ).strip("_")

    # --------------------------------------------------------
    # Timestamp prevents one run from overwriting another.
    # --------------------------------------------------------
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    if sequence_number is not None:

        filename = (
            f"{sequence_number:03d}_" f"{safe_worksheet_name}_" f"{timestamp}.txt"
        )

    else:

        filename = f"{safe_worksheet_name}_" f"{timestamp}.txt"

    raw_file = raw_folder / filename

    raw_file.write_text(
        raw_text,
        encoding="utf-8",
    )

    print(f"Raw Tableau clipboard saved: {raw_file}")

    return raw_file


# ============================================================
# COPY RAW TABLEAU CROSSTAB TEXT
# ============================================================


def get_tableau_crosstab_text(
    worksheet_name: str,
    max_menu_attempts: int = 3,
    tableau_workbook_path: str | None = None,
) -> str:
    """
    Select a Tableau worksheet, copy its crosstab, and return
    the exact raw tab-delimited clipboard text.

    No parsing is performed here.
    """

    # ========================================================
    # 1. SELECT WORKSHEET
    # ========================================================

    select_tableau_worksheet(
        worksheet_name=worksheet_name, tableau_workbook_path=tableau_workbook_path
    )

    time.sleep(1.5)

    # ========================================================
    # 2. RECONNECT TO TABLEAU
    # ========================================================

    tableau_window = _get_tableau_window(tableau_workbook_path=tableau_workbook_path)

    try:

        tableau_window.set_focus()

    except Exception:

        pass

    time.sleep(0.75)

    # ========================================================
    # 3. CLEAR CLIPBOARD
    # ========================================================

    _clear_clipboard()

    print(f"Copying raw crosstab from: " f"{worksheet_name}")

    # ========================================================
    # 4. FIND WORKSHEET -> COPY
    # ========================================================

    copy_item = None

    for attempt in range(
        1,
        max_menu_attempts + 1,
    ):

        print("Finding Worksheet -> Copy " f"{attempt}/{max_menu_attempts}...")

        tableau_window = _get_tableau_window(
            tableau_workbook_path=tableau_workbook_path
        )

        try:

            tableau_window.set_focus()

        except Exception:

            pass

        time.sleep(0.5)

        try:

            tableau_window.type_keys("{ESC}{ESC}")

        except Exception:

            pass

        time.sleep(0.4)

        worksheet_menu = _find_menu_item_in_tableau(
            tableau_window=(tableau_window),
            menu_text=("Worksheet"),
        )

        if worksheet_menu is None:

            time.sleep(0.75)

            continue

        try:

            worksheet_menu.click_input()

        except Exception:

            time.sleep(0.75)

            continue

        time.sleep(1)

        copy_item = _find_visible_popup_menu_item("Copy")

        if copy_item is not None:

            print('Tableau "Copy" menu found.')

            break

        try:

            tableau_window.type_keys("{ESC}{ESC}")

        except Exception:

            pass

        time.sleep(0.75)

    if copy_item is None:

        raise RuntimeError('Tableau "Copy" menu could not be found.')

    # ========================================================
    # 5. OPEN COPY SUBMENU
    # ========================================================

    copy_item.click_input()

    time.sleep(1)

    # ========================================================
    # 6. FIND CROSSTAB
    # ========================================================

    crosstab_item = None

    for attempt in range(
        1,
        max_menu_attempts + 1,
    ):

        crosstab_item = _find_visible_popup_menu_item("Crosstab")

        if crosstab_item is not None:
            break

        time.sleep(0.75)

    if crosstab_item is None:

        raise RuntimeError('Tableau "Crosstab" menu could not be found.')

    # ========================================================
    # 7. COPY CROSSTAB
    # ========================================================

    crosstab_item.click_input()

    print("Tableau Crosstab copy command executed.")

    time.sleep(3)

    # ========================================================
    # 8. RETURN RAW CLIPBOARD CONTENT
    # ========================================================

    clipboard_text = _get_clipboard_text()

    if not clipboard_text.strip():

        raise RuntimeError("Tableau copied an empty crosstab.")

    print("Raw Tableau crosstab captured.")

    return clipboard_text
