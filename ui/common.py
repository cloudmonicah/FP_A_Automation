"""Small helpers shared by the app's pages."""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path
from tkinter import messagebox

log = logging.getLogger(__name__)


def confirm_close_tableau(widget) -> bool:
    """Before a run that opens Tableau: save (or not) and close open workbooks.

    Saving happens here, before the run changes any .twb file; saving later
    would write Tableau's older on-screen values over the run's changes.
    Returns False if the person cancels or a workbook could not be closed.
    """
    try:
        from automations.actuals.tableau_desktop_close_service import (
            close_all_tableau_workbooks,
            list_open_tableau_workbooks,
        )

        open_workbooks = list_open_tableau_workbooks()
    except Exception:
        # Not on Windows, or Tableau can't be inspected: nothing to close.
        return True
    if not open_workbooks:
        return True
    shown = "\n".join(f"  • {t}" for t in open_workbooks[:8])
    more = f"\n  …and {len(open_workbooks) - 8} more" if len(open_workbooks) > 8 else ""
    answer = messagebox.askyesnocancel(
        "Close Tableau workbooks?",
        f"These Tableau workbooks are open and will be closed before the run starts:\n\n"
        f"{shown}{more}\n\n"
        "Save your changes to them first?\n\n"
        "Yes - save and close\nNo - close without saving\nCancel - don't start the run",
        icon="warning",
        parent=widget.winfo_toplevel(),
    )
    if answer is None:
        return False

    top = widget.winfo_toplevel()
    top.configure(cursor="watch")
    top.update_idletasks()
    try:
        result = close_all_tableau_workbooks(save=bool(answer))
    finally:
        top.configure(cursor="")
    if result.get("warning"):
        messagebox.showwarning(
            "Close Tableau workbooks",
            "Some Tableau workbooks are still open, so the run did not start:\n\n"
            + result["warning"].split(": ", 1)[-1]
            + "\n\nSave or close them in Tableau yourself (Tableau may be waiting for "
            "an answer, such as where to save), then click the button again.",
            parent=top,
        )
        return False
    return True


def _flash_taskbar(window) -> None:
    """Flash the taskbar button until the window is focused (Windows only)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes

        class FLASHWINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.UINT),
                ("hwnd", wintypes.HWND),
                ("dwFlags", wintypes.DWORD),
                ("uCount", wintypes.UINT),
                ("dwTimeout", wintypes.DWORD),
            ]

        hwnd = ctypes.windll.user32.GetParent(window.winfo_id()) or window.winfo_id()
        flash_all, flash_until_focused = 0x3, 0xC
        info = FLASHWINFO(ctypes.sizeof(FLASHWINFO), hwnd, flash_all | flash_until_focused, 0, 0)
        ctypes.windll.user32.FlashWindowEx(ctypes.byref(info))
    except Exception:
        log.debug("Could not flash the taskbar", exc_info=True)


def _play_sound(success: bool) -> None:
    if sys.platform != "win32":
        return
    try:
        import winsound

        winsound.MessageBeep(winsound.MB_ICONASTERISK if success else winsound.MB_ICONHAND)
    except Exception:
        pass


def get_attention(widget, *, success: bool = True):
    """Bring the app to the front, flash its taskbar button and play a sound."""
    window = widget.winfo_toplevel()
    try:
        window.deiconify()
        window.lift()
        window.attributes("-topmost", True)
        window.after(1500, lambda: window.attributes("-topmost", False))
        window.focus_force()
    except Exception:
        pass
    _flash_taskbar(window)
    _play_sound(success)
    return window


def notify_finished(widget, title: str, message: str, *, success: bool = True, open_path=None) -> None:
    """Tell the person a run has finished, even if they switched to another window.

    Brings the app to the front, flashes its taskbar button, plays the Windows
    sound, then shows a dialog. With ``open_path``, offers to open that file.
    """
    window = get_attention(widget, success=success)
    if not success:
        messagebox.showerror(title, message, parent=window)
        return
    if open_path and Path(open_path).exists():
        if messagebox.askyesno(title, f"{message}\n\nOpen the report now?", parent=window):
            try:
                os.startfile(str(open_path))  # type: ignore[attr-defined]
            except Exception as exc:
                messagebox.showerror(title, f"Could not open the report:\n{exc}", parent=window)
    else:
        messagebox.showinfo(title, message, parent=window)
