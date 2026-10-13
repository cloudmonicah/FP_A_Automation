"""Open the app and every page once, to catch start-up and layout errors."""

import os
import sys

import pytest

tkinter = pytest.importorskip("tkinter")
pytest.importorskip("ttkbootstrap")

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    pytest.skip("needs a display", allow_module_level=True)


def test_every_page_opens(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    from app import FinanceAutomationApp

    app = FinanceAutomationApp()
    errors = []
    app._report_error = lambda *exc: errors.append(exc)
    app.app.report_callback_exception = app._report_error
    try:
        for page in ("actuals", "budget", "settings", "forecast"):
            app.show_page(page)
            app.app.update()
            assert app.current_page == page
    finally:
        app.app.destroy()
    assert errors == []
