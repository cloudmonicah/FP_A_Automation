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


def test_finish_notification(tmp_path, monkeypatch):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    import ui.common as common
    from app import FinanceAutomationApp

    asked = []
    monkeypatch.setattr(common.messagebox, "askyesno", lambda *a, **k: asked.append(a) or False)
    monkeypatch.setattr(common.messagebox, "showerror", lambda *a, **k: asked.append(a))
    report = tmp_path / "report.xlsx"
    report.write_text("x")

    app = FinanceAutomationApp()
    try:
        app.show_page("actuals")
        view = app.workspace.winfo_children()[0]
        view._finish_actuals_validation({"Revenue": {"overall_status": "PASS", "final_report": str(report)}})
        view._actuals_validation_failed("boom")
        app.app.update()
    finally:
        app.app.destroy()
    assert "Revenue validation finished" in asked[0][0]
    assert "boom" in asked[1][1]
