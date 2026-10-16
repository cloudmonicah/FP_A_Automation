"""Open the app and every page once, to catch start-up and layout errors."""

import os
import sys

import pytest

tkinter = pytest.importorskip("tkinter")
pytest.importorskip("ttkbootstrap")

if sys.platform != "win32" and not os.environ.get("DISPLAY"):
    pytest.skip("needs a display", allow_module_level=True)


@pytest.fixture(scope="module")
def app(tmp_path_factory):
    """One app window for the whole module.

    Creating a second Tk interpreter in the same process fails intermittently
    on the GitHub Windows runners ("Can't find a usable init.tcl"), so the
    tests share one window instead of building a new one each.
    """
    with pytest.MonkeyPatch.context() as patch:
        patch.setenv("APPDATA", str(tmp_path_factory.mktemp("appdata")))
        from app import FinanceAutomationApp

        instance = FinanceAutomationApp()
        try:
            yield instance
        finally:
            instance.app.destroy()


def test_every_page_opens(app):
    errors = []
    app._report_error = lambda *exc: errors.append(exc)
    app.app.report_callback_exception = app._report_error
    for page in ("actuals", "budget", "forecast"):
        app.show_page(page)
        app.app.update()
        assert app.current_page == page
    assert errors == []


def test_finish_notification(app, tmp_path, monkeypatch):
    import ui.common as common

    asked = []
    monkeypatch.setattr(common.messagebox, "askyesno", lambda *a, **k: asked.append(a) or False)
    monkeypatch.setattr(common.messagebox, "showerror", lambda *a, **k: asked.append(a))
    report = tmp_path / "report.xlsx"
    report.write_text("x")

    app.show_page("actuals")
    view = app.workspace.winfo_children()[0]
    view._finish_actuals_validation({"Revenue": {"overall_status": "PASS", "final_report": str(report)}})
    view._actuals_validation_failed("boom")
    app.app.update()
    assert "Revenue validation finished" in asked[0][0]
    assert "boom" in asked[1][1]
