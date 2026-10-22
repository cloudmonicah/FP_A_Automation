"""Closing Tableau: save only when the person asks, before a run; never during it."""

import sys
import types

from automations.actuals import tableau_desktop_close_service as close


class FakeButton:
    def __init__(self, dialog, name):
        self.dialog, self.name = dialog, name

    def exists(self, timeout=0):
        return self.dialog.open and self.name in self.dialog.buttons

    def click_input(self):
        self.dialog.clicked.append(self.name)
        self.dialog.open = False
        self.dialog.workbook.open = False


class FakeDialog:
    def __init__(self, workbook):
        self.workbook, self.open, self.clicked = workbook, True, []
        self.buttons = ["Save", "Don't Save", "Cancel"]

    def window_text(self):
        return "Tableau"

    def child_window(self, title, control_type):
        return FakeButton(self, title)


class FakeWorkbook:
    def __init__(self, title, dirty=True, stuck=False):
        self.title, self.dirty, self.stuck, self.open = title, dirty, stuck, True
        self.dialog = None

    def window_text(self):
        return self.title

    def process_id(self):
        return 1

    def close(self):
        if self.stuck:
            return                      # e.g. Tableau wants to know where to save
        if self.dirty:
            self.dialog = FakeDialog(self)
        else:
            self.open = False


def _desktop(workbooks):
    class Desktop:
        def __init__(self, backend):
            pass

        def windows(self, **kwargs):
            out = [w for w in workbooks if w.open]
            out += [w.dialog for w in workbooks if w.dialog and w.dialog.open]
            return out

    module = types.ModuleType("pywinauto")
    module.Desktop = Desktop
    application = types.ModuleType("pywinauto.application")
    application.process_module = lambda pid: "tableau.exe"
    return {"pywinauto": module, "pywinauto.application": application}


def _run(monkeypatch, workbooks, save):
    for name, module in _desktop(workbooks).items():
        monkeypatch.setitem(sys.modules, name, module)
    monkeypatch.setattr(close.time, "sleep", lambda s: None)
    return close.close_all_tableau_workbooks(wait_seconds=0.2, save=save)


def test_yes_saves_changes_then_closes(monkeypatch):
    book = FakeWorkbook("Tableau - FY27 FORECAST WORKBOOKS")
    result = _run(monkeypatch, [book], save=True)
    assert book.dialog.clicked == ["Save"] and not book.open
    assert result["closed"] == ["Tableau - FY27 FORECAST WORKBOOKS"] and not result["warning"]


def test_runs_close_without_saving(monkeypatch):
    book = FakeWorkbook("Tableau - Legal_Validation_ab12")
    _run(monkeypatch, [book], save=False)
    assert book.dialog.clicked == ["Don't Save"] and not book.open


def test_a_workbook_that_will_not_close_is_reported(monkeypatch):
    stuck = FakeWorkbook("Tableau - Book1", stuck=True)
    result = _run(monkeypatch, [stuck, FakeWorkbook("Tableau - Legal", dirty=False)], save=True)
    assert result["warning"] == "These Tableau workbooks did not close: Tableau - Book1"
