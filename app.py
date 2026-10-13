from __future__ import annotations

import logging
from tkinter import messagebox

import ttkbootstrap as tb

from config import (
    APP_MIN_SIZE,
    APP_SIZE,
    APP_THEME,
    APP_TITLE,
    TEST_MODE,
)

from ui.forecast_refresh_view import ForecastRefreshView
from ui.actuals_validation_view import ActualsValidationView
from ui.budget_setup_view import BudgetSetupView
from ui.settings_view import SettingsView
from shared.logging_setup import log_file

log = logging.getLogger(__name__)


class FinanceAutomationApp:

    def __init__(self) -> None:
        self.app = tb.Window(
            title=APP_TITLE,
            themename=APP_THEME,
            size=APP_SIZE,
            minsize=APP_MIN_SIZE,
        )

        self.current_page = None
        self.nav_buttons = {}
        self.app.report_callback_exception = self._report_error

        self._build_header()
        self._build_body()

    def _build_header(self) -> None:
        header = tb.Frame(
            self.app,
            padding=(24, 16),
        )
        header.pack(fill="x")

        title_block = tb.Frame(header)
        title_block.pack(side="left")

        tb.Label(
            title_block,
            text="FINANCE AUTOMATION CENTER",
            font=("Segoe UI", 24, "bold"),
        ).pack(anchor="w")

        tb.Label(
            title_block,
            text="FP&A workbook and reporting automation",
            font=("Segoe UI", 10),
            bootstyle="secondary",
        ).pack(anchor="w")

        if TEST_MODE:
            tb.Label(
                header,
                text="TEST MODE",
                font=("Segoe UI", 10, "bold"),
                bootstyle="warning-inverse",
                padding=(12, 6),
            ).pack(side="right")

    def _build_body(self) -> None:
        body = tb.Frame(self.app)
        body.pack(fill="both", expand=True)

        sidebar = tb.Frame(
            body,
            width=200,
            padding=(18, 20),
        )
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tb.Label(
            sidebar,
            text="AUTOMATIONS",
            font=("Segoe UI", 11, "bold"),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(0, 14))

        pages = [
            ("forecast", "Monthly Forecast"),
            ("actuals", "Actuals Validation"),
            ("budget", "Initial Budget Setup"),
            ("settings", "Settings"),
        ]

        for key, label in pages:
            button = tb.Button(
                sidebar,
                text=label,
                bootstyle="secondary-outline",
                command=lambda page=key: self.show_page(page),
            )
            button.pack(fill="x", pady=5)
            self.nav_buttons[key] = button

        tb.Label(
            sidebar,
            text="FP&A Automation",
            font=("Segoe UI", 9),
            bootstyle="secondary",
        ).pack(
            side="bottom",
            anchor="w",
            pady=(20, 0),
        )

        self.workspace = tb.Frame(
            body,
            padding=(20, 16, 24, 24),
        )
        self.workspace.pack(
            side="right",
            fill="both",
            expand=True,
        )

        self.show_page("forecast")

    def show_page(self, page: str) -> None:
        if self.current_page == page:
            return

        views = {
            "forecast": ForecastRefreshView,
            "actuals": ActualsValidationView,
            "budget": BudgetSetupView,
            "settings": SettingsView,
        }

        view_class = views[page]

        # Create the next view before removing the current one.
        new_view = view_class(parent=self.workspace)

        for child in self.workspace.winfo_children():
            if child is not new_view:
                child.destroy()

        new_view.pack(fill="both", expand=True)
        self.current_page = page

        for key, button in self.nav_buttons.items():
            button.configure(
                bootstyle=("primary" if key == page else "secondary-outline")
            )

    def _report_error(self, exc_type, exc, tb) -> None:
        """Log any error raised in a button handler and tell the user."""
        log.error("Unhandled error in the app", exc_info=(exc_type, exc, tb))
        messagebox.showerror(
            "Something went wrong",
            f"{exc}\n\nDetails were written to the log file:\n{log_file()}",
            parent=self.app,
        )

    def run(self) -> None:
        self.app.mainloop()


if __name__ == "__main__":
    from main import main

    main()
