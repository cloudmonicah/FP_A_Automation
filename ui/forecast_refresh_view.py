from __future__ import annotations

import logging
import os
import threading
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from automations.forecast.mapping import FORECAST_CATEGORIES, routes_for
from automations.forecast.preflight import check_before_update, check_before_validation
from automations.forecast.service import (
    build_forecast_period,
    fiscal_year_for,
    populate_monthly_forecast_from_tableau,
    prepare_monthly_forecast,
    validate_monthly_forecast,
)
from shared import settings
from shared.forecast_output_paths import get_forecast_validation_output_folder
from shared.forecast_validation_report import counts as result_counts
from shared.forecast_validation_report import load_saved_results, report_path_for
from shared.friendly_errors import explain, friendly
from ui.common import confirm_close_tableau, get_attention

MONTH_LOOKUP = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}

VALIDATION_TOLERANCE = 5.00

log = logging.getLogger(__name__)


class ForecastRefreshView(tb.Frame):
    """
    Monthly Forecast GUI.

    One selected forecast period (month + year; the fiscal year follows from
    it) controls:
      - Excel Forecast Type
      - consolidated Tableau Forecast Start / Forecast Type
      - individual Tableau validation period

    Threading rule: worker threads never touch Tk objects (no .get() on
    variables or widgets, no new Tk objects). Everything a worker needs is
    read on the UI thread first and passed in; results come back through
    ``self._post`` (``self.after``).
    """

    def __init__(self, parent) -> None:
        super().__init__(parent)

        # Start from the paths used last time (or the team defaults).
        remembered = self._remembered_paths()
        self.source_folder_var = tb.StringVar(value=remembered["forecast_source_folder"])
        self.output_folder_var = tb.StringVar(value=remembered["forecast_output_folder"])
        self.tableau_workbook_var = tb.StringVar(value=remembered["forecast_tableau_workbook"])
        self.overwrite_var = tb.StringVar(value="skip")

        self.status_var = tb.StringVar(value="Ready")
        self.status_description_var = tb.StringVar()
        self.period_summary_var = tb.StringVar()
        self.progress_var = tb.IntVar(value=0)

        self.files_created_var = tb.StringVar(value="0")
        self.worksheets_updated_var = tb.StringVar(value="0")
        self.validation_passed_var = tb.StringVar(value="0")
        self.validation_failed_var = tb.StringVar(value="0")
        self.action_required_var = tb.StringVar(value="0")
        self.not_validated_var = tb.StringVar(value="0")
        self.output_var = tb.StringVar(value="Not created yet")
        self.validation_output_var = tb.StringVar()

        self._busy = False
        self._report_path: Path | None = None
        self._step_buttons: list = []
        # Per category: (step 2 status, step 3 status)
        self._category_vars = {
            name: (tb.StringVar(value="—"), tb.StringVar(value="—"))
            for name in (*FORECAST_CATEGORIES, "All")
        }

        self._build_view()
        self._on_period_changed()

    # =========================================================
    # VIEW
    # =========================================================

    def _build_view(self) -> None:
        tb.Label(
            self,
            text="Monthly Forecast Automation",
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")

        tb.Label(
            self,
            text=(
                "Create the monthly forecast, populate it from Tableau, "
                "and validate the final totals."
            ),
            font=("Segoe UI", 10),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(3, 10))

        content = tb.Frame(self)
        content.pack(fill="both", expand=True)

        # 60% Forecast Parameters / 40% Process Status
        content.columnconfigure(0, weight=3, uniform="panels")
        content.columnconfigure(1, weight=2, uniform="panels")
        content.rowconfigure(0, weight=1)

        self._build_parameters_panel(content)
        self._build_status_panel(content)

    def _build_parameters_panel(self, parent) -> None:
        panel = tb.Labelframe(parent, text="Forecast Parameters", padding=14)
        panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

        today = date.today()

        self._label(panel, "Forecast Month")
        self.forecast_month_combo = tb.Combobox(
            panel, values=list(MONTH_LOOKUP), state="readonly", width=32
        )
        self.forecast_month_combo.pack(fill="x", pady=(3, 7))
        self.forecast_month_combo.set(f"{today:%B}")

        self._label(panel, "Forecast Year")
        self.forecast_year_combo = tb.Combobox(
            panel,
            values=[str(year) for year in range(today.year - 1, today.year + 4)],
            state="readonly",
            width=32,
        )
        self.forecast_year_combo.pack(fill="x", pady=(3, 7))
        self.forecast_year_combo.set(str(today.year))

        for combo in (self.forecast_month_combo, self.forecast_year_combo):
            combo.bind("<<ComboboxSelected>>", self._on_period_changed)

        tb.Label(
            panel,
            textvariable=self.period_summary_var,
            bootstyle="info",
            font=("Segoe UI", 9, "bold"),
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        self._path_picker(
            panel, "Previous Forecast Folder", self.source_folder_var, self._browse_source_folder
        )
        self._path_picker(
            panel, "Forecast Output Folder", self.output_folder_var, self._browse_output_folder
        )
        self._path_picker(
            panel,
            "Tableau Forecast Workbook",
            self.tableau_workbook_var,
            self._browse_tableau_workbook,
        )

        self._label(panel, "Existing Files")
        self.overwrite_combo = tb.Combobox(
            panel,
            textvariable=self.overwrite_var,
            values=["skip", "replace", "cancel"],
            state="readonly",
        )
        self.overwrite_combo.pack(fill="x", pady=(5, 16))
        self.overwrite_combo.set("skip")

        tb.Separator(panel).pack(fill="x", pady=(2, 8))

        self.prepare_button = tb.Button(
            panel,
            text="1.  Create Forecast Workbooks",
            bootstyle="success",
            command=self._start_prepare,
        )
        self.prepare_button.pack(fill="x", pady=(0, 5), ipady=2)

        self._step_buttons.append(self.prepare_button)
        self._build_category_grid(panel)

    def _build_category_grid(self, panel) -> None:
        """Steps 2 and 3 per category, like the Actuals page."""
        grid = tb.Frame(panel)
        grid.pack(fill="x", pady=(8, 0))
        for col, weight in enumerate((0, 0, 1, 0, 1)):
            grid.columnconfigure(col, weight=weight)

        headers = ("Category", "2. Update from Tableau", "", "3. Validate", "")
        for col, text in enumerate(headers):
            if text:
                tb.Label(grid, text=text, font=("Segoe UI", 9, "bold")).grid(
                    row=0, column=col, columnspan=2 if col else 1, sticky="w", pady=(0, 4)
                )

        rows = [(name, len(routes_for([name]))) for name in FORECAST_CATEGORIES]
        rows.append(("All", len(routes_for())))
        row = 0
        for name, count in rows:
            row += 1
            if name == "All":
                tb.Separator(grid).grid(row=row, column=0, columnspan=5, sticky="ew", pady=4)
                row += 1
            label = "All categories" if name == "All" else name
            tb.Label(grid, text=f"{label} ({count})", font=("Segoe UI", 9, "bold" if name == "All" else "normal")).grid(
                row=row, column=0, sticky="w", padx=(0, 8), pady=2
            )
            categories = None if name == "All" else [name]
            update_var, validate_var = self._category_vars[name]
            for col, (text, style, command, variable) in enumerate(
                (
                    ("Update", "info-outline" if categories else "info",
                     lambda c=categories: self._start_population(c), update_var),
                    ("Validate", "primary-outline" if categories else "primary",
                     lambda c=categories: self._start_validation(c), validate_var),
                )
            ):
                button = tb.Button(grid, text=text, bootstyle=style, command=command, width=9)
                button.grid(row=row, column=1 + col * 2, sticky="w", pady=2)
                self._step_buttons.append(button)
                tb.Label(grid, textvariable=variable, bootstyle="secondary", wraplength=170, justify="left").grid(
                    row=row, column=2 + col * 2, sticky="w", padx=(6, 10), pady=2
                )

    def _build_status_panel(self, parent) -> None:
        panel = tb.Labelframe(parent, text="Process Status", padding=14)
        panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

        tb.Label(
            panel, textvariable=self.status_var, font=("Segoe UI", 20, "bold")
        ).pack(anchor="w", pady=(6, 4))

        tb.Label(
            panel,
            textvariable=self.status_description_var,
            bootstyle="secondary",
            justify="left",
            wraplength=350,
        ).pack(anchor="w", pady=(0, 10))

        tb.Progressbar(
            panel, variable=self.progress_var, maximum=100, bootstyle="info-striped"
        ).pack(fill="x", pady=(0, 22))

        stats = tb.Frame(panel)
        stats.pack(fill="x", pady=(0, 20))
        for col in range(6):
            stats.columnconfigure(col, weight=1)

        self._stat(stats, 0, self.files_created_var, "Workbooks")
        self._stat(stats, 1, self.worksheets_updated_var, "Sheets Updated")
        self._stat(stats, 2, self.validation_passed_var, "Passed")
        self._stat(stats, 3, self.validation_failed_var, "Failed")
        self._stat(stats, 4, self.action_required_var, "Action Required")
        self._stat(stats, 5, self.not_validated_var, "Not Validated")

        tb.Separator(panel).pack(fill="x", pady=(4, 18))

        self._label(panel, "Forecast files location")
        tb.Label(
            panel,
            textvariable=self.output_var,
            bootstyle="secondary",
            justify="left",
            wraplength=600,
        ).pack(anchor="w", pady=(5, 18))

        self._label(panel, "Validation report location")
        tb.Label(
            panel,
            textvariable=self.validation_output_var,
            bootstyle="secondary",
            justify="left",
            wraplength=600,
        ).pack(anchor="w", pady=(5, 18))

        actions = tb.Frame(panel)
        actions.pack(fill="x")

        self.open_output_button = tb.Button(
            actions,
            text="Open Output Folder",
            bootstyle="secondary-outline",
            command=self._open_output_folder,
        )
        self.open_output_button.pack(side="left", fill="x", expand=True, padx=(0, 6))

        self.open_validation_button = tb.Button(
            actions,
            text="Open Validation Report",
            bootstyle="primary-outline",
            command=self._open_validation_report,
            state="disabled",
        )
        self.open_validation_button.pack(side="left", fill="x", expand=True, padx=6)

        self.reset_button = tb.Button(
            actions,
            text="Reset",
            bootstyle="secondary-outline",
            command=self._reset_view,
        )
        self.reset_button.pack(side="right", fill="x", expand=True, padx=(6, 0))

    # =========================================================
    # SMALL UI HELPERS
    # =========================================================

    @staticmethod
    def _label(parent, text: str) -> None:
        tb.Label(parent, text=text, font=("Segoe UI", 10, "bold")).pack(anchor="w")

    def _path_picker(self, parent, label, variable, command) -> None:
        self._label(parent, label)
        row = tb.Frame(parent)
        row.pack(fill="x", pady=(3, 7))
        tb.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        tb.Button(row, text="Browse", bootstyle="secondary-outline", command=command).pack(
            side="right", padx=(8, 0)
        )

    @staticmethod
    def _stat(parent, column, variable, label) -> None:
        frame = tb.Frame(parent)
        frame.grid(row=0, column=column, sticky="nsew")
        tb.Label(frame, textvariable=variable, font=("Segoe UI", 18, "bold")).pack()
        tb.Label(frame, text=label, font=("Segoe UI", 8), bootstyle="secondary").pack(
            pady=(3, 0)
        )

    def _post(self, callback, *args) -> None:
        """Run ``callback`` on the UI thread. Safe to call from a worker thread.

        If the page was closed while the worker ran, the result is only logged.
        """
        try:
            self.after(0, callback, *args)
        except (RuntimeError, tk.TclError):
            log.warning("Forecast page closed before %s could be shown", callback.__name__)

    def _progress_reporter(self, start: int, end: int, verb: str):
        """A progress callback for the services that maps (done, total) to start..end %."""

        def report(done: int, total: int, label: str) -> None:
            percent = start + int((end - start) * done / total) if total else start
            text = f"{verb} {min(done + 1, total)} of {total}: {label}" if done < total else label
            self._post(self._show_progress, percent, text)

        return report

    def _show_progress(self, percent: int, text: str) -> None:
        self.progress_var.set(percent)
        self.status_description_var.set(text)

    # =========================================================
    # PERIOD
    # =========================================================

    def _period(self) -> dict:
        month = MONTH_LOOKUP[self.forecast_month_combo.get()]
        year = int(self.forecast_year_combo.get())
        return build_forecast_period(
            fiscal_year=fiscal_year_for(month, year),
            forecast_month=month,
            forecast_year=year,
        )

    def _on_period_changed(self, _event=None) -> None:
        """Refresh everything that depends on the month/year (no restart needed)."""
        period = self._period()
        self.period_summary_var.set(
            f"Selected Period: {period['display']} ({period['fiscal_year']})\n"
            f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
            f"Forecast Type: {period['forecast_type']}"
        )
        self._clear_results()
        self._show_ready_message(period)

    def _show_ready_message(self, period: dict) -> None:
        self.status_var.set("Ready")
        if period["forecast_month"] == 7:
            self.status_description_var.set(
                "July starts the fiscal year: create it with Initial Budget Setup. "
                "Steps 2 and 3 can still be run on existing July workbooks."
            )
        else:
            self.status_description_var.set(
                f"Ready for {period['display']}. Select the three locations, "
                "then run the steps in order."
            )

    def _clear_results(self) -> None:
        """Forget results and paths from a previous run or month."""
        self.progress_var.set(0)
        for var in (
            self.files_created_var,
            self.worksheets_updated_var,
            self.validation_passed_var,
            self.validation_failed_var,
            self.action_required_var,
            self.not_validated_var,
        ):
            var.set("0")
        self.output_var.set(self.output_folder_var.get().strip() or "Not created yet")
        self._report_path = None
        self.open_validation_button.configure(state="disabled")
        for update_var, validate_var in self._category_vars.values():
            update_var.set("—")
            validate_var.set("—")
        self._update_validation_output()
        self._load_saved_validation()

    def _load_saved_validation(self) -> None:
        """Show what was already validated this month, so people can pick up where they left off."""
        try:
            period = self._period()
            folder = get_forecast_validation_output_folder(
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                create=False,
            )
            saved = load_saved_results(folder, period["forecast_month"], period["forecast_year"], period["forecast_type"])
        except Exception:
            log.debug("No saved forecast validation results", exc_info=True)
            return
        if saved:
            self._show_validation_status(saved)
            report = report_path_for(folder, period["forecast_month"], period["forecast_year"])
            if report.is_file():
                self._report_path = report
                self.open_validation_button.configure(state="normal")

    @staticmethod
    def _summary_text(results: list[dict]) -> str:
        c = result_counts(results)
        parts = [f"{c['PASS']} pass"]
        for key, label in (("FAIL", "fail"), ("ACTION REQUIRED", "action")):
            if c[key]:
                parts.append(f"{c[key]} {label}")
        if c["NOT VALIDATED"] - c["NOT SET UP"]:
            parts.append(f"{c['NOT VALIDATED'] - c['NOT SET UP']} not validated")
        if c["NOT SET UP"]:
            parts.append(f"{c['NOT SET UP']} not set up")
        return " · ".join(parts)

    def _show_validation_status(self, by_category: dict) -> None:
        everything = []
        for name, entry in by_category.items():
            if name not in self._category_vars:
                continue
            results = entry.get("results", [])
            everything.extend(results)
            when = entry.get("validated_at") or ""
            when = f"{when[5:]}: " if len(when) >= 16 else ""  # "MM-DD HH:MM: "
            self._category_vars[name][1].set(when + self._summary_text(results))
        if len(by_category) > 1 or set(by_category) >= set(FORECAST_CATEGORIES):
            self._category_vars["All"][1].set(self._summary_text(everything) if everything else "—")

    def _show_update_status(self, result: dict) -> None:
        names = result.get("categories") or FORECAST_CATEGORIES
        for name in names:
            passed = sum(item.get("category") == name for item in result["passed"])
            failed = sum(item.get("category") == name for item in result["failed"])
            text = f"{passed} updated" + (f" · {failed} not updated" if failed else "")
            self._category_vars[name][0].set(text if passed or failed else "—")
        if set(names) >= set(FORECAST_CATEGORIES):
            self._category_vars["All"][0].set(
                f"{result['routes_passed']} updated"
                + (f" · {result['routes_failed']} not updated" if result["routes_failed"] else "")
            )

    def _update_validation_output(self) -> None:
        try:
            period = self._period()
            path = get_forecast_validation_output_folder(
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                create=False,
            )
            self.validation_output_var.set(str(path))
        except Exception as exc:
            self.validation_output_var.set(f"Not available: {exc}")

    # =========================================================
    # BROWSERS
    # =========================================================

    def _browse_source_folder(self) -> None:
        value = filedialog.askdirectory(
            parent=self.winfo_toplevel(), title="Select Previous Forecast Folder"
        )
        if value:
            self.source_folder_var.set(value)

    def _browse_output_folder(self) -> None:
        value = filedialog.askdirectory(
            parent=self.winfo_toplevel(), title="Select Forecast Output Folder"
        )
        if value:
            self.output_folder_var.set(value)
            self.output_var.set(value)

    def _browse_tableau_workbook(self) -> None:
        value = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="Select Tableau Forecast Workbook",
            filetypes=[("Tableau Workbook", "*.twb"), ("All Files", "*.*")],
        )
        if value:
            self.tableau_workbook_var.set(value)

    # =========================================================
    # INPUTS (read on the UI thread, passed to workers)
    # =========================================================

    def _error(self, title: str, message: str) -> None:
        messagebox.showerror(title, message, parent=self.winfo_toplevel())

    def _collect_inputs(self, *, source=False, output_exists=False, tableau=False) -> dict | None:
        """Read and check the form. Returns None (after showing why) if invalid."""
        try:
            period = self._period()
        except ValueError as exc:
            self._error("Forecast Period", str(exc))
            return None

        source_text = self.source_folder_var.get().strip()
        output_text = self.output_folder_var.get().strip()
        tableau_text = self.tableau_workbook_var.get().strip()

        # Path("") is ".", which "exists": always test the text first.
        if source and (not source_text or not Path(source_text).is_dir()):
            self._error("Previous Forecast Folder", "Select the folder that holds last month's forecast workbooks.")
            return None

        if not output_text:
            self._error("Forecast Output Folder", "Select a forecast output folder.")
            return None
        if output_exists and not Path(output_text).is_dir():
            self._error(
                "Forecast Output Folder",
                f"This folder does not exist:\n{output_text}\n\nRun step 1 first, or select the folder it created.",
            )
            return None
        if source and Path(output_text).resolve() == Path(source_text).resolve():
            if not messagebox.askyesno(
                "Forecast Output Folder",
                "The output folder is the same as the previous forecast folder. "
                "New workbooks will be created next to last month's. Continue?",
                parent=self.winfo_toplevel(),
            ):
                return None

        if tableau and (
            not tableau_text
            or not Path(tableau_text).is_file()
            or Path(tableau_text).suffix.lower() != ".twb"
        ):
            self._error("Tableau Forecast Workbook", "Select the FY Forecast Tableau .twb workbook.")
            return None

        inputs = {
            "period": period,
            "source_folder": source_text,
            "output_folder": output_text,
            "tableau_workbook": tableau_text,
            "overwrite_mode": self.overwrite_var.get(),
        }
        self._remember_paths(inputs)
        return inputs

    @staticmethod
    def _remembered_paths() -> dict:
        keys = ("forecast_source_folder", "forecast_output_folder", "forecast_tableau_workbook")
        try:
            return {key: str(settings.get_path(key) or "") for key in keys}
        except Exception:
            log.warning("Could not read the remembered forecast paths", exc_info=True)
            return {key: "" for key in keys}

    @staticmethod
    def _remember_paths(inputs: dict) -> None:
        """Save the paths that were just used, for next time (best effort)."""
        values = {
            "forecast_source_folder": inputs["source_folder"],
            "forecast_output_folder": inputs["output_folder"],
            "forecast_tableau_workbook": inputs["tableau_workbook"],
        }
        values = {key: value for key, value in values.items() if value}
        try:
            if any(settings.get(key) != value for key, value in values.items()):
                settings.save_user_settings(values)
        except Exception:
            log.warning("Could not remember the forecast paths", exc_info=True)

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        for widget in (*self._step_buttons, self.reset_button):
            widget.configure(state=state)
        # The period cannot change under a running step.
        combo_state = "disabled" if busy else "readonly"
        for combo in (self.forecast_month_combo, self.forecast_year_combo, self.overwrite_combo):
            combo.configure(state=combo_state)

    def _start_worker(self, target, inputs: dict) -> None:
        self._set_busy(True)
        threading.Thread(target=target, args=(inputs,), daemon=True).start()

    # =========================================================
    # STEP 1 - CREATE THE MONTHLY WORKBOOKS
    # =========================================================

    def _start_prepare(self) -> None:
        inputs = self._collect_inputs(source=True, tableau=True)
        if inputs is None:
            return
        period = inputs["period"]

        if not messagebox.askyesno(
            "Create Forecast Workbooks",
            (
                f"Create the {period['display']} forecast workbooks?\n\n"
                f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"Forecast Type: {period['forecast_type']}\n"
                f"Existing files: {inputs['overwrite_mode']}\n\n"
                "The same period is used for Excel, the consolidated Tableau "
                "workbook and the individual Tableau validation."
            ),
            parent=self.winfo_toplevel(),
        ):
            return

        self._clear_results()
        self.status_var.set("Creating Forecast Workbooks")
        self.status_description_var.set(
            "Checking the consolidated Tableau workbook and creating the new monthly workbooks."
        )
        self.progress_var.set(10)
        self._start_worker(self._run_prepare, inputs)

    def _run_prepare(self, inputs: dict) -> None:
        period = inputs["period"]
        try:
            result = prepare_monthly_forecast(
                source_folder=inputs["source_folder"],
                output_folder=inputs["output_folder"],
                tableau_workbook_path=inputs["tableau_workbook"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                overwrite_mode=inputs["overwrite_mode"],
            )
            self._post(self._prepare_complete, result)
        except Exception as exc:
            log.exception("Forecast step 1 failed")
            self._post(self._process_failed, exc)

    def _prepare_complete(self, result: dict) -> None:
        self._set_busy(False)
        period = result["period"]
        excel = result["excel"]
        failures = excel.get("failures") or []
        older = excel.get("older_sources") or []

        get_attention(self, success=not failures)
        self.files_created_var.set(str(excel.get("files_created", 0)))
        self.output_var.set(result["output_folder"])
        self.progress_var.set(33)

        lines = [
            f"Created: {excel.get('files_created', 0)}",
            f"Skipped (already existed): {excel.get('files_skipped', 0)}",
            f"Failed: {len(failures)}",
        ]
        if failures:
            lines += ["", "These workbooks could not be created:"]
            lines += [f"  • {item['file']}: {item['error']}" for item in failures[:8]]
        if older:
            lines += ["", f"No {excel.get('expected_source_month')} file was found for these; an older month was used:"]
            lines += [f"  • {item['family']} ({item['source_month']})" for item in older[:8]]
        if excel.get("files_skipped") and result["excel"].get("files_created", 0) == 0:
            lines += ["", "Every workbook already existed and was kept. Choose Existing Files = replace to recreate them."]

        if failures:
            self.status_var.set("Workbooks Created with Exceptions")
            self.status_description_var.set(f"{len(failures)} workbook(s) could not be created. See the message.")
        else:
            self.status_var.set("Forecast Workbooks Created")
            self.status_description_var.set(
                f"{period['display']} workbooks are ready. Next: 2. Update Workbooks from Tableau."
            )

        lines += ["", "Next: click 2. Update Workbooks from Tableau. The app opens Tableau itself."]
        (messagebox.showwarning if failures or older else messagebox.showinfo)(
            "Forecast Workbooks", "\n".join(lines), parent=self.winfo_toplevel()
        )

    # =========================================================
    # STEP 2 - CONSOLIDATED TABLEAU -> EXCEL
    # =========================================================

    # ---------------------------------------------------------
    # Quick file check before Tableau (steps 2 and 3)
    # ---------------------------------------------------------

    def _start_check(self, kind: str, inputs: dict) -> None:
        self.status_var.set("Checking files")
        self.status_description_var.set("Checking the forecast and Tableau files before opening Tableau...")
        self._start_worker(self._run_check, {**inputs, "kind": kind})

    def _run_check(self, inputs: dict) -> None:
        try:
            routes = routes_for(inputs["categories"])
            if inputs["kind"] == "update":
                problems = check_before_update(
                    routes, inputs["output_folder"], inputs["tableau_workbook"], inputs["period"]
                )
            else:
                problems = check_before_validation(routes, inputs["output_folder"], inputs["period"])
        except Exception as exc:
            log.exception("Forecast file check failed")
            problems = [{"workbook": "", "worksheet": "", "problem": f"The file check could not finish ({friendly(exc)}). "
                                                                      "The run can still continue."}]
        self._post(self._check_done, inputs, problems)

    def _check_done(self, inputs: dict, problems: list[dict]) -> None:
        self._set_busy(False)
        if problems:
            shown = [
                f"• {p['workbook'].replace('_Aug Forecast', '')} / {p['worksheet']}: {p['problem']}"
                if p.get("workbook") else f"• {p['problem']}"
                for p in problems[:12]
            ]
            if len(problems) > 12:
                shown.append(f"…and {len(problems) - 12} more.")
            step = "updated" if inputs["kind"] == "update" else "validated"
            if not messagebox.askyesno(
                "Check before Tableau",
                f"{len(problems)} problem(s) found before opening Tableau:\n\n"
                + "\n".join(shown)
                + f"\n\nContinue anyway? These sheets will not be {step}; the others will.",
                icon="warning",
                parent=self.winfo_toplevel(),
            ):
                self.status_var.set("Stopped before Tableau")
                self.status_description_var.set("Fix the problems listed, then run again.")
                return
        if not confirm_close_tableau(self):
            self._show_ready_message(inputs["period"])
            return
        if inputs["kind"] == "update":
            self._begin_population(inputs)
        else:
            self._begin_validation(inputs)

    @staticmethod
    def _describe(categories) -> str:
        return "all categories" if not categories else ", ".join(categories)

    def _start_population(self, categories=None) -> None:
        inputs = self._collect_inputs(output_exists=True, tableau=True)
        if inputs is None:
            return
        inputs["categories"] = categories
        period = inputs["period"]

        if not messagebox.askyesno(
            "Update Workbooks from Tableau",
            (
                f"Update the {period['display']} workbooks for "
                f"{self._describe(categories)} from Tableau?\n\n"
                f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"Forecast Type: {period['forecast_type']}\n\n"
                "The app opens a temporary copy of the consolidated workbook, "
                "copies each crosstab and closes it again. Please don't use the "
                "mouse or keyboard until it finishes."
            ),
            parent=self.winfo_toplevel(),
        ):
            return
        self._start_check("update", inputs)

    def _begin_population(self, inputs: dict) -> None:
        categories = inputs["categories"]
        self.status_var.set(f"Updating {self._describe(categories)}")
        self.status_description_var.set("Opening the consolidated Tableau workbook...")
        for name in categories or (*FORECAST_CATEGORIES, "All"):
            self._category_vars[name][0].set("Running…")
        self.progress_var.set(35)
        self._start_worker(self._run_population, inputs)

    def _run_population(self, inputs: dict) -> None:
        period = inputs["period"]
        try:
            result = populate_monthly_forecast_from_tableau(
                forecast_folder=inputs["output_folder"],
                tableau_workbook_path=inputs["tableau_workbook"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                categories=inputs["categories"],
                progress=self._progress_reporter(35, 66, "Copying crosstab"),
            )
            self._post(self._population_complete, result)
        except Exception as exc:
            log.exception("Forecast step 2 failed")
            self._post(self._process_failed, exc)

    def _population_complete(self, result: dict) -> None:
        self._set_busy(False)
        failed = result["failed"]
        get_attention(self, success=not failed)
        self.worksheets_updated_var.set(str(result["worksheets_updated"]))
        self.files_created_var.set(str(result["workbooks_updated"]))
        self.progress_var.set(66)
        self._show_update_status(result)

        if failed:
            self.status_var.set("Updated with Exceptions")
            self.status_description_var.set(
                f"{result['routes_passed']} sheets updated, {result['routes_failed']} not updated. "
                "Sheets that were not updated will show in the validation report."
            )
            lines = [f"• {item['workbook']} / {item['worksheet']}: {item['error']}" for item in failed[:10]]
            if len(failed) > 10:
                lines.append(f"…and {len(failed) - 10} more (see the log).")
            message = (
                f"{result['routes_passed']} sheets were updated. These were not:\n\n"
                + "\n".join(lines)
            )
        else:
            self.status_var.set("Forecast Workbooks Updated")
            self.status_description_var.set(
                f"All {result['routes_passed']} sheets were updated. Next: Validate."
            )
            message = (
                f"{result['workbooks_updated']} workbooks and {result['worksheets_updated']} "
                "sheets were updated.\n\nNext: click Validate for the same category."
            )
        if result.get("close_warning"):
            message += f"\n\nTableau could not be closed automatically: {result['close_warning']}"

        (messagebox.showwarning if failed else messagebox.showinfo)(
            "Update Workbooks from Tableau", message, parent=self.winfo_toplevel()
        )

    # =========================================================
    # STEP 3 - INDIVIDUAL TABLEAU VALIDATION
    # =========================================================

    def _start_validation(self, categories=None) -> None:
        inputs = self._collect_inputs(output_exists=True)
        if inputs is None:
            return
        inputs["categories"] = categories
        period = inputs["period"]

        if not messagebox.askyesno(
            "Validate Final Forecast",
            (
                f"Validate the {period['display']} forecast for "
                f"{self._describe(categories)}?\n\n"
                "Each individual Tableau workbook is opened as a temporary copy "
                f"set to Forecast Type {period['forecast_type']}. A route that "
                "can't be checked is marked in the report and the run continues."
            ),
            parent=self.winfo_toplevel(),
        ):
            return
        self._start_check("validate", inputs)

    def _begin_validation(self, inputs: dict) -> None:
        categories = inputs["categories"]
        self.status_var.set(f"Validating {self._describe(categories)}")
        for name in categories or (*FORECAST_CATEGORIES, "All"):
            self._category_vars[name][1].set("Running…")
        self.status_description_var.set("Comparing Excel Grand Totals to the individual Tableau workbooks.")
        self.progress_var.set(67)
        self._start_worker(self._run_validation, inputs)

    def _run_validation(self, inputs: dict) -> None:
        period = inputs["period"]
        try:
            result = validate_monthly_forecast(
                forecast_folder=inputs["output_folder"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                tolerance=VALIDATION_TOLERANCE,
                categories=inputs["categories"],
                progress=self._progress_reporter(67, 99, "Validating"),
            )
            self._post(self._validation_complete, result)
        except Exception as exc:
            log.exception("Forecast step 3 failed")
            self._post(self._process_failed, exc)

    def _validation_complete(self, result: dict) -> None:
        self._set_busy(False)
        passed = int(result.get("passed", 0))
        failed = int(result.get("failed", 0))
        action = int(result.get("action_required", 0))
        not_set_up = int(result.get("not_configured", 0))
        # "Not validated" here means a suspect line (e.g. wrong month);
        # routes not set up for validation yet are reported separately.
        not_validated = int(result.get("not_validated", 0)) - not_set_up
        clean = not (failed or action or not_validated)
        self._show_validation_status(
            {name: {"validated_at": "just now", "results": [r for r in result.get("results", []) if r.get("category") == name]}
             for name in result.get("categories") or FORECAST_CATEGORIES}
        )

        get_attention(self, success=clean)
        self.validation_passed_var.set(str(passed))
        self.validation_failed_var.set(str(failed))
        self.action_required_var.set(str(action))
        self.not_validated_var.set(str(not_validated))
        self.progress_var.set(100)

        report_path = result.get("validation_report")
        if report_path:
            self._report_path = Path(report_path)
            self.validation_output_var.set(report_path)
            self.open_validation_button.configure(state="normal")

        ran = result.get("categories") or list(FORECAST_CATEGORIES)
        scope = "Forecast" if set(ran) >= set(FORECAST_CATEGORIES) else ", ".join(ran)
        if clean:
            self.status_var.set(f"{scope} Reconciled")
            self.status_description_var.set(
                f"All {passed} validated routes reconciled within ${VALIDATION_TOLERANCE:,.2f}"
                + (f"; {not_set_up} routes are not set up for validation yet." if not_set_up else ".")
            )
        else:
            self.status_var.set(f"{scope}: Review Required")
            self.status_description_var.set(
                f"{passed} passed, {failed} failed, {action} action required, {not_validated} not validated."
            )

        message = (
            "Final forecast validation complete.\n\n"
            f"Passed: {passed}\n"
            f"Failed: {failed}\n"
            f"Action Required: {action}\n"
            f"Not Validated: {not_validated}\n"
            f"Not set up for validation yet: {not_set_up}"
        )
        problems = [
            item for item in result.get("results", [])
            if item.get("status") in ("FAIL", "ACTION REQUIRED", "NOT VALIDATED") and not item.get("not_configured")
        ]
        if problems:
            message += "\n\nFirst items to review:\n" + "\n".join(
                f"• {item.get('forecast_workbook')} / {item.get('worksheet')}: {item.get('note')}"
                for item in problems[:5]
            )
        expected_name = f"Forecast_Validation_{result['period']['forecast_start']:%b_%Y}"
        if report_path and Path(report_path).stem != expected_name:
            message += (
                "\n\nThe previous report was open in Excel, so this one was saved as:\n"
                f"{Path(report_path).name}"
            )
        message += "\n\nOpen the validation report now?"

        if messagebox.askyesno(
            "Forecast Validation Complete", message, icon="info" if clean else "warning",
            parent=self.winfo_toplevel(),
        ):
            self._open_validation_report()

    # =========================================================
    # COMMON
    # =========================================================

    def _process_failed(self, error: Exception) -> None:
        self._set_busy(False)
        for update_var, validate_var in self._category_vars.values():
            for variable in (update_var, validate_var):
                if variable.get() == "Running…":
                    variable.set("Stopped – see message")
        get_attention(self, success=False)
        self.progress_var.set(0)
        self.status_var.set("Process Failed")
        self.status_description_var.set(explain(error))
        self._error("Monthly Forecast Automation", friendly(error))

    def _reset_view(self) -> None:
        if self._busy:
            return
        self._clear_results()
        self._show_ready_message(self._period())

    def _open_validation_report(self) -> None:
        path = self._report_path
        if path is None or not path.is_file():
            self._error("Validation Report", "The validation report has not been created yet.")
            return
        os.startfile(path)

    def _open_output_folder(self) -> None:
        folder = self.output_folder_var.get().strip()
        if not folder:
            self._error("Forecast Output Folder", "Select a forecast output folder first.")
            return
        path = Path(folder)
        if not path.is_dir():
            self._error("Folder Not Found", f"{path}\n\nIt is created by step 1.")
            return
        os.startfile(path)
