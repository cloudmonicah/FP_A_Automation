from __future__ import annotations
from shared.friendly_errors import explain, friendly

import logging

import os
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
from shared.forecast_output_paths import (
    get_forecast_validation_output_folder,
)

import ttkbootstrap as tb

from automations.forecast.service import (
    build_forecast_period,
    populate_monthly_forecast_from_tableau,
    prepare_monthly_forecast,
    validate_monthly_forecast,
)

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


from ui.common import confirm_close_tableau, get_attention

log = logging.getLogger(__name__)

class ForecastRefreshView(tb.Frame):
    """
    Monthly Forecast GUI.

    One selected forecast period controls:
      - Excel Forecast Type
      - consolidated Tableau Forecast Start / Forecast Type
      - individual Tableau validation period
    """

    def __init__(self, parent) -> None:
        super().__init__(parent)

        self.source_folder_var = tb.StringVar()
        self.output_folder_var = tb.StringVar()
        self.tableau_workbook_var = tb.StringVar()
        self.overwrite_var = tb.StringVar(value="skip")

        self.status_var = tb.StringVar(value="Ready")
        self.status_description_var = tb.StringVar(
            value="Select the forecast period and the three required locations."
        )
        self.period_summary_var = tb.StringVar()
        self.progress_var = tb.IntVar(value=0)

        self.files_created_var = tb.StringVar(value="0")
        self.worksheets_updated_var = tb.StringVar(value="0")
        self.validation_passed_var = tb.StringVar(value="0")
        self.validation_failed_var = tb.StringVar(value="0")
        self.action_required_var = tb.StringVar(value="0")
        self.output_var = tb.StringVar(value="Not created yet")
        self.validation_output_var = tb.StringVar()

        self._build_view()
        self._update_period_summary()

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

        panel = tb.Labelframe(
            parent,
            text="Forecast Parameters",
            padding=14,
        )

        panel.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(0, 8),
        )

        self._label(panel, "Fiscal Year")
        self.fiscal_year_combo = tb.Combobox(
            panel,
            values=["FY27", "FY28", "FY29", "FY30"],
            state="readonly",
            width=42,
        )
        self.fiscal_year_combo.pack(fill="x", pady=(3, 7))
        self.fiscal_year_combo.set("FY27")

        self._label(panel, "Forecast Month")
        self.forecast_month_combo = tb.Combobox(
            panel,
            values=list(MONTH_LOOKUP),
            state="readonly",
            width=32,
        )
        self.forecast_month_combo.pack(fill="x", pady=(3, 7))
        self.forecast_month_combo.set("October")

        self._label(panel, "Forecast Year")
        self.forecast_year_combo = tb.Combobox(
            panel,
            values=["2026", "2027", "2028", "2029", "2030"],
            state="readonly",
            width=32,
        )
        self.forecast_year_combo.pack(fill="x", pady=(3, 7))
        self.forecast_year_combo.set("2026")

        for combo in (
            self.fiscal_year_combo,
            self.forecast_month_combo,
            self.forecast_year_combo,
        ):
            combo.bind(
                "<<ComboboxSelected>>",
                self._update_period_summary,
            )

        tb.Label(
            panel,
            textvariable=self.period_summary_var,
            bootstyle="info",
            font=("Segoe UI", 9, "bold"),
            justify="left",
        ).pack(anchor="w", pady=(0, 10))

        self._path_picker(
            panel,
            "Previous Forecast Folder",
            self.source_folder_var,
            self._browse_source_folder,
        )
        self._path_picker(
            panel,
            "Forecast Output Folder",
            self.output_folder_var,
            self._browse_output_folder,
        )
        self._path_picker(
            panel,
            "Tableau Forecast Workbook",
            self.tableau_workbook_var,
            self._browse_tableau_workbook,
        )

        self._label(panel, "Existing Files")
        overwrite = tb.Combobox(
            panel,
            textvariable=self.overwrite_var,
            values=["skip", "replace", "cancel"],
            state="readonly",
        )
        overwrite.pack(fill="x", pady=(5, 16))
        overwrite.set("skip")

        tb.Separator(panel).pack(fill="x", pady=(2, 8))

        self.prepare_button = tb.Button(
            panel,
            text="1.  Generate Forecast & Open Tableau",
            bootstyle="success",
            command=self._start_prepare,
        )
        self.prepare_button.pack(fill="x", pady=(0, 5), ipady=2)

        self.populate_button = tb.Button(
            panel,
            text="2.  Update Workbooks from Tableau",
            bootstyle="info",
            command=self._start_population,
            state="normal",
        )
        self.populate_button.pack(fill="x", pady=(0, 5), ipady=2)

        self.validate_button = tb.Button(
            panel,
            text="3.  Validate Final Forecast",
            bootstyle="primary",
            command=self._start_validation,
            state="normal",
        )
        self.validate_button.pack(fill="x", ipady=2)

    def _build_status_panel(self, parent) -> None:
        panel = tb.Labelframe(
            parent,
            text="Process Status",
            padding=14,
        )

        panel.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(8, 0),
        )

        tb.Label(
            panel,
            textvariable=self.status_var,
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="w", pady=(6, 4))

        tb.Label(
            panel,
            textvariable=self.status_description_var,
            bootstyle="secondary",
            justify="left",
            wraplength=350,
        ).pack(anchor="w", pady=(0, 10))

        tb.Progressbar(
            panel,
            variable=self.progress_var,
            maximum=100,
            bootstyle="info-striped",
        ).pack(fill="x", pady=(0, 22))

        stats = tb.Frame(panel)
        stats.pack(fill="x", pady=(0, 20))
        for col in range(5):
            stats.columnconfigure(col, weight=1)

        self._stat(stats, 0, self.files_created_var, "Workbooks")
        self._stat(stats, 1, self.worksheets_updated_var, "Sheets Updated")
        self._stat(stats, 2, self.validation_passed_var, "Validated")
        self._stat(stats, 3, self.validation_failed_var, "Failed")
        self._stat(stats, 4, self.action_required_var, "Action Required")

        tb.Separator(panel).pack(fill="x", pady=(4, 18))

        self._label(panel, "Forecast files location")
        tb.Label(
            panel,
            textvariable=self.output_var,
            bootstyle="secondary",
            justify="left",
            wraplength=600,
        ).pack(anchor="w", pady=(5, 18))

        self._label(
            panel,
            "Validation report location",
        )

        tb.Label(
            panel,
            textvariable=self.validation_output_var,
            bootstyle="secondary",
            justify="left",
            wraplength=600,
        ).pack(
            anchor="w",
            pady=(5, 18),
        )

        actions = tb.Frame(panel)
        actions.pack(fill="x")

        self.open_output_button = tb.Button(
            actions,
            text="Open Output Folder",
            bootstyle="secondary-outline",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_output_button.pack(
            side="left",
            fill="x",
            expand=True,
            padx=(0, 6),
        )

        self.open_validation_button = tb.Button(
            actions,
            text="Open Validation Report",
            bootstyle="primary-outline",
            command=self._open_validation_report,
            state="disabled",
        )

        self.open_validation_button.pack(
            side="left",
            fill="x",
            expand=True,
            padx=6,
        )

        tb.Button(
            actions,
            text="Reset",
            bootstyle="secondary-outline",
            command=self._reset_view,
        ).pack(side="right", fill="x", expand=True, padx=(6, 0))

    # =========================================================
    # SMALL UI HELPERS
    # =========================================================

    @staticmethod
    def _label(parent, text: str) -> None:
        tb.Label(
            parent,
            text=text,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

    def _path_picker(self, parent, label, variable, command) -> None:
        self._label(parent, label)
        row = tb.Frame(parent)
        row.pack(fill="x", pady=(3, 7))
        tb.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        tb.Button(
            row,
            text="Browse",
            bootstyle="secondary-outline",
            command=command,
        ).pack(side="right", padx=(8, 0))

    @staticmethod
    def _stat(parent, column, variable, label) -> None:
        frame = tb.Frame(parent)
        frame.grid(row=0, column=column, sticky="nsew")
        tb.Label(
            frame,
            textvariable=variable,
            font=("Segoe UI", 18, "bold"),
        ).pack()
        tb.Label(
            frame,
            text=label,
            font=("Segoe UI", 8),
            bootstyle="secondary",
        ).pack(pady=(3, 0))

    def _update_validation_output(self) -> None:
        period = self._period()

        path = get_forecast_validation_output_folder(
            fiscal_year=period["fiscal_year"],
            forecast_month=period["forecast_month"],
            forecast_year=period["forecast_year"],
            create=False,
        )

        self.validation_output_var.set(str(path))

    def _open_validation_report(self) -> None:

        report_path = self.validation_output_var.get().strip()

        if not report_path:
            return

        path = Path(report_path)

        if not path.is_file():
            messagebox.showerror(
                "Validation Report",
                "The validation report has not been created yet.",
                parent=self.winfo_toplevel(),
            )
            return

        os.startfile(path)

    # =========================================================
    # PERIOD
    # =========================================================

    def _period(self) -> dict:
        return build_forecast_period(
            fiscal_year=self.fiscal_year_combo.get(),
            forecast_month=MONTH_LOOKUP[self.forecast_month_combo.get()],
            forecast_year=int(self.forecast_year_combo.get()),
        )

    @staticmethod
    def _fiscal_year_for(month: int, year: int) -> str:
        """July-June fiscal year: July 2026 .. June 2027 -> FY27."""
        return f"FY{(year + 1 if month >= 7 else year) % 100:02d}"

    def _update_period_summary(self, _event=None) -> None:
        # The fiscal year always follows the month and year, so the two can
        # never disagree (e.g. FY28 with October 2026).
        self.fiscal_year_combo.configure(state="normal")
        self.fiscal_year_combo.set(
            self._fiscal_year_for(
                MONTH_LOOKUP[self.forecast_month_combo.get()],
                int(self.forecast_year_combo.get()),
            )
        )
        self.fiscal_year_combo.configure(state="disabled")
        period = self._period()
        self.period_summary_var.set(
            f"Selected Period: {period['display']}\n"
            f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
            f"Forecast Type: {period['forecast_type']}"
        )
        self._update_validation_output()

    # =========================================================
    # BROWSERS
    # =========================================================

    def _browse_source_folder(self) -> None:
        value = filedialog.askdirectory(
            parent=self.winfo_toplevel(),
            title="Select Previous Forecast Folder",
        )
        if value:
            self.source_folder_var.set(value)

    def _browse_output_folder(self) -> None:
        value = filedialog.askdirectory(
            parent=self.winfo_toplevel(),
            title="Select Forecast Output Folder",
        )
        if value:
            self.output_folder_var.set(value)

    def _browse_tableau_workbook(self) -> None:
        value = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="Select Tableau Forecast Workbook",
            filetypes=[
                ("Tableau Workbook", "*.twb"),
                ("All Files", "*.*"),
            ],
        )
        if value:
            self.tableau_workbook_var.set(value)

    # =========================================================
    # INPUT VALIDATION
    # =========================================================

    def _validate_inputs(self) -> bool:
        source_text = self.source_folder_var.get().strip()
        output_text = self.output_folder_var.get().strip()
        tableau_text = self.tableau_workbook_var.get().strip()
        # Path("") is ".", which always "exists" - check the text first.
        source = Path(source_text)
        output = Path(output_text)
        tableau = Path(tableau_text)

        if not source_text or not source.is_dir():
            messagebox.showerror(
                "Previous Forecast Folder",
                "Select a valid previous forecast folder.",
                parent=self.winfo_toplevel(),
            )
            return False

        if not output_text:
            messagebox.showerror(
                "Forecast Output Folder",
                "Select a forecast output folder.",
                parent=self.winfo_toplevel(),
            )
            return False

        if not tableau_text or not tableau.is_file() or tableau.suffix.lower() != ".twb":
            messagebox.showerror(
                "Tableau Forecast Workbook",
                "Select the FY Forecast Tableau .twb workbook.",
                parent=self.winfo_toplevel(),
            )
            return False

        if output.resolve() == source.resolve() and not messagebox.askyesno(
            "Forecast Output Folder",
            "The output folder is the same as the Previous Forecast Folder, so the "
            "new workbooks will be created next to last month's. Continue?",
            parent=self.winfo_toplevel(),
        ):
            return False

        output.mkdir(parents=True, exist_ok=True)
        return True

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.prepare_button.configure(state=state)
        if busy:
            self.populate_button.configure(state="disabled")
            self.validate_button.configure(state="disabled")

    # =========================================================
    # STEP 1 - CREATE + PREPARE TABLEAU
    # =========================================================

    def _start_prepare(self) -> None:
        if not self._validate_inputs():
            return

        period = self._period()
        confirmed = messagebox.askyesno(
            "Run Monthly Forecast",
            (
                f"Run the forecast for {period['display']}?\n\n"
                f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"Forecast Type: {period['forecast_type']}\n\n"
                "These same period values will be used for Excel, "
                "the consolidated Tableau workbook, and individual "
                "Tableau validation."
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return
        if not confirm_close_tableau(self):
            return

        # Read everything here, on the UI thread. The worker thread must not
        # touch Tk (that is what caused the "Tcl_AsyncDelete" crash).
        inputs = {
            "period": period,
            "source_folder": self.source_folder_var.get().strip(),
            "output_folder": self.output_folder_var.get().strip(),
            "tableau_workbook": self.tableau_workbook_var.get().strip(),
            "overwrite_mode": self.overwrite_var.get(),
        }

        self._set_busy(True)
        self.status_var.set("Preparing Monthly Forecast")
        self.status_description_var.set(
            "Creating the new monthly workbooks and setting the "
            "consolidated Tableau forecast period."
        )
        self.progress_var.set(15)

        threading.Thread(
            target=self._run_prepare,
            args=(inputs,),
            daemon=True,
        ).start()

    def _run_prepare(self, inputs: dict) -> None:
        try:
            period = inputs["period"]
            result = prepare_monthly_forecast(
                source_folder=inputs["source_folder"],
                output_folder=inputs["output_folder"],
                tableau_workbook_path=inputs["tableau_workbook"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                overwrite_mode=inputs["overwrite_mode"],
                open_tableau=True,
            )
            self.after(0, lambda r=result: self._prepare_complete(r))
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, lambda e=exc: self._process_failed(e))

    def _prepare_complete(self, result: dict) -> None:
        get_attention(self)
        period = result["period"]
        excel = result["excel"]

        self.files_created_var.set(str(excel.get("files_created", 0)))
        self.output_var.set(result["output_folder"])
        self.open_output_button.configure(state="normal")

        self.progress_var.set(40)
        self.status_var.set("Waiting for Tableau")
        self.status_description_var.set(
            f"Tableau is open for {period['display']}. "
            f"Forecast Start = {period['forecast_start']:%m/%d/%Y}; "
            f"Forecast Type = {period['forecast_type']}. "
            "Allow Tableau to finish loading, then continue."
        )

        self.prepare_button.configure(state="normal")
        self.populate_button.configure(state="normal")
        self.validate_button.configure(state="normal")

        failures = excel.get("failures") or []
        skipped = excel.get("skipped_files") or []
        older = excel.get("older_sources") or []

        lines = [
            f"{period['display']} forecast workbooks:",
            f"  Created: {excel.get('files_created', 0)}",
            f"  Kept (already existed): {len(skipped)}",
            f"  Failed: {len(failures)}",
        ]
        if failures:
            lines += ["", "These could not be created:"]
            lines += [f"  - {item['file']}: {item['error']}" for item in failures[:8]]
            if len(failures) > 8:
                lines.append(f"  ...and {len(failures) - 8} more (see the log).")
        if skipped:
            lines += [
                "",
                "Existing files were kept as they are (Existing Files = \"skip\"). "
                "Choose \"replace\" to recreate them.",
            ]
        if older:
            lines += ["", f"No {excel.get('expected_source_month')} file was found for these, so an older month was used:"]
            lines += [f"  - {item['family']} ({item['source_month']})" for item in older[:8]]
        lines += [
            "",
            f"Forecast Start: {period['forecast_start']:%m/%d/%Y}",
            f"Forecast Type: {period['forecast_type']}",
            "",
            "Tableau is opening. Wait for it to finish loading and refreshing, then click:",
            "2.  Update Workbooks from Tableau",
        ]
        (messagebox.showwarning if failures or older else messagebox.showinfo)(
            "Forecast Workbooks Created",
            "\n".join(lines),
            parent=self.winfo_toplevel(),
        )

    # =========================================================
    # STEP 2 - CONSOLIDATED TABLEAU -> EXCEL
    # =========================================================

    def _start_population(self) -> None:
        if not self._validate_inputs():
            return

        period = self._period()
        confirmed = messagebox.askyesno(
            "Update Workbooks from Tableau",
            (
                "Confirm the consolidated Tableau workbook is ready.\n\n"
                f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"Forecast Type: {period['forecast_type']}\n\n"
                "Continue with crosstab extraction and Excel population?"
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return

        # Read everything on the UI thread; the worker must not touch Tk.
        inputs = {
            "period": period,
            "output_folder": self.output_folder_var.get().strip(),
            "tableau_workbook": self.tableau_workbook_var.get().strip(),
        }

        self._set_busy(True)
        self.status_var.set("Updating from Tableau")
        self.status_description_var.set(
            "Copying each mapped Tableau tab and pasting it into its "
            "forecast sheet."
        )
        self.progress_var.set(50)

        threading.Thread(
            target=self._run_population,
            args=(inputs,),
            daemon=True,
        ).start()

    def _run_population(self, inputs: dict) -> None:
        try:
            period = inputs["period"]
            result = populate_monthly_forecast_from_tableau(
                forecast_folder=inputs["output_folder"],
                tableau_workbook_path=inputs["tableau_workbook"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
            )
            self.after(0, lambda r=result: self._population_complete(r))
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, lambda e=exc: self._process_failed(e))

    def _population_complete(self, result: dict) -> None:
        get_attention(self)
        failed = result["failed"]
        self.worksheets_updated_var.set(str(result["worksheets_updated"]))
        self.files_created_var.set(str(result["workbooks_updated"]))
        self.progress_var.set(75)

        if failed:
            self.status_var.set("Population Completed with Exceptions")
            self.status_description_var.set(
                f"{result['routes_passed']} routes passed and "
                f"{result['routes_failed']} had exceptions. "
                "Continue to final validation; exceptions will be included "
                "in the validation results."
            )
        else:
            self.status_var.set("Forecast Workbooks Updated")
            self.status_description_var.set(
                f"All {result['routes_passed']} mapped Tableau routes "
                "were written to the forecast workbooks. "
                "Final individual Tableau validation is ready."
            )

        self.prepare_button.configure(state="normal")
        self.populate_button.configure(state="normal")
        self.validate_button.configure(state="normal")

        if failed:
            lines = [
                f"{item['workbook']} / {item['worksheet']}: {item['error']}"
                for item in failed[:10]
            ]
            messagebox.showwarning(
                "Forecast Population Completed with Exceptions",
                "\n".join(lines),
                parent=self.winfo_toplevel(),
            )
        else:
            messagebox.showinfo(
                "Forecast Population Complete",
                (
                    f"{result['workbooks_updated']} workbooks updated.\n"
                    f"{result['worksheets_updated']} worksheets updated.\n"
                    f"Months pasted from Tableau: {result.get('months_pasted') or 'none'}\n"
                    "Formulas in the workbooks were kept.\n\n"
                    "Click '3.  Validate Final Forecast' to compare "
                    "the completed Excel totals to the individual "
                    "Tableau workbooks."
                ),
                parent=self.winfo_toplevel(),
            )

    # =========================================================
    # STEP 3 - INDIVIDUAL TABLEAU VALIDATION
    # =========================================================

    def _start_validation(self) -> None:
        folder = Path(self.output_folder_var.get().strip())

        if not self.output_folder_var.get().strip() or not folder.is_dir():
            messagebox.showerror(
                "Forecast Output Folder",
                "Select the existing forecast output folder before validation.",
                parent=self.winfo_toplevel(),
            )
            return
        period = self._period()

        confirmed = messagebox.askyesno(
            "Validate Final Forecast",
            (
                f"Validate the completed {period['display']} forecast?\n\n"
                "Each individual Tableau workbook will use the same "
                "Forecast Start and Forecast Type/Forecast Measure "
                "for this selected period."
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return
        if not confirm_close_tableau(self):
            return

        self._set_busy(True)
        self.status_var.set("Validating Final Forecast")
        self.status_description_var.set(
            "Comparing final Excel Grand Totals to individual Tableau workbooks."
        )
        self.progress_var.set(82)

        threading.Thread(
            target=self._run_validation,
            daemon=True,
        ).start()

    def _run_validation(self) -> None:
        try:
            period = self._period()
            result = validate_monthly_forecast(
                forecast_folder=self.output_folder_var.get().strip(),
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                tolerance=5.00,
            )
            self.after(0, lambda r=result: self._validation_complete(r))
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, lambda e=exc: self._process_failed(e))

    def _validation_complete(self, result: dict) -> None:
        get_attention(self)
        passed = int(result.get("passed", 0))
        failed = int(result.get("failed", 0))
        action = int(result.get("action_required", 0))

        self.validation_passed_var.set(str(passed))
        self.validation_failed_var.set(str(failed))
        self.action_required_var.set(str(action))
        self.progress_var.set(100)
        report_path = result.get("validation_report")

        if report_path:

            self.validation_output_var.set(report_path)

            self.open_validation_button.configure(state="normal")

        if failed or action:
            self.status_var.set("Validation Review Required")
            self.status_description_var.set(
                f"{passed} passed, {failed} failed, " f"{action} require action."
            )
        else:
            self.status_var.set("Forecast Reconciled")
            self.status_description_var.set(
                f"All {passed} validated routes reconciled within tolerance."
            )

        self.prepare_button.configure(state="normal")
        self.populate_button.configure(state="normal")
        self.validate_button.configure(state="normal")

        message = (
            "Final forecast validation complete.\n\n"
            f"Passed: {passed}\n"
            f"Failed: {failed}\n"
            f"Action Required: {action}"
        )

        if failed or action:
            messagebox.showwarning(
                "Forecast Validation Complete",
                message,
                parent=self.winfo_toplevel(),
            )
        else:
            messagebox.showinfo(
                "Forecast Validation Complete",
                message,
                parent=self.winfo_toplevel(),
            )

    # =========================================================
    # COMMON
    # =========================================================

    def _process_failed(self, error: Exception) -> None:
        get_attention(self, success=False)
        self.progress_var.set(0)
        self.status_var.set("Process Failed")
        self.status_description_var.set(explain(error))
        self.prepare_button.configure(state="normal")
        self.populate_button.configure(state="normal")
        self.validate_button.configure(state="normal")

        messagebox.showerror(
            "Monthly Forecast Automation",
            friendly(error),
            parent=self.winfo_toplevel(),
        )

    def _reset_view(self) -> None:
        self.status_var.set("Ready")
        self.status_description_var.set(
            "Select the forecast period and the three required locations."
        )
        self.progress_var.set(0)
        self.files_created_var.set("0")
        self.worksheets_updated_var.set("0")
        self.validation_passed_var.set("0")
        self.validation_failed_var.set("0")
        self.action_required_var.set("0")
        self.output_var.set("Not created yet")
        self.open_output_button.configure(state="disabled")
        self.populate_button.configure(state="normal")
        self.validate_button.configure(state="normal")
        self.prepare_button.configure(state="normal")
        self.open_validation_button.configure(state="disabled")
        self._update_validation_output()

    def _open_output_folder(self) -> None:
        folder = self.output_folder_var.get().strip()
        if not folder:
            return
        path = Path(folder)
        if not path.exists():
            messagebox.showerror(
                "Folder Not Found",
                str(path),
                parent=self.winfo_toplevel(),
            )
            return
        os.startfile(path)
