from __future__ import annotations
from shared.friendly_errors import PlainLanguageError, explain, friendly

import logging

import os
import threading
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox
from shared.forecast_output_paths import (
    get_forecast_validation_output_folder,
)

import ttkbootstrap as tb

from automations.forecast.mapping import FORECAST_CATEGORIES, routes_for
from automations.forecast.service import (
    build_forecast_period,
    check_tableau_forecast_type,
    default_forecast_start,
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
from ui.forecast_messages import status_tiles, step2_message, step3_message

log = logging.getLogger(__name__)

class ForecastRefreshView(tb.Frame):
    """
    Monthly Forecast GUI.

    One selected forecast period controls:
      - Excel Forecast Type
      - Tableau Forecast Start and Forecast Type filter (the page's
        "Tableau Forecast Type"), in the consolidated and every individual .twb
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
        self.forecast_start_var = tb.StringVar()
        self.tableau_forecast_type_var = tb.StringVar()
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
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="w")

        tb.Label(
            self,
            text=(
                "Create the monthly forecast, populate it from Tableau, "
                "and validate the final totals."
            ),
            font=("Segoe UI", 10),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(2, 8))

        # Scrolls on small or zoomed-in screens so nothing is ever cut off.
        try:
            from ttkbootstrap.scrolled import ScrolledFrame

            content = ScrolledFrame(self, autohide=True)
        except Exception:
            content = tb.Frame(self)
        content.pack(fill="both", expand=True)

        # 60% parameters and steps / 40% process status
        content.columnconfigure(0, weight=3, uniform="panels")
        content.columnconfigure(1, weight=2, uniform="panels")
        content.rowconfigure(0, weight=1)

        left = tb.Frame(content)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_parameters_panel(left)
        self._build_steps_panel(left)
        self._build_status_panel(content)

    def _build_parameters_panel(self, parent) -> None:
        panel = tb.Labelframe(parent, text="Forecast Parameters", padding=(14, 8))
        panel.pack(fill="x")
        panel.columnconfigure(1, weight=1)

        # Period: month, year and fiscal year on one row.
        period = tb.Frame(panel)
        period.grid(row=0, column=0, columnspan=3, sticky="ew")
        for col in range(5):
            period.columnconfigure(col, weight=1, uniform="period")

        def field(col, text):
            tb.Label(period, text=text, font=("Segoe UI", 9, "bold")).grid(
                row=0, column=col, sticky="w", padx=(0 if col == 0 else 8, 0)
            )

        field(0, "Forecast Month")
        field(1, "Forecast Year")
        field(2, "Fiscal Year")
        field(3, "Forecast Start")
        field(4, "Tableau Forecast Type")

        self.forecast_month_combo = tb.Combobox(
            period, values=list(MONTH_LOOKUP), state="readonly"
        )
        self.forecast_month_combo.grid(row=1, column=0, sticky="ew", pady=(2, 0))
        self.forecast_month_combo.set("October")

        self.forecast_year_combo = tb.Combobox(
            period,
            values=["2026", "2027", "2028", "2029", "2030"],
            state="readonly",
        )
        self.forecast_year_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(2, 0))
        self.forecast_year_combo.set("2026")

        self.fiscal_year_combo = tb.Combobox(
            period, values=["FY27", "FY28", "FY29", "FY30"], state="readonly"
        )
        self.fiscal_year_combo.grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=(2, 0))
        self.fiscal_year_combo.set("FY27")

        for combo in (
            self.fiscal_year_combo,
            self.forecast_month_combo,
            self.forecast_year_combo,
        ):
            combo.bind("<<ComboboxSelected>>", self._month_changed)

        # Defaults to the 1st of the forecast month; can be typed over (MM/DD/YYYY).
        self.forecast_start_entry = tb.Entry(period, textvariable=self.forecast_start_var)
        self.forecast_start_entry.grid(row=1, column=3, sticky="ew", padx=(8, 0), pady=(2, 0))
        self.forecast_start_entry.bind("<FocusOut>", self._update_period_summary)
        self.forecast_start_entry.bind("<Return>", self._update_period_summary)

        # Forecast Type for the Tableau filters only - consolidated and every
        # individual workbook. The Excel workbooks' Forecast Type column always
        # follows the selected month. Defaults to the month's; can be typed over.
        self.tableau_forecast_type_combo = tb.Combobox(
            period, textvariable=self.tableau_forecast_type_var
        )
        self.tableau_forecast_type_combo.grid(row=1, column=4, sticky="ew", padx=(8, 0), pady=(2, 0))
        self.tableau_forecast_type_combo.bind("<<ComboboxSelected>>", self._update_period_summary)
        self.tableau_forecast_type_combo.bind("<FocusOut>", self._update_period_summary)
        self.tableau_forecast_type_combo.bind("<Return>", self._update_period_summary)

        tb.Label(
            panel,
            textvariable=self.period_summary_var,
            bootstyle="info",
            font=("Segoe UI", 9, "bold"),
            justify="left",
        ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(6, 8))

        # Paths: label | box | Browse, one line each.
        for row, (label, variable, command) in enumerate(
            (
                ("Previous Forecast Folder", self.source_folder_var, self._browse_source_folder),
                ("Forecast Output Folder", self.output_folder_var, self._browse_output_folder),
                ("Tableau Forecast Workbook", self.tableau_workbook_var, self._browse_tableau_workbook),
            ),
            start=2,
        ):
            tb.Label(panel, text=label, font=("Segoe UI", 9, "bold")).grid(
                row=row, column=0, sticky="w", padx=(0, 10), pady=3
            )
            tb.Entry(panel, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=3)
            tb.Button(
                panel, text="Browse", bootstyle="secondary-outline", command=command
            ).grid(row=row, column=2, sticky="e", padx=(8, 0), pady=3)

        tb.Label(panel, text="Existing Files", font=("Segoe UI", 9, "bold")).grid(
            row=5, column=0, sticky="w", padx=(0, 10), pady=3
        )
        overwrite = tb.Combobox(
            panel,
            textvariable=self.overwrite_var,
            values=["skip", "replace", "cancel"],
            state="readonly",
            width=12,
        )
        overwrite.grid(row=5, column=1, sticky="w", pady=3)
        overwrite.set("skip")


    def _build_steps_panel(self, parent) -> None:
        """Step 1 across the top; steps 2 and 3 as a grid: one column per category."""
        panel = tb.Labelframe(parent, text="Run", padding=(14, 8))
        panel.pack(fill="x", pady=(8, 0))

        self.prepare_button = tb.Button(
            panel,
            text="1.  Generate Forecast & Open Tableau",
            bootstyle="success",
            command=self._start_prepare,
        )
        self.prepare_button.pack(fill="x", ipady=2)

        grid = tb.Frame(panel)
        grid.pack(fill="x", pady=(10, 0))
        grid.columnconfigure(0, weight=0)
        for col in range(1, len(FORECAST_CATEGORIES) + 1):
            grid.columnconfigure(col, weight=1, uniform="categories")

        # Header: category and its sheet count.
        for col, category in enumerate(FORECAST_CATEGORIES, start=1):
            label = "Other (SAC, BTL)" if category == "Other" else category
            tb.Label(
                grid,
                text=f"{label}\n{len(routes_for([category]))} sheets",
                font=("Segoe UI", 9, "bold"),
                justify="center",
            ).grid(row=0, column=col, pady=(0, 4))

        self.populate_buttons, self.category_status_vars = self._step_row(
            grid,
            row=1,
            title="2.  Update from\n     Tableau",
            button_text="Update",
            style="info-outline",
            command=self._start_population,
        )
        tb.Separator(grid).grid(row=3, column=0, columnspan=5, sticky="ew", pady=6)
        self.validate_buttons, self.validation_status_vars = self._step_row(
            grid,
            row=4,
            title="3.  Validate",
            button_text="Validate",
            style="primary-outline",
            command=self._start_validation,
        )

    def _step_row(self, grid, *, row, title, button_text, style, command):
        """One step: a button and a status line per category."""
        tb.Label(grid, text=title, font=("Segoe UI", 10, "bold"), justify="left").grid(
            row=row, column=0, rowspan=2, sticky="w", padx=(0, 10)
        )
        buttons, statuses = {}, {}
        for col, category in enumerate(FORECAST_CATEGORIES, start=1):
            button = tb.Button(
                grid,
                text=button_text,
                bootstyle=style,
                command=lambda c=category: command(c),
            )
            button.grid(row=row, column=col, sticky="ew", padx=(0 if col == 1 else 4, 0))
            buttons[category] = button
            status = tb.StringVar(value="Not run")
            statuses[category] = status
            tb.Label(
                grid,
                textvariable=status,
                bootstyle="secondary",
                font=("Segoe UI", 8),
                wraplength=130,
                justify="center",
            ).grid(row=row + 1, column=col, pady=(2, 0))
        return buttons, statuses

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
        self._stat(stats, 1, self.worksheets_updated_var, "Sheets")
        self._stat(stats, 2, self.validation_passed_var, "Passed")
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
        period = build_forecast_period(  # the folder depends only on the month
            fiscal_year=self.fiscal_year_combo.get(),
            forecast_month=MONTH_LOOKUP[self.forecast_month_combo.get()],
            forecast_year=int(self.forecast_year_combo.get()),
        )

        try:
            path = get_forecast_validation_output_folder(
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                create=False,
            )
        except Exception as exc:  # shown on the page; a run stops with the full message
            self.validation_output_var.set(
                "Not available: the shared Validation Results folder is not in your OneDrive. "
                + friendly(exc).split("\n\n", 1)[-1]
            )
            return

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

    def _forecast_start(self) -> date:
        text = self.forecast_start_var.get().strip()
        for pattern in ("%m/%d/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, pattern).date()
            except ValueError:
                pass
        raise PlainLanguageError(
            f'Forecast Start "{text}" is not a date. Type it as MM/DD/YYYY, for example 10/01/2026.'
        )

    def _period(self) -> dict:
        period = build_forecast_period(
            fiscal_year=self.fiscal_year_combo.get(),
            forecast_month=MONTH_LOOKUP[self.forecast_month_combo.get()],
            forecast_year=int(self.forecast_year_combo.get()),
            forecast_start=self._forecast_start(),
        )
        chosen = self.tableau_forecast_type_var.get().strip() or period["forecast_type"]
        period["tableau_forecast_type"] = check_tableau_forecast_type(chosen, period["fiscal_year"])
        return period

    def _tableau_type_choices(self) -> list[str]:
        """The selected month's Forecast Type and the two before it."""
        month = MONTH_LOOKUP[self.forecast_month_combo.get()]
        year = int(self.forecast_year_combo.get())
        fiscal = self.fiscal_year_combo.get()
        choices = []
        for _ in range(3):
            choices.append(f"{fiscal} Forecast_{month:02d}{year}")
            month, year = (12, year - 1) if month == 1 else (month - 1, year)
        return choices

    def _checked_period(self) -> dict | None:
        """The period, or None after telling the user what is wrong with it."""
        try:
            return self._period()
        except PlainLanguageError as exc:
            messagebox.showerror("Forecast Parameters", str(exc), parent=self.winfo_toplevel())
            return None

    def _month_changed(self, _event=None) -> None:
        """A new month resets Forecast Start to the 1st of that month."""
        start = default_forecast_start(
            MONTH_LOOKUP[self.forecast_month_combo.get()],
            int(self.forecast_year_combo.get()),
        )
        self.forecast_start_var.set(f"{start:%m/%d/%Y}")
        self.tableau_forecast_type_var.set("")   # back to the new month's default
        self._update_period_summary()

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
        if not self.forecast_start_var.get().strip():
            start = default_forecast_start(
                MONTH_LOOKUP[self.forecast_month_combo.get()],
                int(self.forecast_year_combo.get()),
            )
            self.forecast_start_var.set(f"{start:%m/%d/%Y}")
        self.tableau_forecast_type_combo.configure(values=self._tableau_type_choices())
        if not self.tableau_forecast_type_var.get().strip():
            self.tableau_forecast_type_var.set(self._tableau_type_choices()[0])
        try:
            period = self._period()
        except PlainLanguageError as exc:
            self.period_summary_var.set(str(exc))
            return
        self.period_summary_var.set(
            f"Selected Period: {period['display']}\n"
            f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
            f"Forecast Type: {period['forecast_type']} (Excel workbooks); "
            f"{period['tableau_forecast_type']} (Tableau filters)"
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

    def _set_validate_state(self, state: str) -> None:
        for button in self.validate_buttons.values():
            button.configure(state=state)

    def _set_populate_state(self, state: str) -> None:
        for button in self.populate_buttons.values():
            button.configure(state=state)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.prepare_button.configure(state=state)
        if busy:
            self._set_populate_state("disabled")
            self._set_validate_state("disabled")

    # =========================================================
    # STEP 1 - CREATE + PREPARE TABLEAU
    # =========================================================

    def _start_prepare(self) -> None:
        if not self._validate_inputs():
            return

        period = self._checked_period()
        if period is None:
            return
        confirmed = messagebox.askyesno(
            "Run Monthly Forecast",
            (
                f"Run the forecast for {period['display']}?\n\n"
                f"Excel workbooks - Forecast Type column: {period['forecast_type']}\n\n"
                f"Tableau workbooks (consolidated and individual):\n"
                f"  Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"  Forecast Type filter: {period['tableau_forecast_type']}"
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
                forecast_start=period["forecast_start"],
                overwrite_mode=inputs["overwrite_mode"],
                tableau_forecast_type=period["tableau_forecast_type"],
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
        for figure in (self.worksheets_updated_var, self.validation_passed_var,
                       self.validation_failed_var, self.action_required_var):
            figure.set("-")
        self.output_var.set(result["output_folder"])
        self.open_output_button.configure(state="normal")

        self.progress_var.set(40)
        self.status_var.set("Waiting for Tableau")
        self.status_description_var.set(
            f"Tableau is open for {period['display']}. "
            f"Forecast Start = {period['forecast_start']:%m/%d/%Y}. "
            f"Forecast Type filter = {period['tableau_forecast_type']}. "
            "Let Tableau finish refreshing, then click Update."
        )

        self.prepare_button.configure(state="normal")
        self._set_populate_state("normal")
        self._set_validate_state("normal")

        failures = excel.get("failures") or []
        skipped = excel.get("skipped_files") or []
        older = excel.get("older_sources") or []
        warnings = excel.get("warnings") or []

        lines = [
            f"{period['display']} forecast workbooks:",
            f"  Created: {excel.get('files_created', 0)}",
            f"  Kept (already existed): {len(skipped)}",
            f"  Failed: {len(failures)}",
            "",
            self._formula_check_line(excel.get("formula_audits") or []),
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
        if warnings:
            lines += ["", "Check these before continuing:"]
            lines += [
                f"  - {item['file']}: {'; '.join(item['problems'])}" for item in warnings[:8]
            ]
            if len(warnings) > 8:
                lines.append(f"  ...and {len(warnings) - 8} more (see the log).")
        lines += [
            "",
            f"Excel Forecast Type column: {period['forecast_type']}",
            f"Tableau: Forecast Start {period['forecast_start']:%m/%d/%Y}, "
            f"Forecast Type filter {period['tableau_forecast_type']}",
            "",
            "Tableau is opening. When it has finished loading and refreshing, click "
            "Update for each category under step 2.",
        ]
        (messagebox.showwarning if failures or older or warnings else messagebox.showinfo)(
            "Forecast Workbooks Created",
            "\n".join(lines),
            parent=self.winfo_toplevel(),
        )

    @staticmethod
    def _formula_check_line(audits: list[dict]) -> str:
        """One line saying whether every formula came through Step 1 unchanged."""
        if not audits:
            return "Formula check: no new workbooks were created, so nothing was checked."
        total = sum(item.get("source_formula_count", 0) for item in audits)
        bad = [item for item in audits if not item.get("formulas_preserved", True)]
        if not bad:
            return (
                f"Formula check: all {total:,} formulas in {len(audits)} workbook(s) "
                "are the same as last month's."
            )
        problems = sum(
            item.get("missing_formula_count", 0) + item.get("changed_formula_count", 0)
            for item in bad
        )
        return (
            f"Formula check: {problems:,} of {total:,} formulas are missing or changed "
            f"in {len(bad)} of {len(audits)} workbook(s) (listed below)."
        )

    # =========================================================
    # STEP 2 - CONSOLIDATED TABLEAU -> EXCEL
    # =========================================================

    def _start_population(self, category: str) -> None:
        if not self._validate_inputs():
            return

        period = self._checked_period()
        if period is None:
            return
        confirmed = messagebox.askyesno(
            "Update Workbooks from Tableau",
            (
                f"Update the {category} workbooks ({len(routes_for([category]))} sheets)?\n\n"
                "The app copies what the Tableau workbook from Step 1 shows "
                "(it is opened again if it was closed). Nothing in it is changed.\n\n"
                f"Tableau: Forecast Start {period['forecast_start']:%m/%d/%Y}, "
                f"Forecast Type filter {period['tableau_forecast_type']}\n\n"
                "Continue?"
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return

        # Read everything on the UI thread; the worker must not touch Tk.
        inputs = {
            "period": period,
            "category": category,
            "output_folder": self.output_folder_var.get().strip(),
            "tableau_workbook": self.tableau_workbook_var.get().strip(),
        }

        self._set_busy(True)
        self.category_status_vars[category].set("Running...")
        self.status_var.set(f"Updating {category} from Tableau")
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
                forecast_start=period["forecast_start"],
                categories=[inputs["category"]],
                tableau_forecast_type=period["tableau_forecast_type"],
            )
            self.after(0, lambda r=result, c=inputs["category"]: self._population_complete(c, r))
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, lambda e=exc, c=inputs["category"]: self._population_failed(c, e))

    def _population_failed(self, category: str, error: Exception) -> None:
        self.category_status_vars[category].set("Stopped - see message")
        self._process_failed(error)

    def _population_complete(self, category: str, result: dict) -> None:
        get_attention(self)
        failed = result["failed"]
        self.category_status_vars[category].set(
            f"{result['routes_passed']} updated"
            + (f", {result['routes_failed']} not" if failed else "")
        )
        self._show_tiles(2, category, result)
        self.progress_var.set(75)

        self.prepare_button.configure(state="normal")
        self._set_populate_state("normal")
        self._set_validate_state("normal")

        attention, text = step2_message(category, result)
        (messagebox.showwarning if attention else messagebox.showinfo)(
            f"Update from Tableau - {category}",
            text,
            parent=self.winfo_toplevel(),
        )

    # =========================================================
    # STEP 3 - INDIVIDUAL TABLEAU VALIDATION
    # =========================================================

    def _start_validation(self, category: str) -> None:
        folder = Path(self.output_folder_var.get().strip())

        if not self.output_folder_var.get().strip() or not folder.is_dir():
            messagebox.showerror(
                "Forecast Output Folder",
                "Select the existing forecast output folder before validation.",
                parent=self.winfo_toplevel(),
            )
            return
        period = self._checked_period()
        if period is None:
            return

        confirmed = messagebox.askyesno(
            "Validate Final Forecast",
            (
                f"Validate the {category} sheets for {period['display']}?\n\n"
                f"Forecast Start: {period['forecast_start']:%m/%d/%Y}\n"
                f"Forecast Type: {period['tableau_forecast_type']} (same as the consolidated "
                "workbook)\n\n"
                "Each individual Tableau workbook is set to this period and saved, then "
                "opened, and its total is compared to the sheet's Grand Total "
                "(within $1,000). The month's report keeps the other categories' results."
            ),
            parent=self.winfo_toplevel(),
        )
        if not confirmed:
            return
        if not confirm_close_tableau(self):
            return

        # Read everything on the UI thread; the worker must not touch Tk.
        inputs = {
            "period": period,
            "category": category,
            "output_folder": self.output_folder_var.get().strip(),
        }

        self._set_busy(True)
        self.validation_status_vars[category].set("Running...")
        self.status_var.set(f"Validating {category}")
        self.status_description_var.set(
            "Comparing final Excel Grand Totals to individual Tableau workbooks."
        )
        self.progress_var.set(82)

        threading.Thread(
            target=self._run_validation,
            args=(inputs,),
            daemon=True,
        ).start()

    def _run_validation(self, inputs: dict) -> None:
        category = inputs["category"]
        try:
            period = inputs["period"]
            result = validate_monthly_forecast(
                forecast_folder=inputs["output_folder"],
                fiscal_year=period["fiscal_year"],
                forecast_month=period["forecast_month"],
                forecast_year=period["forecast_year"],
                forecast_start=period["forecast_start"],
                tolerance=1000.00,
                categories=[category],
                tableau_forecast_type=period["tableau_forecast_type"],
            )
            self.after(0, lambda r=result, c=category: self._validation_complete(c, r))
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, lambda e=exc, c=category: self._validation_failed(c, e))

    def _validation_failed(self, category: str, error: Exception) -> None:
        self.validation_status_vars[category].set("Stopped - see message")
        self._process_failed(error)

    def _validation_complete(self, category: str, result: dict) -> None:
        get_attention(self)
        passed = int(result.get("passed", 0))
        failed = int(result.get("failed", 0))
        action = int(result.get("action_required", 0))
        self.validation_status_vars[category].set(
            f"{passed} pass"
            + (f", {failed} fail" if failed else "")
            + (f", {action} action" if action else "")
        )

        self._show_tiles(3, category, result)
        self.progress_var.set(100)
        report_path = result.get("validation_report")
        if report_path:
            self.validation_output_var.set(report_path)
            self.open_validation_button.configure(state="normal")

        self.prepare_button.configure(state="normal")
        self._set_populate_state("normal")
        self._set_validate_state("normal")

        attention, message = step3_message(category, result)
        (messagebox.showwarning if attention else messagebox.showinfo)(
            f"Validate - {category}",
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
        self._set_populate_state("normal")
        self._set_validate_state("normal")

        messagebox.showerror(
            "Monthly Forecast Automation",
            friendly(error),
            parent=self.winfo_toplevel(),
        )

    def _show_tiles(self, step: int, category: str, result: dict) -> None:
        """All five figures and the title come from this one run (one step, one category)."""
        tiles, title, description = status_tiles(step, category, result)
        self.files_created_var.set(str(tiles["workbooks"]))
        self.worksheets_updated_var.set(str(tiles["sheets"]))
        self.validation_passed_var.set(str(tiles["passed"]))
        self.validation_failed_var.set(str(tiles["failed"]))
        self.action_required_var.set(str(tiles["action"]))
        self.status_var.set(title)
        self.status_description_var.set(description)

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
        self._set_populate_state("normal")
        self._set_validate_state("normal")
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
