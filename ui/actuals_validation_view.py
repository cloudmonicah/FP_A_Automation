from __future__ import annotations
from shared.friendly_errors import explain, friendly

import logging

from datetime import date
from pathlib import Path

import os
import threading

import ttkbootstrap as tb
from tkinter import filedialog, messagebox

from automations.actuals.validation_output_folders import (
    create_validation_output_folders,
)

from automations.actuals.opex_validation_service import (
    run_live_opex_validation,
)

from automations.actuals.direct_cost_validation_service import (
    run_live_direct_cost_validation,
)
from ui.common import confirm_close_tableau, notify_finished

log = logging.getLogger(__name__)


class ActualsValidationView(tb.Frame):

    def __init__(
        self,
        parent,
    ) -> None:
        super().__init__(parent)

        # =====================================================
        # CLOSED PERIOD VARIABLES
        # =====================================================

        # Default period used while we are testing the build.
        self.closed_period_month_var = tb.StringVar(
            value="July",
        )

        self.closed_period_year_var = tb.StringVar(
            value="2026",
        )

        # =====================================================
        # FINANCIAL STATEMENT SOURCE VARIABLES
        # =====================================================

        self.aliv_fs_path_var = tb.StringVar(value="")

        self.cbl_fs_path_var = tb.StringVar(value="")

        # Values calculated after the user confirms the period.
        self.fiscal_year_var = tb.StringVar(
            value="Not resolved",
        )

        self.output_period_var = tb.StringVar(
            value="Not resolved",
        )

        # =====================================================
        # VALIDATION OUTPUT LOCATION
        # =====================================================

        self.validation_base_folder_var = tb.StringVar(value="")

        self.validation_base_folder = None

        self.validation_output_paths = None

        self.validation_month_folder = None

        # =====================================================
        # WORKSPACE STATUS VARIABLES
        # =====================================================

        self.workspace_status_var = tb.StringVar(
            value="Select and confirm the closed period.",
        )

        self.workspace_path_var = tb.StringVar(
            value="",
        )

        # =====================================================
        # VALIDATION RUN STATE
        # =====================================================
        #
        # These variables drive the validation status area.
        #
        # The actual validation work runs on a background
        # thread so Tableau automation does not freeze the GUI.
        # =====================================================

        self.validation_running = False

        self.overall_validation_status_var = tb.StringVar(value="NOT RUN")

        self.opex_status_var = tb.StringVar(value="NOT RUN")

        self.direct_cost_status_var = tb.StringVar(value="NOT RUN")

        self.revenue_status_var = tb.StringVar(value="NOT RUN")

        self.validation_progress_var = tb.StringVar(
            value=("Set the validation period and select the " "financial statements.")
        )
        # =====================================================
        # BUILD PAGE
        # =====================================================

        self.opex_tableau_path_var = tb.StringVar(value="")
        self.direct_cost_tableau_path_var = tb.StringVar(value="")
        self.revenue_tableau_path_var = tb.StringVar(value="")
        self.active_category = None

        # User-selected Tableau settings
        self.forecast_type_var = tb.StringVar(value="FY27 Forecast_092026")

        self.forecast_start_var = tb.StringVar(value="2026-09-01")

        self.company_selection_var = tb.StringVar(value="Both")

        # Additional filter/parameter values
        self.additional_filters_var = tb.StringVar(value="")
        self.additional_parameters_var = tb.StringVar(value="")
        self._build_view()

    # =========================================================
    # BUILD VIEW
    # =========================================================

    def _build_view(self) -> None:
        "Select consolidated sources; individual TWBs come from OPEX_MAPPINGS."
        self.columnconfigure(0, weight=1)
        tb.Label(self, text="Actuals Validation", font=("Segoe UI", 20, "bold")).grid(
            row=0, column=0, sticky="w", pady=(0, 4)
        )
        tb.Label(
            self,
            text="Compare closed-period actuals with finalized financial statements.",
            bootstyle="secondary",
        ).grid(row=1, column=0, sticky="w", pady=(0, 10))
        body = tb.Frame(self)
        body.grid(row=2, column=0, sticky="ew")
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        setup = tb.Labelframe(body, text="Period and source files", padding=12)
        setup.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        setup.columnconfigure(1, weight=1)
        period = tb.Frame(setup)
        period.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        period.columnconfigure(1, weight=1)
        tb.Label(period, text="Closed month").grid(
            row=0, column=0, sticky="w", padx=(0, 8)
        )
        self.closed_period_month_combo = tb.Combobox(
            period,
            textvariable=self.closed_period_month_var,
            state="readonly",
            width=13,
            values=[
                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December",
            ],
        )
        self.closed_period_month_combo.grid(row=0, column=1, sticky="ew")
        tb.Label(period, text="Year").grid(row=0, column=2, padx=(12, 8))
        self.closed_period_year_entry = tb.Entry(
            period, textvariable=self.closed_period_year_var, width=7
        )
        self.closed_period_year_entry.grid(row=0, column=3)
        sources = [
            (
                "ALIV statement",
                self.aliv_fs_path_var,
                self._browse_aliv_financial_statement,
            ),
            (
                "CBL statement",
                self.cbl_fs_path_var,
                self._browse_cbl_financial_statement,
            ),
        ]
        sources.extend(
            [
                (
                    "OPEX Consolidated Tableau",
                    self.opex_tableau_path_var,
                    lambda: self._browse_tableau_source(self.opex_tableau_path_var),
                ),
                (
                    "Direct Cost Consolidated Tableau",
                    self.direct_cost_tableau_path_var,
                    lambda: self._browse_tableau_source(
                        self.direct_cost_tableau_path_var
                    ),
                ),
                (
                    "Revenue Consolidated Tableau",
                    self.revenue_tableau_path_var,
                    lambda: self._browse_tableau_source(self.revenue_tableau_path_var),
                ),
            ]
        )
        for row, (label, variable, callback) in enumerate(sources, 1):
            tb.Label(setup, text=label).grid(
                row=row, column=0, sticky="w", padx=(0, 8), pady=5
            )
            tb.Entry(setup, textvariable=variable, width=24).grid(
                row=row, column=1, sticky="ew", pady=5
            )
            tb.Button(
                setup, text="Browse", command=callback, bootstyle="secondary-outline"
            ).grid(row=row, column=2, padx=(8, 0), pady=5)

        # =========================================================
        # TABLEAU FILTERS AND PARAMETERS
        # =========================================================

        settings_row = len(sources) + 4

        settings = tb.Labelframe(
            setup,
            text="Tableau Filters & Parameters",
            padding=12,
            bootstyle="primary",
        )

        settings.grid(
            row=settings_row,
            column=0,
            columnspan=3,
            sticky="ew",
            pady=(18, 12),
        )

        settings.columnconfigure(1, weight=1)

        # Forecast Type and Forecast Start Date
        fields = [
            ("Forecast Type", self.forecast_type_var),
            ("Forecast Start Date", self.forecast_start_var),
        ]

        for index, (label, variable) in enumerate(fields):

            tb.Label(
                settings,
                text=label,
            ).grid(
                row=index,
                column=0,
                sticky="w",
                padx=(0, 10),
                pady=5,
            )

            tb.Entry(
                settings,
                textvariable=variable,
            ).grid(
                row=index,
                column=1,
                sticky="ew",
                pady=5,
            )

        # Company selection - OUTSIDE the loop
        tb.Label(
            settings,
            text="Company",
        ).grid(
            row=len(fields),
            column=0,
            sticky="w",
            pady=5,
        )

        tb.Combobox(
            settings,
            textvariable=self.company_selection_var,
            values=["Both", "ALIV", "CBL"],
            state="readonly",
        ).grid(
            row=len(fields),
            column=1,
            sticky="ew",
            pady=5,
        )

        resolved = tb.Frame(setup)
        resolved.grid(
            row=len(sources) + 1, column=0, columnspan=3, sticky="ew", pady=(10, 0)
        )
        tb.Button(
            resolved,
            text="Set Validation Period",
            command=self._confirm_closed_period,
            bootstyle="primary",
        ).pack(side="left", padx=(0, 12))
        tb.Label(resolved, textvariable=self.fiscal_year_var, bootstyle="success").pack(
            side="left"
        )
        tb.Label(
            resolved, textvariable=self.output_period_var, bootstyle="secondary"
        ).pack(side="left", padx=8)
        self.workspace_status_label = tb.Label(
            setup,
            textvariable=self.workspace_status_var,
            bootstyle="secondary",
            wraplength=480,
        )
        self.workspace_status_label.grid(
            row=len(sources) + 2, column=0, columnspan=3, sticky="ew", pady=(8, 0)
        )
        self.workspace_path_label = tb.Label(
            setup,
            textvariable=self.workspace_path_var,
            bootstyle="secondary",
            wraplength=480,
        )
        self.workspace_path_label.grid(
            row=len(sources) + 3, column=0, columnspan=3, sticky="ew", pady=(0, 0)
        )
        run = tb.Labelframe(body, text="Validation progress", padding=12)
        run.grid(row=0, column=1, sticky="nsew")
        run.columnconfigure(0, weight=1)
        for row, (name, variable, attribute) in enumerate(
            [
                ("OPEX", self.opex_status_var, "opex_status_label"),
                (
                    "Direct Cost",
                    self.direct_cost_status_var,
                    "direct_cost_status_label",
                ),
                ("Revenue", self.revenue_status_var, "revenue_status_label"),
            ]
        ):
            tb.Label(run, text=name, font=("Segoe UI", 10, "bold")).grid(
                row=row, column=0, sticky="w", pady=6
            )
            label = tb.Label(run, textvariable=variable, bootstyle="secondary")
            label.grid(row=row, column=1, sticky="e", padx=(12, 0), pady=6)
            setattr(self, attribute, label)
        tb.Separator(run).grid(row=3, column=0, columnspan=2, sticky="ew", pady=8)
        self.validation_progress_label = tb.Label(
            run,
            textvariable=self.validation_progress_var,
            bootstyle="secondary",
            wraplength=330,
        )
        self.validation_progress_label.grid(
            row=4, column=0, columnspan=2, sticky="ew", pady=4
        )
        self.validation_progress_bar = tb.Progressbar(
            run, mode="indeterminate", bootstyle="primary-striped"
        )
        self.validation_progress_bar.grid(
            row=5, column=0, columnspan=2, sticky="ew", pady=8
        )
        tb.Label(run, text="Overall result", font=("Segoe UI", 10, "bold")).grid(
            row=6, column=0, sticky="w", pady=6
        )
        self.overall_validation_status_label = tb.Label(
            run,
            textvariable=self.overall_validation_status_var,
            bootstyle="secondary",
            font=("Segoe UI", 10, "bold"),
        )
        self.overall_validation_status_label.grid(row=6, column=1, sticky="e", pady=6)
        actions = tb.Frame(self)
        actions.grid(row=3, column=0, sticky="ew", pady=(12, 0))
        self.run_validation_button = tb.Button(
            actions,
            text="Run OPEX",
            bootstyle="primary",
            command=lambda: self._start_actuals_validation("OPEX"),
        )
        self.run_validation_button.pack(side="left", padx=(0, 8))
        self.run_direct_cost_button = tb.Button(
            actions,
            text="Run Direct Cost",
            bootstyle="info",
            command=lambda: self._start_actuals_validation("Direct Cost"),
        )
        self.run_direct_cost_button.pack(side="left", padx=(0, 8))
        self.run_revenue_button = tb.Button(
            actions,
            text="Run Revenue",
            bootstyle="warning",
            command=lambda: self._start_actuals_validation("Revenue"),
        )
        self.run_revenue_button.pack(side="left", padx=(0, 8))
        self.open_results_button = tb.Button(
            actions,
            text="Open Results Folder",
            bootstyle="secondary-outline",
            command=self._open_validation_results,
            state="disabled",
        )
        self.open_results_button.pack(side="left", padx=10)

        def fit_text(event):
            self.workspace_status_label.configure(wraplength=max(180, event.width - 30))
            self.workspace_path_label.configure(wraplength=max(180, event.width - 30))

        setup.bind("<Configure>", fit_text)
        run.bind(
            "<Configure>",
            lambda event: self.validation_progress_label.configure(
                wraplength=max(140, event.width - 30)
            ),
        )

    def _browse_tableau_source(self, variable):
        path = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="Select Tableau source workbook",
            filetypes=[("Tableau workbook", "*.twb")],
        )
        if path:
            variable.set(path)

    def _get_closed_period(
        self,
    ) -> date:

        month_name = self.closed_period_month_var.get().strip()

        year_text = self.closed_period_year_var.get().strip()

        # -----------------------------------------------------
        # MONTH LOOKUP
        # -----------------------------------------------------

        month_lookup = {
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

        # -----------------------------------------------------
        # VALIDATE MONTH
        # -----------------------------------------------------

        if month_name not in month_lookup:
            raise ValueError("Please select a valid closed month.")

        # -----------------------------------------------------
        # VALIDATE YEAR
        # -----------------------------------------------------

        if not year_text.isdigit():
            raise ValueError("Closed year must be numeric.")

        year = int(year_text)

        if year < 2000 or year > 2100:
            raise ValueError("Please enter a valid four-digit year.")

        month = month_lookup[month_name]

        return date(
            year,
            month,
            1,
        )

    # =========================================================
    # GET FISCAL YEAR
    # =========================================================

    def _get_fiscal_year(
        self,
        closed_period: date,
    ) -> str:
        """
        Resolve the July-June fiscal year.

        Examples:
            July 2026     -> FY27
            December 2026 -> FY27
            January 2027  -> FY27
            June 2027     -> FY27
            July 2027     -> FY28
        """

        if closed_period.month >= 7:

            fiscal_year_end = closed_period.year + 1

        else:

            fiscal_year_end = closed_period.year

        return f"FY{str(fiscal_year_end)[-2:]}"

    # =========================================================
    # GET OUTPUT PERIOD
    # =========================================================

    def _get_output_period(
        self,
        closed_period: date,
    ) -> str:
        """
        Create the standard sortable month folder name.

        Example:
            July 2026 -> 2026-07 - July
        """

        return f"{closed_period:%Y-%m} - " f"{closed_period:%B}"

    # =========================================================
    # CONFIRM CLOSED PERIOD
    # =========================================================

    def _confirm_closed_period(
        self,
    ) -> None:
        """
        Confirm the selected closed period and create/reuse the
        appropriate validation workspace.

        This method does NOT perform actual validation yet.
        """

        try:

            # -------------------------------------------------
            # GET CLOSED PERIOD
            # -------------------------------------------------

            closed_period = self._get_closed_period()

            # -------------------------------------------------
            # RESOLVE FISCAL YEAR
            # -------------------------------------------------

            fiscal_year = self._get_fiscal_year(
                closed_period,
            )

            # -------------------------------------------------
            # RESOLVE OUTPUT PERIOD
            # -------------------------------------------------

            output_period = self._get_output_period(
                closed_period,
            )

            # -------------------------------------------------
            # UPDATE PERIOD VALUES ON SCREEN
            # -------------------------------------------------

            self.fiscal_year_var.set(
                fiscal_year,
            )

            self.output_period_var.set(
                output_period,
            )

            # -------------------------------------------------
            # VALIDATE OUTPUT LOCATION
            # -------------------------------------------------

            from automations.actuals.validation_output_folders import (
                find_actuals_validation_results_folder,
            )

            self.validation_base_folder = Path(find_actuals_validation_results_folder())
            self.validation_base_folder_var.set(str(self.validation_base_folder))
            # -------------------------------------------------
            # CREATE / REUSE OUTPUT FOLDERS
            # -------------------------------------------------

            output_paths = create_validation_output_folders(
                base_output_folder=(self.validation_base_folder),
                closed_period=closed_period,
            )

            # -------------------------------------------------
            # SAVE OUTPUT PATHS FOR LATER STEPS
            # -------------------------------------------------

            self.validation_output_paths = output_paths

            self.validation_month_folder = output_paths["month_folder"]

            # -------------------------------------------------
            # UPDATE GUI STATUS
            # -------------------------------------------------

            self.workspace_status_var.set("✓ Validation period ready")

            self.workspace_path_var.set(str(self.validation_month_folder))

            self.workspace_status_label.configure(
                bootstyle="success",
            )

            # -------------------------------------------------
            # TERMINAL TEST OUTPUT
            # -------------------------------------------------

            print()
            print("=" * 70)
            print("ACTUALS VALIDATION WORKSPACE")
            print("=" * 70)

            print(
                "Closed Period:",
                closed_period.strftime("%B %Y"),
            )

            print(
                "Fiscal Year:",
                fiscal_year,
            )

            print(
                "Output Period:",
                output_period,
            )

            print(
                "Workspace:",
                self.validation_month_folder,
            )

            print()
            print("Validation folders:")

            for name, path in output_paths.items():

                print(f"  {name}: {path}")

            print("=" * 70)

        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)

            # -------------------------------------------------
            # UPDATE GUI WITH ERROR STATE
            # -------------------------------------------------

            self.workspace_status_var.set("Validation period could not be created.")

            self.workspace_path_var.set("")

            self.workspace_status_label.configure(
                bootstyle="danger",
            )

            # -------------------------------------------------
            # SHOW ACTUAL ERRORS TO USER
            # -------------------------------------------------

            messagebox.showerror(
                "Validation Workspace Error",
                friendly(exc),
                parent=self.winfo_toplevel(),
            )

        # =========================================================

    # BROWSE - ALIV FINANCIAL STATEMENT
    # =========================================================

    def _browse_aliv_financial_statement(
        self,
    ) -> None:
        """
        Allow the user to select the finalized ALIV
        month-end financial statement workbook.
        """

        selected_file = filedialog.askopenfilename(
            title="Select ALIV Financial Statement",
            filetypes=[
                (
                    "Excel Workbooks",
                    "*.xlsx *.xlsm",
                ),
            ],
        )

        if selected_file:
            self.aliv_fs_path_var.set(selected_file)

    # =========================================================
    # BROWSE - CBL FINANCIAL STATEMENT
    # =========================================================

    def _browse_cbl_financial_statement(
        self,
    ) -> None:
        """
        Allow the user to select the finalized CBL
        month-end financial statement workbook.
        """

        selected_file = filedialog.askopenfilename(
            title="Select CBL Financial Statement",
            filetypes=[
                (
                    "Excel Workbooks",
                    "*.xlsx *.xlsm",
                ),
            ],
        )

        if selected_file:
            self.cbl_fs_path_var.set(selected_file)

    # =========================================================
    # BROWSE - VALIDATION OUTPUT LOCATION
    # =========================================================

    def _browse_validation_output_folder(
        self,
    ) -> None:
        """
        Allow the user to choose the base folder where
        Actuals Validation results should be stored.

        Python will create Forecast_Validation_Results
        underneath the selected folder.
        """

        selected_folder = filedialog.askdirectory(
            title="Select Validation Output Location",
        )

        # User cancelled the folder picker.
        if not selected_folder:
            return

        # Save selected location.
        self.validation_base_folder_var.set(selected_folder)

        self.validation_base_folder = Path(selected_folder)

        # Reset status because the output location changed.
        self.workspace_status_var.set(
            "Output location selected. Set the validation period."
        )

        self.workspace_path_var.set("")

        self.workspace_status_label.configure(
            bootstyle="secondary",
        )

    # =========================================================
    # START ACTUALS VALIDATION
    # =========================================================

    def _start_actuals_validation(self, category="OPEX") -> None:

        if self.validation_running:

            return

        self.selected_tableau_sources = {
            "opex": self.opex_tableau_path_var.get().strip(),
            "direct_cost": self.direct_cost_tableau_path_var.get().strip(),
            "revenue": self.revenue_tableau_path_var.get().strip(),
        }

        # Save the user's Tableau settings for this validation run.
        try:
            self.selected_tableau_settings = self._get_tableau_settings()
        except ValueError as exc:
            messagebox.showerror(
                "Tableau Settings",
                friendly(exc),
                parent=self.winfo_toplevel(),
            )
            return
        selected_key = {
            "OPEX": "opex",
            "Direct Cost": "direct_cost",
            "Revenue": "revenue",
        }[category]
        if (
            not Path(self.selected_tableau_sources[selected_key]).is_file()
            or Path(self.selected_tableau_sources[selected_key]).suffix.lower()
            != ".twb"
        ):
            messagebox.showerror(
                "Source files",
                f"Select a valid {category} consolidated TWB file.",
                parent=self.winfo_toplevel(),
            )
            return
        try:

            # -------------------------------------------------
            # CLOSED PERIOD
            # -------------------------------------------------

            closed_period = self._get_closed_period()

            if (
                self.validation_output_paths is None
                or self.output_period_var.get()
                != self._get_output_period(closed_period)
            ):

                self.validation_output_paths = None
                self._confirm_closed_period()
                if self.validation_output_paths is None:
                    return

            # -------------------------------------------------
            # FINANCIAL STATEMENTS
            # -------------------------------------------------

            aliv_fs_path = Path(self.aliv_fs_path_var.get().strip())

            cbl_fs_path = Path(self.cbl_fs_path_var.get().strip())

            if not str(aliv_fs_path).strip():

                raise ValueError("Please select the ALIV Financial Statement.")

            if not str(cbl_fs_path).strip():

                raise ValueError("Please select the CBL Financial Statement.")

            if not aliv_fs_path.exists():

                raise ValueError(
                    "The selected ALIV Financial Statement " "could not be found."
                )

            if not cbl_fs_path.exists():

                raise ValueError(
                    "The selected CBL Financial Statement " "could not be found."
                )

            # -------------------------------------------------
            # VALIDATION WORKSPACE
            # -------------------------------------------------

            if self.validation_output_paths is None:

                raise ValueError(
                    "Please click Set Validation Period before "
                    "running Actuals Validation."
                )

            # -------------------------------------------------
            # LOCK GUI
            # -------------------------------------------------

            if not confirm_close_tableau(self):
                return

            self.validation_running = True
            self.active_category = category

            self.run_validation_button.configure(state="disabled")
            self.run_direct_cost_button.configure(state="disabled")
            self.run_revenue_button.configure(state="disabled")

            self.open_results_button.configure(
                state="disabled",
            )

            {
                "OPEX": self.opex_status_var,
                "Direct Cost": self.direct_cost_status_var,
                "Revenue": self.revenue_status_var,
            }[category].set("RUNNING")

            self.overall_validation_status_var.set("RUNNING")

            self.validation_progress_var.set("Starting Actuals Validation...")

            self.validation_progress_bar.start(10)

            # -------------------------------------------------
            # BACKGROUND THREAD
            # -------------------------------------------------

            validation_thread = threading.Thread(
                target=self._run_actuals_validation_worker,
                kwargs={
                    "category": category,
                    "closed_period": closed_period,
                    "aliv_fs_path": aliv_fs_path,
                    "cbl_fs_path": cbl_fs_path,
                },
                daemon=True,
            )

            validation_thread.start()

        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)

            messagebox.showerror(
                "Actuals Validation",
                friendly(exc),
                parent=self.winfo_toplevel(),
            )

    def _start_individual_opex_only(self) -> None:
        """Capture individual OPEX into an EXISTING report; never rerun consolidated."""
        if self.validation_running:
            return
        report_name = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="Select EXISTING OPEX validation workbook (not the backup)",
            filetypes=[("Excel workbooks", "*.xlsx")],
        )
        if not report_name:
            return
        report = Path(report_name)
        if not report.is_file():
            messagebox.showerror(
                "OPEX", "Validation workbook not found.", parent=self.winfo_toplevel()
            )
            return
        try:
            closed_period = self._get_closed_period()
            settings = self._get_tableau_settings()
            company = str(settings.get("company", "ALIV")).upper()
            if company not in ("ALIV", "CBL", "BOTH"):
                raise ValueError("Select ALIV, CBL, or Both.")
            aliv = Path(self.aliv_fs_path_var.get().strip())
            cbl_text = self.cbl_fs_path_var.get().strip()
            cbl = Path(cbl_text) if cbl_text else None
            if company in ("ALIV", "BOTH") and not aliv.is_file():
                raise ValueError("Select the ALIV Financial Statement.")
            if company in ("CBL", "BOTH") and (cbl is None or not cbl.is_file()):
                raise ValueError("Select the CBL Financial Statement.")
            if not messagebox.askyesno(
                "Confirm OPEX-only test",
                "Run individual OPEX only?\n\n"
                "Consolidated OPEX, Direct Cost and Revenue will NOT run.\n"
                "Existing individual rows for selected departments may be updated.\n\n"
                f"Report: {report.name}\nCompany: {company}",
                parent=self.winfo_toplevel(),
            ):
                return
            if not confirm_close_tableau(self):
                return
            from shutil import copy2
            from datetime import datetime

            backup = report.with_name(
                report.stem
                + "_before_opex_gui_"
                + datetime.now().strftime("%Y%m%d_%H%M%S")
                + report.suffix
            )
            copy2(report, backup)
            print(f"[OPEX ONLY] Report backup: {backup}", flush=True)
            self.validation_running = True
            self.run_validation_button.configure(state="disabled")

            self.open_results_button.configure(state="disabled")
            self.opex_status_var.set("RUNNING")
            self.direct_cost_status_var.set("NOT RUN")
            self.revenue_status_var.set("NOT RUN")
            self.overall_validation_status_var.set("RUNNING")
            self.validation_progress_var.set("Capturing individual OPEX only...")
            self.validation_progress_bar.start(10)
            self.validation_month_folder = report.parent
            threading.Thread(
                target=self._individual_opex_only_worker,
                kwargs=dict(
                    report=report,
                    aliv=aliv,
                    cbl=cbl,
                    period=closed_period,
                    settings=settings,
                ),
                daemon=True,
            ).start()
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            messagebox.showerror(
                "OPEX-only test", friendly(exc), parent=self.winfo_toplevel()
            )

    def _individual_opex_only_worker(self, *, report, aliv, cbl, period, settings):
        try:
            from automations.actuals.individual_actuals_validation_service import (
                run_individual_opex_pilot,
            )
            from automations.actuals.actuals_report_layout import report_status

            raw = report.parent / "Tableau_Raw_Copies"
            raw.mkdir(parents=True, exist_ok=True)
            result = run_individual_opex_pilot(
                report_path=report,
                aliv_fs_path=aliv,
                cbl_fs_path=cbl,
                output_paths={"tableau_raw_copies": str(raw)},
                closed_period=period,
                tableau_settings=settings,
                include_all_mapped=True,
            )
            status = report_status(report)
            self.after(0, self._finish_individual_opex_only, status, result)
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, self._actuals_validation_failed, explain(exc))

    def _finish_individual_opex_only(self, status, result):
        self.validation_progress_bar.stop()
        self.validation_running = False
        self.run_validation_button.configure(state="normal")

        self.open_results_button.configure(state="normal")
        self.opex_status_var.set(status)
        self.direct_cost_status_var.set("NOT RUN")
        self.revenue_status_var.set("NOT RUN")
        self.overall_validation_status_var.set(status)
        self.overall_validation_status_label.configure(
            bootstyle=self._get_status_bootstyle(status)
        )
        self.validation_progress_var.set(
            f"Individual OPEX finished: {result['status']}. "
            "Consolidated, Direct Cost and Revenue were not run."
        )
        notify_finished(
            self,
            "Individual OPEX finished",
            f"Individual OPEX is complete.\n\nOverall result: {status}",
        )

    # =========================================================
    # REVENUE LIVE CAPTURE (DIAGNOSTIC ONLY)
    # =========================================================
    def _start_revenue_live_test(self):
        import threading
        from pathlib import Path
        from tkinter import messagebox

        if self.validation_running:
            return
        selected = self.revenue_tableau_path_var.get().strip()
        source = Path(selected)
        if not source.is_file() or source.suffix.lower() != ".twb":
            messagebox.showerror(
                "Revenue source",
                "Select a valid consolidated Revenue .twb file.",
                parent=self.winfo_toplevel(),
            )
            return
        try:
            period = self._get_closed_period()
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            messagebox.showerror(
                "Reporting period", friendly(exc), parent=self.winfo_toplevel()
            )
            return
        self.validation_running = True
        self.run_revenue_button.configure(state="disabled")
        self.run_validation_button.configure(state="disabled")
        self.run_direct_cost_button.configure(state="disabled")
        self.revenue_status_var.set("RUNNING")
        self.validation_progress_var.set(
            "Capturing consolidated Revenue MONTH and YTD..."
        )
        threading.Thread(
            target=self._revenue_live_test_worker, args=(source, period), daemon=True
        ).start()

    def _revenue_live_test_worker(self, source, period):
        import csv
        import time
        from io import StringIO
        from pathlib import Path
        from uuid import uuid4
        from automations.actuals.revenue_validation_service import prepare_revenue_copy
        from automations.forecast.tableau_service import open_tableau_workbook
        from automations.actuals.tableau_desktop_export_service import (
            get_tableau_crosstab_text,
        )
        from automations.actuals.tableau_desktop_close_service import (
            close_tableau_forecast_workbook,
        )

        worksheet = "CBL FS Rollup View Group Monthly Revenue YTD_SUMMARY"
        from shared.settings import user_dir

        folder = user_dir() / "Revenue_Live_Test_Results" / f"{period:%Y-%m}_GUI"
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for basis in ("MONTH", "YTD"):
                # Work directly on the source workbook: set the filters, save it.
                temporary = Path(source)
                opened = False
                try:
                    prepare_revenue_copy(source, temporary, [worksheet], period, basis)
                    open_tableau_workbook(tableau_workbook_path=str(temporary))
                    opened = True
                    time.sleep(8)
                    data = get_tableau_crosstab_text(
                        worksheet_name=worksheet, tableau_workbook_path=str(temporary)
                    )
                    if not data.strip():
                        raise ValueError(f"{basis} capture returned empty data")
                    rows = list(csv.reader(StringIO(data), delimiter="\t"))
                    (folder / f"Revenue_Group_{basis}.tsv").write_text(
                        data, encoding="utf-8-sig"
                    )
                    print(
                        f"[REVENUE GUI] {basis}: {len(rows)} rows; saved to {folder}",
                        flush=True,
                    )
                finally:
                    if opened:
                        try:
                            close_tableau_forecast_workbook(
                                tableau_workbook_path=str(temporary)
                            )
                        except Exception as exc:
                            log.exception("Error in %s", type(self).__name__)
                            print(f"[REVENUE GUI] Close warning: {exc}", flush=True)
            self.after(0, self._revenue_live_test_finished, str(folder), None)
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, self._revenue_live_test_finished, str(folder), explain(exc))

    def _revenue_live_test_finished(self, folder, error):
        from tkinter import messagebox

        self.validation_running = False
        self.run_revenue_button.configure(state="normal")
        self.run_validation_button.configure(state="normal")
        self.run_direct_cost_button.configure(state="normal")
        if error:
            self.revenue_status_var.set("ERROR")
            self.validation_progress_var.set(f"Revenue capture failed: {error}")
            messagebox.showerror(
                "Revenue live test", error, parent=self.winfo_toplevel()
            )
        else:
            self.revenue_status_var.set("REVIEW REQUIRED")
            self.validation_progress_var.set(f"Revenue raw captures saved: {folder}")
            messagebox.showinfo(
                "Revenue capture completed",
                f"MONTH and YTD raw data saved to:\n{folder}\n\nAmounts are NOT yet reconciled.",
                parent=self.winfo_toplevel(),
            )

    # =========================================================
    # BACKGROUND VALIDATION WORKER
    # =========================================================

    def _run_actuals_validation_worker(
        self,
        *,
        category: str,
        closed_period: date,
        aliv_fs_path: Path,
        cbl_fs_path: Path,
    ) -> None:
        try:
            self.after(
                0,
                self._set_validation_progress,
                category,
                "RUNNING",
                f"Validating {category} consolidated Tableau against statements...",
            )
            if category == "OPEX":
                result = run_live_opex_validation(
                    tableau_workbook_path=self.selected_tableau_sources["opex"],
                    aliv_fs_path=aliv_fs_path,
                    cbl_fs_path=cbl_fs_path,
                    output_paths=self.validation_output_paths,
                    closed_period=closed_period,
                    open_tableau_first=True,
                    run_individual_validation=False,
                )
                self.after(
                    0,
                    self._set_validation_progress,
                    category,
                    "RUNNING",
                    "Running individual OPEX Tableau validation...",
                )
                from automations.actuals.individual_actuals_validation_service import (
                    run_individual_opex_pilot,
                )

                individual = run_individual_opex_pilot(
                    report_path=result["validation_workbook"],
                    aliv_fs_path=aliv_fs_path,
                    cbl_fs_path=cbl_fs_path,
                    output_paths=self.validation_output_paths,
                    closed_period=closed_period,
                    tableau_settings=self.selected_tableau_settings,
                )
                result["individual_actuals"] = individual
                from automations.actuals.actuals_report_layout import report_status

                result["overall_status"] = report_status(result["validation_workbook"])
            elif category == "Direct Cost":
                result = run_live_direct_cost_validation(
                    tableau_workbook_path=self.selected_tableau_sources["direct_cost"],
                    aliv_fs_path=aliv_fs_path,
                    cbl_fs_path=cbl_fs_path,
                    output_paths=self.validation_output_paths,
                    closed_period=closed_period,
                    open_tableau_first=True,
                )
                # Run all mapped individual Direct Cost TWBs in one live batch.
                self.after(
                    0,
                    self._set_validation_progress,
                    category,
                    "RUNNING",
                    "Capturing individual Direct Cost worksheets to Tableau Raw Copies...",
                )
                from automations.actuals.individual_direct_cost_live_service import (
                    run_individual_direct_cost_live,
                )

                individual = run_individual_direct_cost_live(
                    output_paths=self.validation_output_paths,
                    closed_period=closed_period,
                    tableau_settings=self.selected_tableau_settings,
                )
                result["individual_actuals"] = individual
                # Raw captures are not financial reconciliation results.
                result["overall_status"] = "REVIEW REQUIRED"
                self.after(
                    0,
                    self._set_validation_progress,
                    category,
                    "REVIEW REQUIRED",
                    f"Captured {individual['captured']}/{individual['total']} individual views; see Tableau Raw Copies.",
                )
            elif category == "Revenue":
                from automations.actuals.revenue_validation_service import (
                    run_live_revenue_validation,
                )

                group_path = Path(self.selected_tableau_sources["revenue"])
                revenue_sources = {
                    "group": str(group_path),
                    "consumer": str(
                        group_path.parent / "Cable Bahamas Group_CONSUMER REVENUE.twb"
                    ),
                    "corporate": str(
                        group_path.parent / "Cable Bahamas Group_CORPORATE REVENUE.twb"
                    ),
                    "finance": str(
                        group_path.parent / "CBL Group Revenue_Finance View.twb"
                    ),
                    "media": str(
                        group_path.parent / "Cable Bahamas Group_MEDIA REVENUE.twb"
                    ),
                }
                result = run_live_revenue_validation(
                    aliv_fs_path=aliv_fs_path,
                    cbl_fs_path=cbl_fs_path,
                    tableau_sources=revenue_sources,
                    output_paths=self.validation_output_paths,
                    closed_period=closed_period,
                )
                result["overall_status"] = "REVIEW REQUIRED"
            else:
                raise ValueError(f"Unsupported validation category: {category}")

            # Format the SAME validation workbook in place, like OPEX.
            # No second *_Final_AO.xlsx output is created.
            from automations.actuals.actuals_report_layout import (
                finalize_actuals_report_in_place,
            )

            source = Path(result["validation_workbook"])
            # Revenue reports the values it actually set in Tableau.
            result["final_report"] = finalize_actuals_report_in_place(
                source,
                forecast_type=result.get("forecast_type")
                or self.selected_tableau_settings.get("forecast_type"),
                forecast_start=result.get("forecast_start")
                or self.selected_tableau_settings.get("forecast_start"),
            )
            if category == "Revenue":
                from automations.actuals.report_builder import add_run_notes_sheet

                add_run_notes_sheet(
                    result["final_report"],
                    filters=result.get("filters_used", []),
                    notes=result.get("capture_notes", []),
                )
            if category == "Direct Cost":
                from automations.actuals.direct_cost_individual_merge import (
                    merge_direct_cost_crosstabs,
                )

                merge_result = merge_direct_cost_crosstabs(
                    report_path=result["final_report"],
                    crosstab_path=individual["crosstab_workbook"],
                    closed_period=closed_period,
                )
                result["individual_merge"] = merge_result

            # The final report is the single source of truth for the status.
            from automations.actuals.actuals_report_layout import report_status

            result["overall_status"] = report_status(result["final_report"])
            self.after(0, self._finish_actuals_validation, {category: result})
        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            self.after(0, self._actuals_validation_failed, explain(exc))

    # =========================================================
    # UPDATE VALIDATION STATUS
    # =========================================================

    def _set_validation_progress(
        self,
        area: str,
        status: str,
        message: str,
    ) -> None:
        """
        Update one validation area and the progress message.
        """

        status_var = None
        status_label = None

        if area == "OPEX":

            status_var = self.opex_status_var
            status_label = self.opex_status_label

        elif area == "Direct Cost":

            status_var = self.direct_cost_status_var
            status_label = self.direct_cost_status_label

        elif area == "Revenue":

            status_var = self.revenue_status_var
            status_label = self.revenue_status_label

        if status_var is not None:

            status_var.set(status)

        if status_label is not None:

            status_label.configure(bootstyle=self._get_status_bootstyle(status))

        self.validation_progress_var.set(message)

    # =========================================================
    # STATUS STYLE
    # =========================================================

    def _get_status_bootstyle(
        self,
        status: str,
    ) -> str:
        """
        Convert validation status into a ttkbootstrap style.
        """

        normalized_status = str(status).strip().upper()

        if normalized_status == "PASS":

            return "success"

        if normalized_status in {
            "FAIL",
            "REVIEW REQUIRED",
        }:

            return "danger"

        if normalized_status in {
            "RUNNING",
            "PENDING",
        }:

            return "warning"

        return "secondary"

    # =========================================================
    # FINISH VALIDATION
    # =========================================================

    def _finish_actuals_validation(
        self,
        results: dict,
    ) -> None:
        """
        Finish the current validation run and update the GUI.
        """

        self.validation_progress_bar.stop()

        self.validation_running = False

        self.run_validation_button.configure(
            state="normal",
        )

        self.run_direct_cost_button.configure(state="normal")
        self.run_revenue_button.configure(state="normal")

        self.open_results_button.configure(
            state="normal",
        )

        category, result = next(iter(results.items()))
        overall = result.get("overall_status", "PENDING")
        self.overall_validation_status_var.set(overall)
        self.overall_validation_status_label.configure(
            bootstyle=self._get_status_bootstyle(overall)
        )
        self._set_validation_progress(
            category, overall, f"{category} completed. Review report and exceptions."
        )
        report = result.get("final_report")
        self.validation_progress_var.set(
            f"{category}: {'Final report: ' + report if report else 'Final report was not generated.'}"
        )
        notify_finished(
            self,
            f"{category} validation finished",
            f"{category} validation is complete.\n\nOverall result: {overall}",
            open_path=report,
        )

    # =========================================================
    # VALIDATION FAILED
    # =========================================================

    def _actuals_validation_failed(
        self,
        error_message: str,
    ) -> None:
        """
        Restore the GUI after a validation engine error.
        """

        self.validation_progress_bar.stop()

        self.validation_running = False

        self.run_validation_button.configure(
            state="normal",
        )

        self.run_direct_cost_button.configure(state="normal")
        self.run_revenue_button.configure(state="normal")

        self.overall_validation_status_var.set("ERROR")

        self.overall_validation_status_label.configure(
            bootstyle="danger",
        )

        self.validation_progress_var.set(
            "Actuals Validation stopped because an error occurred."
        )

        notify_finished(
            self,
            "Actuals Validation Error",
            f"The validation stopped because of an error:\n\n{error_message}",
            success=False,
        )

    # =========================================================
    # OPEN RESULTS FOLDER
    # =========================================================

    def _open_validation_results(
        self,
    ) -> None:
        """
        Open the resolved monthly validation results folder.
        """

        if self.validation_month_folder is None:

            messagebox.showwarning(
                "Actuals Validation",
                "No validation results folder has been created.",
                parent=self.winfo_toplevel(),
            )

            return

        results_folder = Path(self.validation_month_folder)

        if not results_folder.exists():

            messagebox.showerror(
                "Actuals Validation",
                "The validation results folder could not be found.",
                parent=self.winfo_toplevel(),
            )

            return

        os.startfile(str(results_folder))

    def _get_tableau_settings(self):

        def parse_pairs(value):
            result = {}

            for item in value.split(";"):
                item = item.strip()

                if not item:
                    continue

                if "=" not in item:
                    raise ValueError(
                        f"Invalid Tableau setting: {item}. " "Use Name=Value."
                    )

                name, setting_value = item.split("=", 1)
                name = name.strip()
                setting_value = setting_value.strip()

                if not name or not setting_value:
                    raise ValueError(f"Invalid Tableau setting: {item}")

                result[name] = setting_value

            return result

        from datetime import date

        start_date = date.fromisoformat(self.forecast_start_var.get().strip())

        return {
            "forecast_type": self.forecast_type_var.get().strip(),
            "forecast_start": start_date,
            "company": self.company_selection_var.get(),
            "filters": parse_pairs(self.additional_filters_var.get()),
            "parameters": parse_pairs(self.additional_parameters_var.get()),
        }
