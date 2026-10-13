from __future__ import annotations

import logging

import os
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from automations.budget_setup.service import create_initial_forecasts
from shared import settings


log = logging.getLogger(__name__)

class BudgetSetupView(tb.Frame):
    def __init__(self, parent) -> None:
        super().__init__(parent)

        self._build_variables()
        self._build_workspace()

    def _build_variables(self) -> None:
        self.source_fy_var = tb.StringVar(value="FY26")
        self.target_fy_var = tb.StringVar(value="FY27")

        self.source_folder_var = tb.StringVar(value=str(settings.get_path("budget_source_folder") or ""))
        self.output_folder_var = tb.StringVar(value=str(settings.get_path("budget_output_folder") or ""))

        self.overwrite_var = tb.StringVar(value="skip")

        self.status_var = tb.StringVar(value="Ready")
        self.status_detail_var = tb.StringVar(
            value=(
                "Select the budget source and forecast output folders, "
                "then create the initial forecast workbooks."
            )
        )

        self.files_created_var = tb.StringVar(value="0")
        self.files_skipped_var = tb.StringVar(value="0")
        self.files_failed_var = tb.StringVar(value="0")
        self.sheets_updated_var = tb.StringVar(value="0")
        self.output_location_var = tb.StringVar(value="Not created yet")

    def _build_workspace(self) -> None:
        title_frame = tb.Frame(self)
        title_frame.pack(fill="x", pady=(0, 14))

        tb.Label(
            title_frame,
            text="Initial Forecast Month Setup",
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w")

        tb.Label(
            title_frame,
            text=(
                "Create the initial fiscal-year forecast workbooks "
                "from approved budget submissions."
            ),
            font=("Segoe UI", 10),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(3, 0))

        tb.Label(
            title_frame,
            text=(
                "Use this process once when establishing "
                "the new fiscal-year forecast workbooks."
            ),
            font=("Segoe UI", 10, "bold"),
            bootstyle="warning",
        ).pack(anchor="w", pady=(14, 0))

        cards = tb.Frame(self)
        cards.pack(fill="both", expand=True)

        self._build_setup_card(cards)
        self._build_status_card(cards)

    def _build_setup_card(self, parent) -> None:
        card = tb.Labelframe(
            parent,
            text=" Initial Forecast Month Setup ",
            padding=22,
            bootstyle="primary",
        )
        card.pack(
            side="left",
            fill="both",
            expand=True,
            padx=(0, 10),
        )

        form = tb.Frame(card)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        self._add_row_label(
            form,
            "Source fiscal year",
            0,
        )

        tb.Combobox(
            form,
            textvariable=self.source_fy_var,
            values=["FY25", "FY26", "FY27", "FY28"],
            state="readonly",
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            pady=8,
        )

        self._add_row_label(
            form,
            "Target fiscal year",
            1,
        )

        tb.Combobox(
            form,
            textvariable=self.target_fy_var,
            values=["FY26", "FY27", "FY28", "FY29"],
            state="readonly",
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            pady=8,
        )

        self._folder_row(
            parent=form,
            row=2,
            label="Budget source folder",
            variable=self.source_folder_var,
            command=self._browse_source,
        )

        self._folder_row(
            parent=form,
            row=3,
            label="Forecast output folder",
            variable=self.output_folder_var,
            command=self._browse_output,
        )

        self._add_row_label(
            form,
            "Existing files",
            4,
        )

        tb.Combobox(
            form,
            textvariable=self.overwrite_var,
            values=[
                "skip",
                "replace",
                "cancel",
            ],
            state="readonly",
        ).grid(
            row=4,
            column=1,
            sticky="ew",
            pady=8,
        )

        tb.Separator(card).pack(
            fill="x",
            pady=(18, 14),
        )

        tb.Label(
            card,
            text="Initial forecast month: July",
            font=("Segoe UI", 11, "bold"),
            bootstyle="info",
        ).pack(anchor="w")

        tb.Label(
            card,
            text=(
                "The approved monthly budget values are retained "
                "for July through June. The fiscal month headers "
                "are rolled forward to the new fiscal year, while "
                "formulas, formatting, and workbook structure remain intact."
            ),
            font=("Segoe UI", 9),
            bootstyle="secondary",
            wraplength=650,
            justify="left",
        ).pack(
            anchor="w",
            pady=(6, 16),
        )

        self.run_button = tb.Button(
            card,
            text="Create Initial Forecast Workbooks",
            bootstyle="success",
            command=self._start_processing,
            padding=(14, 10),
        )
        self.run_button.pack(
            fill="x",
            pady=(10, 6),
        )

    def _build_status_card(self, parent) -> None:
        card = tb.Labelframe(
            parent,
            text=" Process Status ",
            padding=22,
            bootstyle="secondary",
            width=470,
        )
        card.pack(
            side="right",
            fill="both",
            padx=(10, 0),
        )
        card.pack_propagate(False)

        tb.Label(
            card,
            textvariable=self.status_var,
            font=("Segoe UI", 20, "bold"),
            bootstyle="primary",
        ).pack(anchor="w")

        tb.Label(
            card,
            textvariable=self.status_detail_var,
            font=("Segoe UI", 10),
            bootstyle="secondary",
            wraplength=420,
            justify="left",
        ).pack(
            anchor="w",
            pady=(5, 18),
        )

        self.progress = tb.Progressbar(
            card,
            mode="indeterminate",
            bootstyle="info-striped",
        )
        self.progress.pack(
            fill="x",
            pady=(0, 22),
        )

        metrics = tb.Frame(card)
        metrics.pack(fill="x")

        self._metric(
            metrics,
            "Files created",
            self.files_created_var,
        )
        self._metric(
            metrics,
            "Files skipped",
            self.files_skipped_var,
        )
        self._metric(
            metrics,
            "Files failed",
            self.files_failed_var,
        )

        tb.Separator(card).pack(
            fill="x",
            pady=22,
        )

        tb.Label(
            card,
            text="Sheets updated",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        tb.Label(
            card,
            textvariable=self.sheets_updated_var,
            font=("Segoe UI", 20, "bold"),
        ).pack(
            anchor="w",
            pady=(4, 16),
        )

        tb.Label(
            card,
            text="Forecast files location",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        tb.Label(
            card,
            textvariable=self.output_location_var,
            font=("Segoe UI", 9),
            bootstyle="secondary",
            wraplength=420,
            justify="left",
        ).pack(
            anchor="w",
            pady=(5, 18),
        )

        folder_buttons = tb.Frame(card)
        folder_buttons.pack(fill="x")

        self.open_output_button = tb.Button(
            folder_buttons,
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

        tb.Button(
            folder_buttons,
            text="Reset",
            bootstyle="secondary-outline",
            command=self._reset_status,
        ).pack(
            side="right",
            fill="x",
            expand=True,
            padx=(6, 0),
        )

        # tb.Separator(card).pack(
        #     fill="x",
        #     pady=22,
        # )

        # tb.Label(
        #     card,
        #     text="Next step",
        #     font=("Segoe UI", 10, "bold"),
        # ).pack(anchor="w")

        # tb.Label(
        #     card,
        #     text=(
        #         "After reviewing the newly created workbooks, "
        #         "open Tableau Prep and manually run the flow."
        #     ),
        #     font=("Segoe UI", 9),
        #     bootstyle="secondary",
        #     wraplength=420,
        #     justify="left",
        # ).pack(
        #     anchor="w",
        #     pady=(5, 12),
        # )

        # self.open_prep_button = tb.Button(
        #     card,
        #     text="Open Tableau Prep Flow",
        #     bootstyle="info",
        #     command=self._confirm_and_open_prep,
        #     state="disabled",
        #     padding=(12, 9),
        # )
        # self.open_prep_button.pack(fill="x")

    def _start_processing(self) -> None:
        if not self._validate_inputs():
            return

        confirmed = messagebox.askyesno(
            "Confirm initial forecast setup",
            (
                f"Source fiscal year: {self.source_fy_var.get()}\n"
                f"Target fiscal year: {self.target_fy_var.get()}\n"
                f"Existing files: {self.overwrite_var.get()}\n\n"
                "The approved budget files will be copied and "
                "converted into initial forecast workbooks.\n\n"
                "Continue?"
            ),
        )

        if not confirmed:
            return

        self.run_button.configure(state="disabled")
        self.open_output_button.configure(state="disabled")

        self.status_var.set("Processing")
        self.status_detail_var.set(
            "Copying budget workbooks and preparing " "the initial forecast files."
        )

        self.files_created_var.set("0")
        self.files_skipped_var.set("0")
        self.files_failed_var.set("0")
        self.sheets_updated_var.set("0")
        self.output_location_var.set("Processing...")

        self.progress.start(12)

        threading.Thread(
            target=self._run_processing,
            daemon=True,
        ).start()

    def _run_processing(self) -> None:
        try:
            result = create_initial_forecasts(
                source_folder=self.source_folder_var.get(),
                output_folder=self.output_folder_var.get(),
                source_fiscal_year=self.source_fy_var.get(),
                target_fiscal_year=self.target_fy_var.get(),
                overwrite_mode=self.overwrite_var.get(),
            )

            self.after(
                0,
                lambda value=result: self._processing_complete(value),
            )

        except Exception as exc:
            log.exception("Error in %s", type(self).__name__)
            error_message = f"{type(exc).__name__}: {exc}"

            self.after(
                0,
                lambda value=error_message: self._processing_failed(value),
            )

    def _processing_complete(self, result: dict) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")

        created = result.get("files_created", 0)
        skipped = result.get("files_skipped", 0)
        failed = result.get("files_failed", 0)
        failures = result.get("failures", [])
        sheets = result.get(
            "worksheets_updated",
            result.get("worksheets_processed", 0),
        )

        self.files_created_var.set(str(created))
        self.files_skipped_var.set(str(skipped))
        self.files_failed_var.set(str(failed))
        self.sheets_updated_var.set(str(sheets))
        self.output_location_var.set(
            result.get(
                "output_folder",
                self.output_folder_var.get(),
            )
        )

        if failed:
            self.status_var.set("Completed with warnings")
            self.status_detail_var.set(
                f"{failed} file(s) could not be processed. " "Review the error details."
            )

            failure_lines = []

            for failure in failures:
                file_name = failure.get(
                    "file",
                    "Unknown file",
                )
                error_message = failure.get(
                    "error",
                    "Unknown error",
                )

                failure_lines.append(f"{file_name}\n" f"Reason: {error_message}")

            failure_text = "\n\n".join(failure_lines)

            messagebox.showwarning(
                "Files failed",
                (f"{failed} file(s) failed:\n\n" f"{failure_text}"),
            )
        else:
            self.status_var.set("Initial Forecast Month Setup completed")
            self.status_detail_var.set(
                "Review the newly created forecast workbooks"
                "and confirm that all monthly budget values were retained."
            )

        self.open_output_button.configure(state="normal")

    def _processing_failed(
        self,
        error_message: str,
    ) -> None:
        self.progress.stop()
        self.run_button.configure(state="normal")

        self.status_var.set("Failed")
        self.status_detail_var.set(error_message)
        self.output_location_var.set("No output created")

        messagebox.showerror(
            "Initial Forecast Month Setupfailed",
            error_message,
        )

    def _validate_inputs(self) -> bool:
        source = Path(self.source_folder_var.get())
        output = Path(self.output_folder_var.get())

        if not source.exists():
            messagebox.showerror(
                "Source folder unavailable",
                ("The budget source folder " f"could not be found:\n\n{source}"),
            )
            return False

        try:
            output.mkdir(
                parents=True,
                exist_ok=True,
            )
        except OSError as exc:
            messagebox.showerror(
                "Output folder unavailable",
                str(exc),
            )
            return False

        matching_files = [
            file
            for file in source.glob("*.xls*")
            if not file.name.startswith("~$")
            and self.source_fy_var.get().upper() in file.name.upper()
            and "BUDGET" in file.name.upper()
        ]

        if not matching_files:
            messagebox.showwarning(
                "No matching budget files",
                (f"No {self.source_fy_var.get()} " "budget workbooks were found."),
            )
            return False

        return True

    def _open_output_folder(self) -> None:
        folder = Path(self.output_location_var.get())

        if folder.exists():
            os.startfile(str(folder))
        else:
            messagebox.showwarning(
                "Folder unavailable",
                "The forecast output folder could not be found.",
            )

    # def _confirm_and_open_prep(self) -> None:
    #     flow_path = self.prep_flow_var.get().strip()

    #     if not flow_path:
    #         messagebox.showwarning(
    #             "No Tableau Prep flow selected",
    #             "Select a .tfl or .tflx file first.",
    #         )
    #         return

    #     confirmed = messagebox.askyesno(
    #         "Open Tableau Prep flow",
    #         (
    #             "Before opening Tableau Prep:\n\n"
    #             "• Review the newly created workbooks.\n"
    #             "• Close all Excel files.\n"
    #             "• Confirm network or VPN access.\n\n"
    #             "The flow will open, but it will not run automatically.\n"
    #             "You must click Run Flow manually.\n\n"
    #             "Continue?"
    #         ),
    #     )

    #     if not confirmed:
    #         return

    #     try:
    #         open_prep_flow(flow_path)

    #         self.status_var.set("Tableau Prep opened")
    #         self.status_detail_var.set(
    #             "Review any warnings, then click " "Run Flow manually."
    #         )

    #     except Exception as exc:
    #         messagebox.showerror(
    #             "Unable to open Tableau Prep",
    #             f"{type(exc).__name__}: {exc}",
    #         )

    def _reset_status(self) -> None:
        self.progress.stop()

        self.status_var.set("Ready")
        self.status_detail_var.set(
            "Select the budget source and forecast output "
            "folders, then create the initial forecast workbooks."
        )

        self.files_created_var.set("0")
        self.files_skipped_var.set("0")
        self.files_failed_var.set("0")
        self.sheets_updated_var.set("0")
        self.output_location_var.set("Not created yet")

        self.open_output_button.configure(state="disabled")

    def _browse_source(self) -> None:
        selected = filedialog.askdirectory(
            title="Select approved budget folder",
            initialdir=self.source_folder_var.get(),
        )

        if selected:
            self.source_folder_var.set(selected)

    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(
            title="Select forecast output folder",
            initialdir=self.output_folder_var.get(),
        )

        if selected:
            self.output_folder_var.set(selected)

    # def _browse_prep_flow(self) -> None:
    #     selected = filedialog.askopenfilename(
    #         title="Select Tableau Prep flow",
    #         initialdir=r"T:\TableauPrep",
    #         filetypes=[
    #             (
    #                 "Tableau Prep flows",
    #                 "*.tfl *.tflx",
    #             ),
    #             (
    #                 "All files",
    #                 "*.*",
    #             ),
    #         ],
    #     )

    #     if selected:
    #         self.prep_flow_var.set(selected)

    def _add_row_label(
        self,
        parent,
        text: str,
        row: int,
    ) -> None:
        tb.Label(
            parent,
            text=text,
            font=("Segoe UI", 10, "bold"),
        ).grid(
            row=row,
            column=0,
            sticky="w",
            padx=(0, 18),
        )

    def _folder_row(
        self,
        parent,
        row: int,
        label: str,
        variable: tb.StringVar,
        command,
        button_text: str = "Browse",
    ) -> None:
        self._add_row_label(
            parent,
            label,
            row,
        )

        row_frame = tb.Frame(parent)
        row_frame.grid(
            row=row,
            column=1,
            sticky="ew",
            pady=8,
        )
        row_frame.columnconfigure(0, weight=1)

        tb.Entry(
            row_frame,
            textvariable=variable,
        ).grid(
            row=0,
            column=0,
            sticky="ew",
        )

        tb.Button(
            row_frame,
            text=button_text,
            bootstyle="secondary-outline",
            command=command,
        ).grid(
            row=0,
            column=1,
            padx=(8, 0),
        )

    def _metric(
        self,
        parent,
        label: str,
        variable: tb.StringVar,
    ) -> None:
        frame = tb.Frame(
            parent,
            padding=8,
        )
        frame.pack(
            side="left",
            fill="both",
            expand=True,
        )

        tb.Label(
            frame,
            textvariable=variable,
            font=("Segoe UI", 21, "bold"),
        ).pack()

        tb.Label(
            frame,
            text=label,
            font=("Segoe UI", 9),
            bootstyle="secondary",
        ).pack()
