from __future__ import annotations

import os
from pathlib import Path
from tkinter import filedialog, messagebox

import ttkbootstrap as tb

from shared import settings
from shared.logging_setup import log_dir

# (key, label, help text, picker: "folder" or None)
PATH_FIELDS = (
    (
        "workbooks_root",
        "Tableau workbooks folder",
        "Folder holding the OPEX, SAC, Revenue and Direct Cost workbook folders.",
        "folder",
    ),
    (
        "validation_output_root",
        "Actuals validation output folder",
        "Where validation reports are saved. Leave blank to find the synced OneDrive folder.",
        "folder",
    ),
    (
        "forecast_validation_output_root",
        "Forecast validation output folder",
        "Leave blank to use the synced OneDrive shortcut folder.",
        "folder",
    ),
    (
        "budget_source_folder",
        "Budget source folder",
        "Default source folder on the Initial Budget Setup page.",
        "folder",
    ),
    (
        "budget_output_folder",
        "Budget output folder",
        "Default output folder on the Initial Budget Setup page.",
        "folder",
    ),
)


class SettingsView(tb.Frame):
    def __init__(self, parent) -> None:
        super().__init__(parent)
        current = settings.load()
        self.vars = {key: tb.StringVar(value=str(current[key] or "")) for key, *_ in PATH_FIELDS}
        self.test_mode_var = tb.BooleanVar(value=bool(current["test_mode"]))
        self.log_level_var = tb.StringVar(value=str(current["log_level"]))
        self._build()

    def _build(self) -> None:
        tb.Label(self, text="Settings", font=("Segoe UI", 22, "bold")).pack(anchor="w")
        tb.Label(
            self,
            text=(
                "Your settings are saved for your Windows user only. "
                "Restart the app after changing folders."
            ),
            font=("Segoe UI", 10),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(0, 16))

        form = tb.Labelframe(self, text="Folders", padding=16)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        for row, (key, label, help_text, picker) in enumerate(PATH_FIELDS):
            tb.Label(form, text=label, font=("Segoe UI", 10, "bold")).grid(
                row=row * 2, column=0, sticky="w", padx=(0, 12), pady=(8, 0)
            )
            tb.Entry(form, textvariable=self.vars[key]).grid(
                row=row * 2, column=1, sticky="ew", pady=(8, 0)
            )
            if picker == "folder":
                tb.Button(
                    form,
                    text="Browse",
                    bootstyle="secondary-outline",
                    command=lambda k=key: self._browse(k),
                ).grid(row=row * 2, column=2, padx=(8, 0), pady=(8, 0))
            tb.Label(form, text=help_text, font=("Segoe UI", 9), bootstyle="secondary").grid(
                row=row * 2 + 1, column=1, sticky="w"
            )

        other = tb.Labelframe(self, text="Other", padding=16)
        other.pack(fill="x", pady=(16, 0))
        tb.Checkbutton(
            other,
            text="Show the TEST MODE badge",
            variable=self.test_mode_var,
            bootstyle="round-toggle",
        ).grid(row=0, column=0, sticky="w")
        tb.Label(other, text="Log detail").grid(row=1, column=0, sticky="w", pady=(10, 0))
        tb.Combobox(
            other,
            textvariable=self.log_level_var,
            values=("DEBUG", "INFO", "WARNING", "ERROR"),
            state="readonly",
            width=12,
        ).grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

        actions = tb.Frame(self)
        actions.pack(fill="x", pady=(20, 0))
        tb.Button(actions, text="Save settings", bootstyle="primary", command=self._save).pack(
            side="left"
        )
        tb.Button(
            actions,
            text="Open log folder",
            bootstyle="secondary-outline",
            command=lambda: self._open(log_dir()),
        ).pack(side="left", padx=(10, 0))
        tb.Button(
            actions,
            text="Open settings folder",
            bootstyle="secondary-outline",
            command=lambda: self._open(settings.user_dir()),
        ).pack(side="left", padx=(10, 0))

        tb.Label(
            self,
            text=f"Shared defaults file: {settings.shared_settings_file()}",
            font=("Segoe UI", 9),
            bootstyle="secondary",
        ).pack(anchor="w", pady=(16, 0))

    def _browse(self, key: str) -> None:
        current = self.vars[key].get().strip()
        selected = filedialog.askdirectory(
            parent=self.winfo_toplevel(),
            initialdir=current if current and Path(current).is_dir() else None,
        )
        if selected:
            self.vars[key].set(str(Path(selected)))

    def _save(self) -> None:
        values = {key: var.get().strip() for key, var in self.vars.items()}
        missing = [
            label
            for key, label, *_ in PATH_FIELDS
            if values[key] and not Path(os.path.expandvars(values[key])).is_dir()
        ]
        if missing and not messagebox.askyesno(
            "Folder not found",
            "These folders were not found:\n\n- "
            + "\n- ".join(missing)
            + "\n\nSave anyway? (For example, if the shared drive is not connected right now.)",
            parent=self.winfo_toplevel(),
        ):
            return
        values["test_mode"] = bool(self.test_mode_var.get())
        values["log_level"] = self.log_level_var.get()
        path = settings.save_user_settings(values)
        messagebox.showinfo(
            "Settings saved",
            f"Saved to:\n{path}\n\nRestart the app to apply folder changes.",
            parent=self.winfo_toplevel(),
        )

    def _open(self, folder: Path) -> None:
        folder.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(folder)  # type: ignore[attr-defined]  # Windows only
        except AttributeError:
            messagebox.showinfo("Folder", str(folder), parent=self.winfo_toplevel())
