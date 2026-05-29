from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from bex_payroll_import.config import AppConfig, load_config, save_config
from bex_payroll_import.models import RunInputs
from bex_payroll_import.runner import run_payroll_import


class PayrollImportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BEX Payroll Import")
        self.geometry("720x360")
        self.resizable(False, False)

        self.config_state = load_config()
        self.template_path = tk.StringVar(value=str(self.config_state.template_path or ""))
        self.tdr_path = tk.StringVar(value="")
        self.expense_path = tk.StringVar(value="")
        self.expense_skipped = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Select files to generate a payroll import.")

        self.build()
        self.refresh_generate_state()
        if not self.config_state.template_path or not self.config_state.template_path.exists():
            self.choose_template()

    def build(self) -> None:
        padding = {"padx": 12, "pady": 6}

        tk.Label(self, text="Payroll Upload Worksheet template").grid(row=0, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.template_path, width=70, state="readonly").grid(row=0, column=1, **padding)
        tk.Button(self, text="Change", command=self.choose_template).grid(row=0, column=2, **padding)

        tk.Label(self, text="Time Detail Report").grid(row=1, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.tdr_path, width=70, state="readonly").grid(row=1, column=1, **padding)
        tk.Button(self, text="Select", command=self.choose_tdr).grid(row=1, column=2, **padding)

        tk.Label(self, text="Expense Transaction").grid(row=2, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.expense_path, width=70, state="readonly").grid(row=2, column=1, **padding)
        tk.Button(self, text="Select", command=self.choose_expense).grid(row=2, column=2, **padding)

        tk.Checkbutton(
            self,
            text="I confirm there is no Expense Transaction file for this payroll import.",
            variable=self.expense_skipped,
            command=self.on_skip_expense_changed,
        ).grid(row=3, column=1, sticky="w", **padding)

        self.generate_button = tk.Button(self, text="Generate Payroll Import", command=self.generate)
        self.generate_button.grid(row=4, column=1, sticky="e", **padding)

        tk.Label(self, textvariable=self.status_text, wraplength=640, justify="left").grid(
            row=5,
            column=0,
            columnspan=3,
            sticky="w",
            **padding,
        )

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(
            title="Select Payroll Upload Worksheet",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm")],
        )
        if path:
            template = Path(path)
            self.template_path.set(str(template))
            save_config(AppConfig(template_path=template))
            self.refresh_generate_state()

    def choose_tdr(self) -> None:
        path = filedialog.askopenfilename(title="Select Time Detail Report", filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.tdr_path.set(path)
            self.refresh_generate_state()

    def choose_expense(self) -> None:
        path = filedialog.askopenfilename(title="Select Expense Transaction", filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.expense_path.set(path)
            self.expense_skipped.set(False)
            self.refresh_generate_state()

    def on_skip_expense_changed(self) -> None:
        if self.expense_skipped.get():
            self.expense_path.set("")
        self.refresh_generate_state()

    def refresh_generate_state(self) -> None:
        ready = bool(
            self.template_path.get()
            and self.tdr_path.get()
            and (self.expense_path.get() or self.expense_skipped.get())
        )
        self.generate_button.config(state=tk.NORMAL if ready else tk.DISABLED)

    def generate(self) -> None:
        try:
            outputs = run_payroll_import(
                RunInputs(
                    template_path=Path(self.template_path.get()),
                    tdr_path=Path(self.tdr_path.get()),
                    expense_path=Path(self.expense_path.get()) if self.expense_path.get() else None,
                    expense_skipped=self.expense_skipped.get(),
                )
            )
            if outputs.payroll_csv_path:
                self.status_text.set(f"Payroll import created: {outputs.payroll_csv_path}")
                messagebox.showinfo("Payroll import created", f"Output folder:\n{outputs.output_dir}")
            else:
                self.status_text.set(f"Validation stopped the run: {outputs.validation_path}")
                messagebox.showerror(
                    "Validation stopped the run",
                    f"Open this file for details:\n{outputs.validation_path}",
                )
            open_folder(outputs.output_dir)
        except Exception as exc:
            messagebox.showerror("Payroll import failed", str(exc))
            self.status_text.set(f"Payroll import failed: {exc}")


def open_folder(path: Path) -> None:
    if os.name == "nt":
        os.startfile(path)


def main() -> None:
    app = PayrollImportApp()
    app.mainloop()
