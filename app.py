from __future__ import annotations

import sys
from pathlib import Path
from typing import Optional
import tkinter as tk
from tkinter import messagebox, ttk

from estimator import PaymentEstimator


WORKBOOK_PATH = Path(__file__).with_name("Payment Terms.xlsx")


def resource_path(filename: str) -> Path:
    bundle_root = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return bundle_root / filename

PALETTE = {
    "bg_main": "#DFF3E5",
    "bg_panel": "#DFF3E5",
    "text_primary": "#1B4332",
    "text_secondary": "#52796F",
    "accent": "#2D6A4F",
    "accent_hover": "#1B4332",
    "border": "#BFE2CD",
    "toast_bg": "#1B4332",
    "toast_text": "#F1FAEE",
}


class PaymentEstimatorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Payment Terms Estimator")
        self.geometry("860x520")
        self.minsize(860, 520)

        self.estimator = PaymentEstimator(resource_path("Payment Terms.xlsx"))
        self.loan_types = sorted({row.loan_type for row in self.estimator.rows})
        self.terms = sorted({row.term_months for row in self.estimator.rows})

        self.amount_var = tk.StringVar()
        self.deposit_var = tk.StringVar(value="0")
        self.term_var = tk.StringVar(value=str(self.terms[0]))
        self.loan_type_var = tk.StringVar(value=self.loan_types[0])
        self.condition_var = tk.StringVar(value="new")
        self.points_var = tk.StringVar(value="2")

        self.result_vars = {
            "rate": tk.StringVar(value="-"),
            "residual": tk.StringVar(value="-"),
            "balloon": tk.StringVar(value="-"),
            "monthly": tk.StringVar(value="-"),
            "total": tk.StringVar(value="-"),
            "status": tk.StringVar(value="Enter values and calculate."),
        }
        self.quote_text = ""
        self.toast_window: Optional[tk.Toplevel] = None

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self) -> None:
        self.configure(bg=PALETTE["bg_main"])

        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("TFrame", background=PALETTE["bg_main"])
        style.configure(
            "TLabelframe",
            background=PALETTE["bg_panel"],
            bordercolor=PALETTE["border"],
            borderwidth=1,
        )
        style.configure(
            "TLabelframe.Label",
            background=PALETTE["bg_panel"],
            foreground=PALETTE["text_secondary"],
            font=("Segoe UI", 10, "bold"),
        )
        style.configure(
            "TLabel",
            background=PALETTE["bg_panel"],
            foreground=PALETTE["text_primary"],
            font=("Segoe UI", 10),
        )
        style.configure(
            "Title.TLabel",
            background=PALETTE["bg_main"],
            foreground=PALETTE["text_primary"],
            font=("Segoe UI", 20, "bold"),
        )
        style.configure(
            "Value.TLabel",
            background=PALETTE["bg_panel"],
            foreground=PALETTE["text_primary"],
            font=("Segoe UI", 11, "bold"),
        )
        style.configure(
            "TButton",
            background=PALETTE["accent"],
            foreground="#FFFFFF",
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            padding=(10, 6),
        )
        style.map(
            "TButton",
            background=[("active", PALETTE["accent_hover"]), ("pressed", PALETTE["accent_hover"])],
        )
        style.configure(
            "TEntry",
            fieldbackground="#FFFFFF",
            foreground=PALETTE["text_primary"],
            bordercolor=PALETTE["border"],
            lightcolor=PALETTE["border"],
            darkcolor=PALETTE["border"],
            padding=5,
        )
        style.configure(
            "TCombobox",
            fieldbackground="#FFFFFF",
            foreground=PALETTE["text_primary"],
            bordercolor=PALETTE["border"],
            lightcolor=PALETTE["border"],
            darkcolor=PALETTE["border"],
            arrowsize=14,
            padding=4,
        )

    def _build_ui(self) -> None:
        self.configure(padx=14, pady=14)
        input_frame = ttk.LabelFrame(self, text="Inputs", padding=14)
        input_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        result_frame = ttk.LabelFrame(self, text="Results", padding=14)
        result_frame.grid(row=0, column=1, columnspan=3, sticky="nsew")

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        result_frame.columnconfigure(1, weight=1)

        self._add_input_row(input_frame, 0, "Amount", ttk.Entry(input_frame, textvariable=self.amount_var, width=24))
        self._add_input_row(input_frame, 1, "Deposit", ttk.Entry(input_frame, textvariable=self.deposit_var, width=24))
        self._add_input_row(input_frame, 2, "Term Months", ttk.Combobox(input_frame, textvariable=self.term_var, values=[str(term) for term in self.terms], state="readonly", width=22))
        self._add_input_row(input_frame, 3, "Bus Condition", ttk.Combobox(input_frame, textvariable=self.condition_var, values=["new", "used"], state="readonly", width=22))
        self._add_input_row(input_frame, 4, "Loan Type", ttk.Combobox(input_frame, textvariable=self.loan_type_var, values=self.loan_types, state="readonly", width=22))
        self._add_input_row(input_frame, 5, "Points", ttk.Combobox(input_frame, textvariable=self.points_var, values=["0", "1", "2", "3", "4"], state="readonly", width=22))

        ttk.Button(input_frame, text="Calculate", command=self.calculate).grid(row=6, column=0, columnspan=2, sticky="ew", pady=(14, 0))

        self._add_result_row(result_frame, 0, "Rate", self.result_vars["rate"])
        self._add_result_row(result_frame, 1, "Residual", self.result_vars["residual"])
        self._add_result_row(result_frame, 2, "Balloon Amount", self.result_vars["balloon"])
        self._add_result_row(result_frame, 3, "Monthly Payment", self.result_vars["monthly"])
        self._add_result_row(result_frame, 4, "Total Payable", self.result_vars["total"])
        self._add_result_row(result_frame, 5, "Status", self.result_vars["status"])

        ttk.Button(result_frame, text="Copy Results", command=self.copy_results).grid(row=6, column=1, sticky="e", pady=(10, 0))

    @staticmethod
    def _add_input_row(parent: ttk.LabelFrame, row: int, label: str, widget: ttk.Widget) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 12))
        widget.grid(row=row, column=1, sticky="ew", pady=6)
        parent.columnconfigure(1, weight=1)

    @staticmethod
    def _add_result_row(parent: ttk.LabelFrame, row: int, label: str, value_var: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=6, padx=(0, 12))
        ttk.Label(parent, textvariable=value_var, style="Value.TLabel").grid(row=row, column=1, sticky="w", pady=6)

    def calculate(self) -> None:
        try:
            amount = float(self.amount_var.get())
            deposit = float(self.deposit_var.get())
            term_months = int(self.term_var.get())
            loan_type = self.loan_type_var.get().strip()
            condition = self.condition_var.get().strip().lower()
            points = int(self.points_var.get())

            result = self.estimator.quote(
                amount=amount,
                deposit=deposit,
                term_months=term_months,
                loan_type=loan_type,
                condition=condition,  # type: ignore[arg-type]
                points=points,
            )

            self.result_vars["rate"].set(f"{result.rate:.4%}")
            self.result_vars["residual"].set(f"{result.residual:.2%}")
            self.result_vars["balloon"].set(f"${result.balloon_amount:,.2f}")
            self.result_vars["monthly"].set(f"${result.monthly_payment:,.2f}")
            self.result_vars["total"].set(f"${result.total_payable:,.2f}")

            if result.used_fallback_points and result.points == 2:
                status = "No exact points match. Used default fallback points = 2."
            else:
                status = f"Matched points = {result.points}."

            self.result_vars["status"].set(status)
            self.quote_text = self._format_quote_output(
                amount=amount,
                deposit=deposit,
                condition=condition,
                loan_type=loan_type,
                term_months=term_months,
                rate=self.result_vars["rate"].get(),
                residual=self.result_vars["residual"].get(),
                balloon=self.result_vars["balloon"].get(),
                monthly=self.result_vars["monthly"].get(),
                total=self.result_vars["total"].get(),
            )
        except Exception as exc:
            messagebox.showerror("Calculation error", str(exc))

    def copy_results(self) -> None:
        if not self.quote_text:
            self._show_toast("Run a calculation first.")
            return

        self.clipboard_clear()
        self.clipboard_append(self.quote_text)
        self.result_vars["status"].set("Results copied to clipboard.")
        self._show_toast("Payment estimate copied to clipboard.")

    def _show_toast(self, text: str, duration_ms: int = 1800) -> None:
        if self.toast_window is not None and self.toast_window.winfo_exists():
            self.toast_window.destroy()

        toast = tk.Toplevel(self)
        self.toast_window = toast
        toast.overrideredirect(True)
        toast.attributes("-topmost", True)

        frame = tk.Frame(toast, bg=PALETTE["toast_bg"], bd=1, relief="solid")
        frame.pack(fill="both", expand=True)
        tk.Label(
            frame,
            text=text,
            bg=PALETTE["toast_bg"],
            fg=PALETTE["toast_text"],
            padx=12,
            pady=8,
            font=("Segoe UI", 10),
        ).pack()

        self.update_idletasks()
        toast.update_idletasks()
        x = self.winfo_rootx() + self.winfo_width() - toast.winfo_reqwidth() - 20
        y = self.winfo_rooty() + self.winfo_height() - toast.winfo_reqheight() - 20
        toast.geometry(f"+{x}+{y}")

        toast.after(duration_ms, toast.destroy)

    @staticmethod
    def _format_quote_output(
        amount: float,
        deposit: float,
        condition: str,
        loan_type: str,
        term_months: int,
        rate: str,
        residual: str,
        balloon: str,
        monthly: str,
        total: str,
    ) -> str:
        return (
            "Payment Estimate\n"
            f"Amount: ${amount:,.2f}\n"
            f"Deposit: ${deposit:,.2f}\n"
            f"Condition: {condition.title()}\n"
            f"Loan Type: {loan_type}\n"
            f"Term: {term_months} months\n"
            f"Rate: {rate}\n"
            f"Residual: {residual}\n"
            f"Balloon Amount: {balloon}\n"
            f"Monthly Payment: {monthly}\n"
            f"Total Payable: {total}"
        )


if __name__ == "__main__":
    app = PaymentEstimatorApp()
    app.mainloop()
