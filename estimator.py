from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from openpyxl import load_workbook


BusCondition = Literal["new", "used"]


@dataclass(frozen=True)
class QuoteRow:
    condition: BusCondition
    points: int
    loan_type: str
    term_months: int
    rate: float
    residual: float
    balloon_amount: float


@dataclass(frozen=True)
class QuoteResult:
    amount: float
    deposit: float
    principal: float
    condition: BusCondition
    points: int
    loan_type: str
    term_months: int
    rate: float
    residual: float
    balloon_amount: float
    monthly_payment: float
    total_payable: float
    used_fallback_points: bool


class PaymentEstimator:
    def __init__(self, workbook_path: str | Path) -> None:
        self.workbook_path = Path(workbook_path)
        self.rows = self._load_rows()

    def _load_rows(self) -> list[QuoteRow]:
        workbook = load_workbook(self.workbook_path, data_only=True)
        rows: list[QuoteRow] = []

        for sheet_name in workbook.sheetnames:
            lower_name = sheet_name.lower()
            if "new" in lower_name:
                condition: BusCondition = "new"
            elif "used" in lower_name:
                condition = "used"
            else:
                continue

            sheet = workbook[sheet_name]
            for row_index in range(2, sheet.max_row + 1):
                points = sheet.cell(row_index, 1).value
                loan_type = sheet.cell(row_index, 2).value
                term_months = sheet.cell(row_index, 3).value
                rate = sheet.cell(row_index, 4).value
                residual = sheet.cell(row_index, 6).value
                balloon_amount = sheet.cell(row_index, 7).value

                if None in (points, loan_type, term_months, rate):
                    continue

                # Some product types in the sheet leave residual/balloon empty.
                # Keep those rows and treat the missing values as zero.
                residual_value = 0.0 if residual is None else float(residual)
                balloon_value = 0.0 if balloon_amount is None else float(balloon_amount)

                rows.append(
                    QuoteRow(
                        condition=condition,
                        points=int(points),
                        loan_type=str(loan_type),
                        term_months=int(term_months),
                        rate=float(rate),
                        residual=residual_value,
                        balloon_amount=balloon_value,
                    )
                )

        if not rows:
            raise ValueError("No pricing rows were loaded from the workbook.")

        return rows

    @staticmethod
    def _payment(principal: float, annual_rate: float, term_months: int, balloon_amount: float) -> float:
        if term_months <= 0:
            raise ValueError("term_months must be greater than zero")

        monthly_rate = annual_rate / 12.0
        if monthly_rate == 0:
            return (principal - balloon_amount) / term_months

        growth = (1 + monthly_rate) ** term_months
        return (monthly_rate * (principal * growth - balloon_amount)) / (growth - 1)

    def quote(
        self,
        amount: float,
        deposit: float,
        term_months: int,
        loan_type: str,
        condition: BusCondition,
        points: int,
    ) -> QuoteResult:
        if amount <= 0:
            raise ValueError("amount must be greater than zero")
        if deposit < 0:
            raise ValueError("deposit cannot be negative")

        principal = amount - deposit
        if principal <= 0:
            raise ValueError("deposit must be smaller than the amount")

        matches = [
            row
            for row in self.rows
            if row.condition == condition
            and row.loan_type.lower() == loan_type.lower()
            and row.term_months == term_months
        ]

        if not matches:
            raise ValueError("No worksheet row matches the selected condition, loan type, and term")

        selected = next((row for row in matches if row.points == points), None)
        used_fallback_points = False
        if selected is None:
            selected = next((row for row in matches if row.points == 2), None)
            used_fallback_points = True

        if selected is None:
            raise ValueError("No worksheet row matched the selected points or the default fallback of 2")

        effective_balloon_amount = principal * selected.residual if selected.residual > 0 else selected.balloon_amount

        monthly_payment = self._payment(
            principal=principal,
            annual_rate=selected.rate,
            term_months=term_months,
            balloon_amount=effective_balloon_amount,
        )

        return QuoteResult(
            amount=amount,
            deposit=deposit,
            principal=principal,
            condition=condition,
            points=selected.points,
            loan_type=selected.loan_type,
            term_months=selected.term_months,
            rate=selected.rate,
            residual=selected.residual,
            balloon_amount=round(effective_balloon_amount, 2),
            monthly_payment=round(monthly_payment, 2),
            total_payable=round(monthly_payment * term_months + deposit, 2),
            used_fallback_points=used_fallback_points,
        )
