from __future__ import annotations

import csv
from dataclasses import dataclass, field
from io import StringIO
from pathlib import Path
from typing import Literal

from bex_payroll_import.models import FINAL_CSV_COLUMN_COUNT


Severity = Literal["ERROR", "WARNING"]
LOOKUP_ERROR_VALUES = {"NL", "#N/A", "#REF!", "#VALUE!", "#NAME?"}


@dataclass(frozen=True)
class ValidationMessage:
    severity: Severity
    message: str
    suggested_fix: str
    row_number: int | None = None
    source: str = ""


@dataclass
class ValidationReport:
    messages: list[ValidationMessage] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(message.severity == "ERROR" for message in self.messages)

    @property
    def error_count(self) -> int:
        return sum(1 for message in self.messages if message.severity == "ERROR")

    @property
    def warning_count(self) -> int:
        return sum(1 for message in self.messages if message.severity == "WARNING")

    def add_error(
        self,
        message: str,
        suggested_fix: str,
        row_number: int | None = None,
        source: str = "",
    ) -> None:
        self.messages.append(
            ValidationMessage(
                severity="ERROR",
                message=message,
                suggested_fix=suggested_fix,
                row_number=row_number,
                source=source,
            )
        )

    def add_warning(
        self,
        message: str,
        suggested_fix: str,
        row_number: int | None = None,
        source: str = "",
    ) -> None:
        self.messages.append(
            ValidationMessage(
                severity="WARNING",
                message=message,
                suggested_fix=suggested_fix,
                row_number=row_number,
                source=source,
            )
        )


def write_validation_csv(report: ValidationReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=["severity", "row_number", "source", "message", "suggested_fix"],
        )
        writer.writeheader()
        for message in report.messages:
            writer.writerow(
                {
                    "severity": message.severity,
                    "row_number": message.row_number or "",
                    "source": message.source,
                    "message": message.message,
                    "suggested_fix": message.suggested_fix,
                }
            )


def validate_final_csv(
    csv_path: Path,
    report: ValidationReport,
    *,
    expected_batch_code: str | None = None,
    expected_output_rows: int | None = None,
) -> None:
    no_valid_rows_message = "Final PU output has no valid rows."
    no_valid_rows_fix = "Confirm the Time Detail Report has payroll rows and rerun."

    if not csv_path.exists():
        report.add_error(
            no_valid_rows_message,
            no_valid_rows_fix,
            source="PU CSV",
        )
        return

    text = csv_path.read_text(encoding="utf-8")
    if not text.strip():
        report.add_error(
            no_valid_rows_message,
            no_valid_rows_fix,
            source="PU CSV",
        )
        return

    rows = list(csv.reader(StringIO(text)))
    if not any(cell.strip() for row in rows for cell in row):
        report.add_error(
            no_valid_rows_message,
            no_valid_rows_fix,
            source="PU CSV",
        )
        return

    first_cells = [cell.strip().lower() for cell in rows[0]]
    if "batch code" in first_cells or "employee code" in first_cells:
        report.add_error(
            "Final CSV still contains a header row.",
            "Export the PU data rows without headers and rerun.",
            source="PU CSV",
        )
        return

    if expected_output_rows is not None and len(rows) != expected_output_rows:
        report.add_error(
            f"Expected {expected_output_rows} payroll rows but found {len(rows)}.",
            "Do not import this file. Confirm every source row reached the PU output and rerun.",
            source="PU CSV",
        )

    for row_number, row in enumerate(rows, start=1):
        validate_final_csv_row(row, row_number, report, expected_batch_code)


def validate_final_csv_row(
    row: list[str],
    row_number: int,
    report: ValidationReport,
    expected_batch_code: str | None,
) -> None:
    if len(row) != FINAL_CSV_COLUMN_COUNT:
        report.add_error(
            f"Expected {FINAL_CSV_COLUMN_COUNT} columns but found {len(row)}.",
            "Do not import this file. Regenerate the payroll CSV using the app.",
            row_number=row_number,
            source="PU CSV",
        )
        return

    values = [cell.strip() for cell in row]
    for column_number, value in enumerate(values, start=1):
        if value.upper() in LOOKUP_ERROR_VALUES:
            report.add_error(
                f"Excel lookup returned {value} in column {column_number}.",
                "Correct the source data or template lookup and rerun.",
                row_number=row_number,
                source="PU CSV",
            )

    required_fields = {
        "Batch Code": values[0],
        "Employee Code": values[1],
        "Department": values[2],
        "Pay Type": values[3],
        "Date": values[8],
    }
    for field_name, value in required_fields.items():
        require_output_value(field_name, value, row_number, report)

    if expected_batch_code is not None and values[0] != expected_batch_code:
        report.add_error(
            f"Batch Code must be {expected_batch_code}, found {values[0] or 'blank'}.",
            "Regenerate the payroll CSV so every row uses the current batch code.",
            row_number=row_number,
            source="PU CSV",
        )

    department = values[2].upper()
    pay_type = values[3].upper()
    is_expense = pay_type.startswith("EXP REIM")

    if is_expense:
        require_output_value("Dollars", values[21], row_number, report)
        require_output_match("Expense rows require Department 1", department, "1", row_number, report)
        require_output_match(
            "Expense rows require Pay Type EXP REIMB",
            pay_type,
            "EXP REIMB",
            row_number,
            report,
        )
        require_blank_output_value("Hours", values[4], row_number, report)
        require_blank_output_value("Job", values[5], row_number, report)
        require_blank_output_value("Phase", values[6], row_number, report)
        require_blank_output_value("Cost Type", values[7], row_number, report)
        return

    if values[21]:
        report.add_error(
            f"Time rows require blank Dollars, found {values[21]}.",
            "Regenerate the payroll CSV so Dollars are populated only on expense rows.",
            row_number=row_number,
            source="PU CSV",
        )

    if department == "DLPTO":
        require_output_value("Hours", values[4], row_number, report)
        if pay_type != "V":
            report.add_error(
                f"DLPTO time rows require Pay Type V, found {values[3] or 'blank'}.",
                "Correct the DLPTO pay rule and rerun.",
                row_number=row_number,
                source="PU CSV",
            )

    if department not in {"1", "5", "DLPTO"}:
        require_output_value("Job", values[5], row_number, report)
        require_output_value("Phase", values[6], row_number, report)
        require_output_value("Cost Type", values[7], row_number, report)


def require_output_value(
    field_name: str,
    value: str,
    row_number: int,
    report: ValidationReport,
) -> None:
    if value:
        return
    report.add_error(
        f"{field_name} is required but blank.",
        "Correct the source data or template formula and rerun.",
        row_number=row_number,
        source="PU CSV",
    )


def require_output_match(
    message: str,
    value: str,
    expected: str,
    row_number: int,
    report: ValidationReport,
) -> None:
    if value == expected:
        return
    report.add_error(
        f"{message}, found {value or 'blank'}.",
        "Regenerate the payroll CSV using the app's accepted expense mapping.",
        row_number=row_number,
        source="PU CSV",
    )


def require_blank_output_value(
    field_name: str,
    value: str,
    row_number: int,
    report: ValidationReport,
) -> None:
    if not value:
        return
    report.add_error(
        f"Expense rows require blank {field_name}, found {value}.",
        "Regenerate the payroll CSV using the app's accepted expense mapping.",
        row_number=row_number,
        source="PU CSV",
    )
