from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path


TDR_SHEET_NAME = "TDR"
PU_SHEET_NAME = "PU"
KEY_SHEET_NAME = "Key"
FINAL_CSV_COLUMN_COUNT = 22

FINAL_OUTPUT_COLUMNS = [
    "Batch code",
    "Employee code",
    "Department",
    "Pay type",
    "Hours",
    "Job",
    "Phase",
    "Cost type",
    "Date",
    "$",
]


@dataclass(frozen=True)
class EmployeeIdentity:
    first_name: str
    last_name: str


@dataclass(frozen=True)
class NormalizedPayrollRow:
    batch_code: str
    employee_code: str
    department: str
    pay_type: str
    hours: Decimal | None
    job: str
    phase: str
    cost_type: str
    work_date: date
    dollars: Decimal | None
    source: str
    source_row_number: int


@dataclass(frozen=True)
class RunInputs:
    template_path: Path
    tdr_path: Path
    expense_path: Path | None
    expense_skipped: bool


@dataclass(frozen=True)
class RunOutputs:
    batch_code: str
    output_dir: Path
    audit_workbook_path: Path
    payroll_csv_path: Path | None
    validation_path: Path
    technical_log_path: Path | None


@dataclass(frozen=True)
class RunClock:
    now: datetime
