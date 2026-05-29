from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from bex_payroll_import.key_lookup import EmployeeLookup
from bex_payroll_import.normalizers import normalize_expense_file, normalize_tdr_file
from tests.helpers import save_workbook


def test_normalize_tdr_file_reads_required_values(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            [
                "EECode",
                "Lastname",
                "Firstname",
                "HomeDepartment",
                "Pay Class",
                "Badge",
                "Date",
                "InPunchTime",
                "OutPunchTime",
                "Department",
                "EarnCode",
                "EarnHours",
                "Dollars",
                "Job",
                "Phase",
                "Cost Type",
            ],
            [
                "E100",
                "Doe",
                "Jane",
                "DLPTO",
                "REG",
                "1",
                date(2026, 4, 12),
                "",
                "",
                "OPS",
                "REG",
                8,
                0,
                "100",
                "200",
                "L",
            ],
        ],
    )

    rows = normalize_tdr_file(path, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].batch_code == "CD0529143708"
    assert rows[0].employee_code == "E100"
    assert rows[0].hours == Decimal("8")
    assert rows[0].work_date == date(2026, 4, 12)


def test_normalize_expense_file_applies_defaults(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [
            ["Employee", "Expense Date", "Paid Amount"],
            ["DOE, JANE", datetime(2026, 4, 15, 8, 30), 123.45],
        ],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    rows = normalize_expense_file(path, lookup, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].employee_code == "E100"
    assert rows[0].department == "DLPTO"
    assert rows[0].pay_type == "EXP REIM"
    assert rows[0].hours == Decimal("0")
    assert rows[0].job == ""
    assert rows[0].phase == ""
    assert rows[0].cost_type == "L"
    assert rows[0].dollars == Decimal("123.45")


def test_normalize_expense_file_skips_whitespace_only_employee_rows(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [
            ["Employee", "Expense Date", "Paid Amount"],
            ["   ", date(2026, 4, 15), 50],
            ["DOE, JANE", date(2026, 4, 16), 75],
        ],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    rows = normalize_expense_file(path, lookup, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].employee_code == "E100"
    assert rows[0].dollars == Decimal("75")


def test_normalize_expense_file_raises_for_unmatched_employee(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [["Employee", "Expense Date", "Paid Amount"], ["DOE, JOHN", date(2026, 4, 15), 50]],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        normalize_expense_file(path, lookup, "CD0529143708")
