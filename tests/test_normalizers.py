from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from bex_payroll_import.key_lookup import EmployeeLookup
from bex_payroll_import.normalizers import (
    load_employee_lookup_from_tdr,
    normalize_expense_file,
    normalize_tdr_file,
)
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


def test_normalize_tdr_file_uses_in_punch_time_when_date_column_missing(tmp_path: Path) -> None:
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
                "InPunchTime",
                "OutPunchTime",
                "Department",
                "EarnCode",
                "EarnHours",
                "Dollars",
            ],
            [
                "0048",
                "BAKER",
                "KENNETH",
                "200",
                "SAL",
                "0049",
                "2026-04-13 12:00 AM",
                "2026-04-13 12:00 AM",
                "9 Mile improvements",
                "R",
                8,
                0,
            ],
        ],
    )

    rows = normalize_tdr_file(path, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].employee_code == "0048"
    assert rows[0].work_date == date(2026, 4, 13)


def test_load_employee_lookup_from_tdr_maps_expense_names_to_codes(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            ["EECode", "Lastname", "Firstname", "InPunchTime", "Department", "EarnCode", "EarnHours"],
            ["0048", "BAKER", "KENNETH", "2026-04-13 12:00 AM", "9 Mile improvements", "R", 8],
            ["1009", "LUMLEY", "MICHAEL", "2026-04-13 12:00 AM", "9 Mile improvements", "R", 8],
        ],
    )

    lookup = load_employee_lookup_from_tdr(path)

    assert lookup.resolve_code("Kenneth", "Baker") == "0048"
    assert lookup.resolve_code("Michael", "Lumley") == "1009"


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


def test_normalize_expense_file_treats_whitespace_paid_amount_as_zero(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [
            ["Employee", "Expense Date", "Paid Amount"],
            ["DOE, JANE", date(2026, 4, 15), "   "],
        ],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    rows = normalize_expense_file(path, lookup, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].dollars == Decimal("0")


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


def test_normalize_tdr_file_treats_whitespace_hours_as_zero(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            ["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"],
            ["E100", "Doe", "Jane", date(2026, 4, 12), "OPS", "REG", "   "],
        ],
    )

    rows = normalize_tdr_file(path, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].hours == Decimal("0")


def test_normalize_tdr_file_treats_whitespace_dollars_as_none(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            ["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours", "Dollars"],
            ["E100", "Doe", "Jane", date(2026, 4, 12), "OPS", "REG", 8, "   "],
        ],
    )

    rows = normalize_tdr_file(path, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].dollars is None


def test_normalize_expense_file_raises_for_unmatched_employee(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [["Employee", "Expense Date", "Paid Amount"], ["DOE, JOHN", date(2026, 4, 15), 50]],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        normalize_expense_file(path, lookup, "CD0529143708")
