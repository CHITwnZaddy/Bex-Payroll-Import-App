from pathlib import Path

import pytest
from openpyxl import Workbook

from bex_payroll_import.key_lookup import (
    EmployeeLookup,
    load_employee_lookup,
    parse_expense_employee,
)


def make_key_workbook(path: Path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Key"
    sheet.append(["Employee code", "Last Name", "First Name"])
    sheet.append(["E100", "Doe", "Jane"])
    workbook.save(path)
    return path


def test_parse_expense_employee_last_comma_first() -> None:
    identity = parse_expense_employee("DOE, JANE")

    assert identity.last_name == "Doe"
    assert identity.first_name == "Jane"


def test_parse_expense_employee_rejects_bad_format() -> None:
    with pytest.raises(ValueError, match="Expected employee format"):
        parse_expense_employee("Jane Doe")


def test_lookup_employee_code_from_key_tab(tmp_path: Path) -> None:
    lookup = load_employee_lookup(make_key_workbook(tmp_path / "template.xlsx"))

    assert lookup.resolve_code("jane", "doe") == "E100"


def test_lookup_raises_for_unmatched_employee() -> None:
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        lookup.resolve_code("John", "Doe")
