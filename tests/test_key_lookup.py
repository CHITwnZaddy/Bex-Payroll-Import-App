from pathlib import Path

import pytest
from openpyxl import Workbook

from bex_payroll_import.key_lookup import (
    EmployeeLookup,
    cell_value,
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


def test_lookup_skips_rows_with_blank_codes_after_stripping(tmp_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Key"
    sheet.append(["Employee code", "Last Name", "First Name"])
    sheet.append(["   ", "Doe", "Jane"])
    template_path = tmp_path / "template.xlsx"
    workbook.save(template_path)

    lookup = load_employee_lookup(template_path)

    with pytest.raises(KeyError, match="No employee code found"):
        lookup.resolve_code("Jane", "Doe")


def test_lookup_raises_for_unmatched_employee() -> None:
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        lookup.resolve_code("John", "Doe")


def test_cell_value_rejects_zero_index() -> None:
    assert cell_value(("first", "last"), 0) is None
