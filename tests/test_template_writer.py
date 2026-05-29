from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from bex_payroll_import.models import NormalizedPayrollRow
from bex_payroll_import.template_writer import copy_template_to_audit, write_tdr_rows


def make_template(path: Path) -> Path:
    workbook = Workbook()
    pu = workbook.active
    pu.title = "PU"
    tdr = workbook.create_sheet("TDR")
    workbook.create_sheet("Key")
    tdr.append(["Batch code", "EECode", "Lastname", "Firstname", "HomeDepartment", "Pay Class", "Badge", "Date", "Text to Col", "InPunchTime", "OutPunchTime", "Department", "EarnCode", "Hours", "Dollars", "Employee Approved", "Supervisor Approved", "Tax Profile", "Home Department Desc", "Dist Department Desc", "Job", "Employee code", "Phase", "Cost Code", "Cost type", "Home Department", "Department", "Phase Adj", "Phase", "Pay type"])
    tdr.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=B2", "=W2", "", "=Y2", "", "=L2", "", "=W2", "=M2"])
    workbook.save(path)
    return path


def test_copy_template_to_audit_preserves_original(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    audit_path = tmp_path / "out" / "audit.xlsx"

    copy_template_to_audit(template_path, audit_path)

    assert template_path.exists()
    assert audit_path.exists()


def test_write_tdr_rows_adds_rows_and_copies_formula_columns(tmp_path: Path) -> None:
    audit_path = tmp_path / "audit.xlsx"
    make_template(audit_path)
    rows = [
        NormalizedPayrollRow("CD0529143708", "E100", "DLPTO", "REG", Decimal("8"), "100", "200", "L", date(2026, 4, 12), Decimal("0"), "Time Detail Report", 2),
        NormalizedPayrollRow("CD0529143708", "E101", "DLPTO", "EXP REIM", Decimal("0"), "", "", "L", date(2026, 4, 15), Decimal("55.25"), "Expense Transaction", 2),
    ]

    write_tdr_rows(audit_path, rows)

    workbook = load_workbook(audit_path, data_only=False)
    sheet = workbook["TDR"]
    assert sheet["A2"].value == "CD0529143708"
    assert sheet["B2"].value == "E100"
    assert sheet["H2"].value == date(2026, 4, 12)
    assert sheet["N3"].value == 0
    assert sheet["O3"].value == 55.25
    assert sheet["V3"].value == "=B3"
