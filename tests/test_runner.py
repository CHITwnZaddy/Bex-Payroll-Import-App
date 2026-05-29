from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from bex_payroll_import.models import RunInputs
from bex_payroll_import.runner import run_payroll_import
from tests.helpers import save_workbook


class FakeExcelBackend:
    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        csv_path.write_text("CD0529143708,E100,DLPTO,REG,8,100,200,L,04/12/2026,0\n", encoding="utf-8")


def make_template(path: Path) -> Path:
    workbook = Workbook()
    pu = workbook.active
    pu.title = "PU"
    workbook.create_sheet("Key")
    tdr = workbook.create_sheet("TDR")
    tdr.append(["Batch code", "EECode", "Lastname", "Firstname", "HomeDepartment", "Pay Class", "Badge", "Date", "Text to Col", "InPunchTime", "OutPunchTime", "Department", "EarnCode", "Hours", "Dollars", "Employee Approved", "Supervisor Approved", "Tax Profile", "Home Department Desc", "Dist Department Desc", "Job", "Employee code", "Phase", "Cost Code", "Cost type", "Home Department", "Department", "Phase Adj", "Phase", "Pay type"])
    tdr.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=B2", "=W2", "", "=Y2", "", "=L2", "", "=W2", "=M2"])
    key = workbook["Key"]
    key.append(["Employee code", "Last Name", "First Name"])
    key.append(["E100", "Doe", "Jane"])
    workbook.save(path)
    return path


def test_run_payroll_import_success_with_expenses(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    tdr_path = save_workbook(
        source_dir / "tdr.xlsx",
        "Time Detail Repor",
        [["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"], ["E100", "Doe", "Jane", "04/12/2026", "DLPTO", "REG", 8]],
    )
    expense_path = save_workbook(
        source_dir / "expense.xlsx",
        "Expense Transacti",
        [["Employee", "Expense Date", "Paid Amount"], ["DOE, JANE", "04/15/2026", 55]],
    )

    outputs = run_payroll_import(
        RunInputs(template_path=template_path, tdr_path=tdr_path, expense_path=expense_path, expense_skipped=False),
        now=datetime(2026, 5, 29, 14, 37, 8),
        excel_backend=FakeExcelBackend(),
    )

    assert outputs.batch_code == "CD0529143708"
    assert outputs.payroll_csv_path is not None
    assert outputs.payroll_csv_path.exists()
    assert outputs.audit_workbook_path.exists()
    assert outputs.validation_path.exists()


def test_run_payroll_import_hard_stops_when_expense_missing_and_not_confirmed(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    source_dir = tmp_path / "source"
    source_dir.mkdir()
    tdr_path = save_workbook(source_dir / "tdr.xlsx", "Time Detail Repor", [["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"]])

    outputs = run_payroll_import(
        RunInputs(template_path=template_path, tdr_path=tdr_path, expense_path=None, expense_skipped=False),
        now=datetime(2026, 5, 29, 14, 37, 8),
        excel_backend=FakeExcelBackend(),
    )

    assert outputs.payroll_csv_path is None
    assert outputs.validation_path.name.startswith("ValidationErrors_")
    assert outputs.validation_path.exists()
