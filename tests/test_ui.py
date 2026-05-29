from pathlib import Path

from bex_payroll_import.models import RunOutputs
from bex_payroll_import.ui import can_generate, is_validation_stopped


def test_is_validation_stopped_when_exported_csv_has_validation_errors(tmp_path: Path) -> None:
    outputs = RunOutputs(
        batch_code="CD0529143708",
        output_dir=tmp_path,
        audit_workbook_path=tmp_path / "audit.xlsx",
        payroll_csv_path=tmp_path / "PayrollImport_CD0529143708.csv",
        validation_path=tmp_path / "ValidationErrors_CD0529143708.csv",
        technical_log_path=None,
    )

    assert is_validation_stopped(outputs) is True


def test_can_generate_rejects_stale_template_path_with_existing_tdr_and_skipped_expense(tmp_path: Path) -> None:
    tdr_path = tmp_path / "tdr.xlsx"
    tdr_path.touch()

    assert can_generate(
        str(tmp_path / "missing-template.xlsx"),
        str(tdr_path),
        "",
        expense_skipped=True,
    ) is False


def test_can_generate_accepts_existing_template_and_tdr_with_skipped_expense(tmp_path: Path) -> None:
    template_path = tmp_path / "template.xlsx"
    tdr_path = tmp_path / "tdr.xlsx"
    template_path.touch()
    tdr_path.touch()

    assert can_generate(str(template_path), str(tdr_path), "", expense_skipped=True) is True


def test_can_generate_accepts_existing_template_tdr_and_expense(tmp_path: Path) -> None:
    template_path = tmp_path / "template.xlsx"
    tdr_path = tmp_path / "tdr.xlsx"
    expense_path = tmp_path / "expense.xlsx"
    template_path.touch()
    tdr_path.touch()
    expense_path.touch()

    assert can_generate(str(template_path), str(tdr_path), str(expense_path), expense_skipped=False) is True


def test_can_generate_rejects_missing_expense_when_not_skipped(tmp_path: Path) -> None:
    template_path = tmp_path / "template.xlsx"
    tdr_path = tmp_path / "tdr.xlsx"
    template_path.touch()
    tdr_path.touch()

    assert can_generate(
        str(template_path),
        str(tdr_path),
        str(tmp_path / "missing-expense.xlsx"),
        expense_skipped=False,
    ) is False
