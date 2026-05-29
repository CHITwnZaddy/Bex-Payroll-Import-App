from pathlib import Path

from bex_payroll_import.models import RunOutputs
from bex_payroll_import.ui import is_validation_stopped


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
