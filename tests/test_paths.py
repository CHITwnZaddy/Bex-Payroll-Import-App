from datetime import datetime
from pathlib import Path

from bex_payroll_import.paths import build_output_paths


def test_build_output_paths_next_to_tdr(tmp_path: Path) -> None:
    tdr_path = tmp_path / "BEX Payroll - #05" / "tdr.xlsx"
    now = datetime(2026, 5, 29, 14, 37, 8)

    outputs = build_output_paths(tdr_path, "CD0529143708", now)

    assert outputs.output_dir == tmp_path / "BEX Payroll - #05" / "Payroll Import Outputs" / "2026-05-29"
    assert outputs.payroll_csv_path.name == "PayrollImport_CD0529143708_2026-05-29_14-37-08.csv"
    assert outputs.audit_workbook_path.name == "PayrollUpload_Audit_CD0529143708_2026-05-29_14-37-08.xlsx"
    assert outputs.validation_path.name == "ValidationReport_CD0529143708_2026-05-29_14-37-08.csv"
    assert outputs.validation_errors_path.name == "ValidationErrors_CD0529143708_2026-05-29_14-37-08.csv"
    assert outputs.technical_log_path.name == "TechnicalLog_CD0529143708_2026-05-29_14-37-08.log"
