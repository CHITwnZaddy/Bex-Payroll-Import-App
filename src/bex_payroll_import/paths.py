from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class PlannedOutputPaths:
    output_dir: Path
    audit_workbook_path: Path
    payroll_csv_path: Path
    validation_path: Path
    validation_errors_path: Path
    technical_log_path: Path


def build_output_paths(tdr_path: Path, batch_code: str, now: datetime) -> PlannedOutputPaths:
    timestamp = f"{now:%Y-%m-%d_%H-%M-%S}"
    output_dir = tdr_path.parent / "Payroll Import Outputs" / f"{now:%Y-%m-%d}"
    return PlannedOutputPaths(
        output_dir=output_dir,
        audit_workbook_path=output_dir / f"PayrollUpload_Audit_{batch_code}_{timestamp}.xlsx",
        payroll_csv_path=output_dir / f"PayrollImport_{batch_code}_{timestamp}.csv",
        validation_path=output_dir / f"ValidationReport_{batch_code}_{timestamp}.csv",
        validation_errors_path=output_dir / f"ValidationErrors_{batch_code}_{timestamp}.csv",
        technical_log_path=output_dir / f"TechnicalLog_{batch_code}_{timestamp}.log",
    )
