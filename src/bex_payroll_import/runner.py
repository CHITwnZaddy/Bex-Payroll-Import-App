from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

from bex_payroll_import.batch import build_batch_code
from bex_payroll_import.excel_exporter import ExcelBackend, export_pu_csv_with_excel
from bex_payroll_import.models import NormalizedPayrollRow, RunInputs, RunOutputs
from bex_payroll_import.normalizers import (
    load_employee_lookup_from_tdr,
    load_employee_lookup_from_template_tdr,
    normalize_expense_file,
    normalize_tdr_file,
)
from bex_payroll_import.paths import PlannedOutputPaths, build_output_paths
from bex_payroll_import.template_writer import copy_template_to_audit, write_tdr_rows
from bex_payroll_import.validation import ValidationReport, validate_final_csv, write_validation_csv


SOURCE_TDR = "Time Detail Report"
SOURCE_EXPENSE = "Expense Transaction"
SOURCE_EXPORT = "PU CSV"
SOURCE_TEMPLATE = "Template"

SOURCE_FIX = "Correct the source file or template lookup data, then rerun."


def run_payroll_import(
    inputs: RunInputs,
    now: datetime | None = None,
    excel_backend: ExcelBackend | None = None,
) -> RunOutputs:
    run_now = now or datetime.now()
    batch_code = build_batch_code(run_now)
    paths = build_output_paths(inputs.tdr_path, batch_code, run_now)
    paths.output_dir.mkdir(parents=True, exist_ok=True)

    report = ValidationReport()
    validate_inputs(inputs, report)
    if report.has_errors:
        return validation_error_outputs(batch_code, paths, report)

    try:
        copy_template_to_audit(inputs.template_path, paths.audit_workbook_path)
    except Exception as exc:
        report.add_error(str(exc), SOURCE_FIX, source=SOURCE_TEMPLATE)
        return validation_error_outputs(batch_code, paths, report)

    if inputs.expense_path is None and not inputs.expense_skipped:
        report.add_error(
            "Expense Transaction decision is missing.",
            "Select the Expense Transaction file or confirm there is no Expense Transaction file for this payroll import.",
            source=SOURCE_EXPENSE,
        )
        return validation_error_outputs(batch_code, paths, report)

    payroll_rows = normalize_source_rows(inputs, batch_code, report, run_now.date())
    if not payroll_rows and not report.has_errors:
        report.add_error(
            "Selected source file has no usable data rows.",
            "Confirm the Time Detail Report contains payroll rows and rerun.",
            source=SOURCE_TDR,
        )

    if report.has_errors:
        return validation_error_outputs(batch_code, paths, report)

    try:
        template_result = write_tdr_rows(paths.audit_workbook_path, payroll_rows)
    except Exception as exc:
        report.add_error(str(exc), SOURCE_FIX, source=SOURCE_TEMPLATE)
        return validation_error_outputs(batch_code, paths, report)

    exported_csv_path: Path | None = None
    try:
        export_pu_csv_with_excel(
            paths.audit_workbook_path,
            paths.payroll_csv_path,
            backend=excel_backend,
            batch_code=batch_code,
        )
        exported_csv_path = paths.payroll_csv_path
        validate_final_csv(
            paths.payroll_csv_path,
            report,
            expected_batch_code=batch_code,
            expected_output_rows=template_result.expected_output_rows,
        )
    except Exception as exc:
        report.add_error(str(exc), SOURCE_FIX, source=SOURCE_EXPORT)

    if report.has_errors:
        return validation_error_outputs(batch_code, paths, report, payroll_csv_path=exported_csv_path)

    write_validation_csv(report, paths.validation_path)
    return RunOutputs(
        batch_code=batch_code,
        output_dir=paths.output_dir,
        audit_workbook_path=paths.audit_workbook_path,
        payroll_csv_path=paths.payroll_csv_path,
        validation_path=paths.validation_path,
        technical_log_path=paths.technical_log_path,
    )


def validate_inputs(inputs: RunInputs, report: ValidationReport) -> None:
    require_existing_file(inputs.template_path, "Template", report)
    require_existing_file(inputs.tdr_path, SOURCE_TDR, report)
    if inputs.expense_path is not None:
        require_existing_file(inputs.expense_path, SOURCE_EXPENSE, report)


def require_existing_file(path: Path, label: str, report: ValidationReport) -> None:
    if not path.exists():
        report.add_error(
            f"{label} file does not exist.",
            f"Select an existing {label} file and rerun.",
            source=label,
        )
        return
    if not path.is_file():
        report.add_error(
            f"{label} path is not a file.",
            f"Select a valid {label} file and rerun.",
            source=label,
        )


def normalize_source_rows(
    inputs: RunInputs,
    batch_code: str,
    report: ValidationReport,
    default_expense_check_date: date,
) -> list[NormalizedPayrollRow]:
    rows: list[NormalizedPayrollRow] = []

    try:
        rows.extend(normalize_tdr_file(inputs.tdr_path, batch_code))
    except Exception as exc:
        report.add_error(str(exc), SOURCE_FIX, source=SOURCE_TDR)

    if inputs.expense_path is None:
        report.add_warning(
            "Expense Transaction file intentionally skipped.",
            "No action needed if this payroll import has no Expense Transaction file.",
            source=SOURCE_EXPENSE,
        )
        return rows

    try:
        lookup = load_employee_lookup_from_tdr(inputs.tdr_path)
        template_lookup = load_employee_lookup_from_template_tdr(inputs.template_path)
        lookup = lookup.with_fallback(template_lookup)
        rows.extend(normalize_expense_file(inputs.expense_path, lookup, batch_code, default_expense_check_date))
    except Exception as exc:
        report.add_error(str(exc), SOURCE_FIX, source=SOURCE_EXPENSE)

    return rows


def validation_error_outputs(
    batch_code: str,
    paths: PlannedOutputPaths,
    report: ValidationReport,
    payroll_csv_path: Path | None = None,
) -> RunOutputs:
    write_validation_csv(report, paths.validation_errors_path)
    return RunOutputs(
        batch_code=batch_code,
        output_dir=paths.output_dir,
        audit_workbook_path=paths.audit_workbook_path,
        payroll_csv_path=payroll_csv_path,
        validation_path=paths.validation_errors_path,
        technical_log_path=paths.technical_log_path,
    )
