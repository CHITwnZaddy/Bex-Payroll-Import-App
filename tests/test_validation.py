from pathlib import Path

from bex_payroll_import.validation import (
    ValidationMessage,
    ValidationReport,
    validate_final_csv,
    write_validation_csv,
)


def test_validation_report_blocks_on_errors() -> None:
    report = ValidationReport()
    report.add_error("Missing column", "Add the Date column and rerun.", row_number=2)
    report.add_warning("Expense skipped", "User confirmed no expense file.")

    assert report.has_errors is True
    assert report.error_count == 1
    assert report.warning_count == 1


def test_write_validation_csv(tmp_path: Path) -> None:
    report = ValidationReport()
    report.messages.append(
        ValidationMessage(
            severity="ERROR",
            message="Missing column",
            suggested_fix="Add the Date column and rerun.",
            row_number=2,
            source="Time Detail Report",
        )
    )

    output = tmp_path / "ValidationErrors_CD0529143708.csv"
    write_validation_csv(report, output)

    text = output.read_text(encoding="utf-8")
    assert "severity,row_number,source,message,suggested_fix" in text
    assert "ERROR,2,Time Detail Report,Missing column,Add the Date column and rerun." in text


def test_validate_final_csv_blocks_header_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text(
        "Batch code,Employee code,Department\nCD0529143708,E100,DLPTO\n",
        encoding="utf-8",
    )
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert "Final CSV still contains a header row." in [
        message.message for message in report.messages
    ]


def test_validate_final_csv_blocks_empty_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text("", encoding="utf-8")
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert "Final PU output has no valid rows." in [
        message.message for message in report.messages
    ]
