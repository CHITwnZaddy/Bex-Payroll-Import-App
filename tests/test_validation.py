import csv
from pathlib import Path

from bex_payroll_import.validation import (
    ValidationMessage,
    ValidationReport,
    validate_final_csv,
    write_validation_csv,
)


NO_VALID_ROWS_MESSAGE = "Final PU output has no valid rows."
NO_VALID_ROWS_FIX = "Confirm the Time Detail Report has payroll rows and rerun."
HEADER_ROW_MESSAGE = "Final CSV still contains a header row."
HEADER_ROW_FIX = "Export the PU data rows without headers and rerun."


def make_pu_row(
    *,
    batch: str = "CD0529143708",
    employee: str = "E100",
    department: str = "6",
    pay_type: str = "R",
    hours: str = "8.00",
    job: str = "25009",
    phase: str = "011010",
    cost_type: str = "L",
    work_date: str = "4/12/2026",
    dollars: str = "",
) -> list[str]:
    row = [batch, employee, department, pay_type, hours, job, phase, cost_type, work_date]
    row.extend([""] * 12)
    row.append(dollars)
    return row


def write_pu_rows(path: Path, rows: list[list[str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as csv_file:
        csv.writer(csv_file).writerows(rows)


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
    assert report.messages[0].message == HEADER_ROW_MESSAGE
    assert report.messages[0].suggested_fix == HEADER_ROW_FIX
    assert report.messages[0].source == "PU CSV"


def test_validate_final_csv_blocks_quoted_header_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text(
        '"Batch code","Employee code",Department\nCD0529143708,E100,DLPTO\n',
        encoding="utf-8",
    )
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert report.messages[0].message == HEADER_ROW_MESSAGE
    assert report.messages[0].suggested_fix == HEADER_ROW_FIX
    assert report.messages[0].source == "PU CSV"


def test_validate_final_csv_blocks_empty_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text("", encoding="utf-8")
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert report.messages[0].message == NO_VALID_ROWS_MESSAGE
    assert report.messages[0].suggested_fix == NO_VALID_ROWS_FIX
    assert report.messages[0].source == "PU CSV"


def test_validate_final_csv_blocks_missing_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert report.messages[0].message == NO_VALID_ROWS_MESSAGE
    assert report.messages[0].suggested_fix == NO_VALID_ROWS_FIX
    assert report.messages[0].source == "PU CSV"


def test_validate_final_csv_blocks_delimiter_only_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text(",,,,\n,,,,\n", encoding="utf-8")
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert report.messages[0].message == NO_VALID_ROWS_MESSAGE
    assert report.messages[0].suggested_fix == NO_VALID_ROWS_FIX
    assert report.messages[0].source == "PU CSV"


def test_validate_final_csv_blocks_when_output_row_count_does_not_match_template(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row()])
    report = ValidationReport()

    validate_final_csv(csv_path, report, expected_output_rows=2)

    assert report.has_errors is True
    assert any("Expected 2 payroll rows but found 1" in message.message for message in report.messages)


def test_validate_final_csv_blocks_wrong_column_count(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row()[:9]])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert any("Expected 22 columns but found 9" in message.message for message in report.messages)


def test_validate_final_csv_blocks_excel_lookup_failures(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row(pay_type="NL")])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert any("Excel lookup returned NL" in message.message for message in report.messages)


def test_validate_final_csv_blocks_missing_project_fields(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row(job="", phase="", cost_type="")])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    messages = [message.message for message in report.messages]
    assert any("Job is required" in message for message in messages)
    assert any("Phase is required" in message for message in messages)
    assert any("Cost Type is required" in message for message in messages)


def test_validate_final_csv_blocks_invalid_dlpto_time_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row(department="DLPTO", pay_type="H", hours="")])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    messages = [message.message for message in report.messages]
    assert any("DLPTO time rows require Pay Type V" in message for message in messages)
    assert any("Hours is required" in message for message in messages)


def test_validate_final_csv_allows_known_good_expense_mapping(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(
        csv_path,
        [
            make_pu_row(
                department="1",
                pay_type="EXP REIMB",
                hours="",
                job="",
                phase="",
                cost_type="",
                dollars="55.25",
            )
        ],
    )
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is False


def test_validate_final_csv_blocks_old_dlpto_expense_mapping(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(
        csv_path,
        [
            make_pu_row(
                department="DLPTO",
                pay_type="EXP REIM",
                hours="0",
                job="",
                phase="",
                cost_type="L",
                dollars="55.25",
            )
        ],
    )
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    messages = [message.message for message in report.messages]
    assert any("Expense rows require Department 1" in message for message in messages)
    assert any("Expense rows require Pay Type EXP REIMB" in message for message in messages)
    assert any("Expense rows require blank Hours" in message for message in messages)
    assert any("Expense rows require blank Cost Type" in message for message in messages)


def test_validate_final_csv_allows_blank_hours_for_project_row_seen_in_known_good_import(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row(department="6", hours="")])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is False


def test_validate_final_csv_blocks_dollars_on_time_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    write_pu_rows(csv_path, [make_pu_row(dollars="0.00")])
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert any("Time rows require blank Dollars" in message.message for message in report.messages)
