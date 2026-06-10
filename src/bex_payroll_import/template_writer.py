from __future__ import annotations

import shutil
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.formula.translate import Translator

from bex_payroll_import.models import NormalizedPayrollRow, TDR_SHEET_NAME


def copy_template_to_audit(template_path: Path, audit_path: Path) -> None:
    audit_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, audit_path)


def write_tdr_rows(audit_workbook_path: Path, rows: list[NormalizedPayrollRow]) -> None:
    workbook = load_workbook(audit_workbook_path)
    workbook.iso_dates = True
    try:
        if TDR_SHEET_NAME not in workbook.sheetnames:
            raise ValueError("Template is missing the TDR tab.")
        sheet = workbook[TDR_SHEET_NAME]
        data_start_row = 2
        formula_cells = capture_formula_cells(sheet, data_start_row)
        clear_existing_tdr_rows(sheet, data_start_row)

        for offset, payroll_row in enumerate(rows):
            target_row = data_start_row + offset
            write_tdr_source_cells(sheet, target_row, payroll_row)
            copy_formula_cells(sheet, formula_cells, target_row)

        workbook.save(audit_workbook_path)
    finally:
        workbook.close()


def clear_existing_tdr_rows(sheet, data_start_row: int) -> None:
    clear_columns = [*range(1, 16), *range(21, 31)]
    for row in range(data_start_row, sheet.max_row + 1):
        for column in clear_columns:
            sheet.cell(row=row, column=column).value = None


def capture_formula_cells(sheet, data_start_row: int) -> dict[int, tuple[str, str]]:
    for row in range(data_start_row, sheet.max_row + 1):
        formula_cells = {}
        for column in range(22, 31):
            cell = sheet.cell(row=row, column=column)
            value = cell.value
            if isinstance(value, str) and value.startswith("="):
                formula_cells[column] = (value, cell.coordinate)
        if formula_cells:
            return formula_cells
    return {}


def write_tdr_source_cells(sheet, row_number: int, payroll_row: NormalizedPayrollRow) -> None:
    sheet.cell(row=row_number, column=1).value = payroll_row.batch_code
    sheet.cell(row=row_number, column=2).value = excel_lookup_code(payroll_row.employee_code)
    sheet.cell(row=row_number, column=8).value = payroll_row.work_date
    sheet.cell(row=row_number, column=12).value = payroll_row.department
    sheet.cell(row=row_number, column=13).value = payroll_row.pay_type
    sheet.cell(row=row_number, column=14).value = float(payroll_row.hours)
    sheet.cell(row=row_number, column=15).value = float(payroll_row.dollars or 0)
    sheet.cell(row=row_number, column=21).value = payroll_row.job
    sheet.cell(row=row_number, column=23).value = payroll_row.phase
    sheet.cell(row=row_number, column=25).value = payroll_row.cost_type


def copy_formula_cells(
    sheet,
    formula_cells: dict[int, tuple[str, str]],
    target_row: int,
) -> None:
    source_value_columns = {23, 25}
    for column, (formula, origin) in formula_cells.items():
        if column in source_value_columns:
            continue
        target_cell = sheet.cell(row=target_row, column=column)
        target_cell.value = Translator(
            formula,
            origin=origin,
        ).translate_formula(target_cell.coordinate)


def excel_lookup_code(value: str) -> str | int:
    text = value.strip()
    if text.isdigit():
        return int(text)
    return text
