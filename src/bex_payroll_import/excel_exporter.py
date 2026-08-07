from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from bex_payroll_import.errors import ExcelAutomationError
from bex_payroll_import.models import FINAL_CSV_COLUMN_COUNT


@dataclass(frozen=True)
class ExcelExportResult:
    csv_path: Path


class ExcelBackend(Protocol):
    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        ...


class Win32ExcelBackend:
    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        error_prefix = "Excel could not recalculate and export the PU tab"
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ExcelAutomationError(
                "pywin32 is required on Windows for Excel export."
            ) from exc

        excel = None
        workbook = None
        csv_workbook = None
        operation_error = None
        cleanup_error = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            workbook = excel.Workbooks.Open(str(workbook_path))
            excel.CalculateFullRebuild()
            workbook.Save()
            worksheet = workbook.Worksheets("PU")
            worksheet.Copy()
            csv_workbook = excel.ActiveWorkbook
            csv_workbook.SaveAs(str(csv_path), FileFormat=6)
            csv_workbook.Close(SaveChanges=False)
            csv_workbook = None
        except Exception as exc:
            operation_error = exc
        finally:
            if csv_workbook is not None:
                try:
                    csv_workbook.Close(SaveChanges=False)
                except Exception as exc:
                    cleanup_error = cleanup_error or exc
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=True)
                except Exception as exc:
                    cleanup_error = cleanup_error or exc
            if excel is not None:
                try:
                    excel.Quit()
                except Exception as exc:
                    cleanup_error = cleanup_error or exc

        if operation_error is not None:
            raise ExcelAutomationError(
                f"{error_prefix}: {operation_error}"
            ) from operation_error
        if cleanup_error is not None:
            raise ExcelAutomationError(
                f"{error_prefix}: {cleanup_error}"
            ) from cleanup_error


def export_pu_csv_with_excel(
    workbook_path: Path,
    csv_path: Path,
    backend: ExcelBackend | None = None,
    batch_code: str | None = None,
) -> ExcelExportResult:
    selected_backend = backend or Win32ExcelBackend()
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    selected_backend.recalculate_and_export(workbook_path, csv_path)
    normalize_pu_csv(csv_path, batch_code)
    return ExcelExportResult(csv_path=csv_path)


def normalize_pu_csv(csv_path: Path, batch_code: str | None) -> None:
    with csv_path.open(newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        return

    first_cells = [cell.strip().lower() for cell in rows[0]]
    if "batch code" in first_cells or "employee code" in first_cells:
        rows = rows[1:]

    normalized_rows = []
    for row in rows:
        employee_code = row[1].strip() if len(row) > 1 else ""
        if not employee_code:
            continue
        normalized_row = row[:FINAL_CSV_COLUMN_COUNT]
        normalized_row.extend([""] * (FINAL_CSV_COLUMN_COUNT - len(normalized_row)))
        if batch_code is not None:
            normalized_row[0] = batch_code
        if normalized_row[3].strip().upper().startswith("EXP REIM"):
            normalized_row[2] = "1"
            normalized_row[3] = "EXP REIMB"
            normalized_row[4:8] = [""] * 4
        else:
            normalized_row[21] = ""
        normalized_rows.append(normalized_row)

    with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(normalized_rows)
