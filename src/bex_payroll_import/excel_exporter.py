from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

from bex_payroll_import.errors import ExcelAutomationError


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
) -> ExcelExportResult:
    selected_backend = backend or Win32ExcelBackend()
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    selected_backend.recalculate_and_export(workbook_path, csv_path)
    remove_pu_header_row(csv_path)
    return ExcelExportResult(csv_path=csv_path)


def remove_pu_header_row(csv_path: Path) -> None:
    with csv_path.open(newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        return

    first_cells = [cell.strip().lower() for cell in rows[0]]
    if "batch code" not in first_cells and "employee code" not in first_cells:
        return

    with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.writer(csv_file)
        writer.writerows(rows[1:])
