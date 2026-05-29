from __future__ import annotations

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
        except Exception as exc:
            operation_error = exc
        finally:
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
    return ExcelExportResult(csv_path=csv_path)
