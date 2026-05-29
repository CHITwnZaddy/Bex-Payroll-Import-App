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
        try:
            import win32com.client  # type: ignore[import-not-found]
        except ImportError as exc:
            raise ExcelAutomationError(
                "pywin32 is required on Windows for Excel export."
            ) from exc

        excel = None
        workbook = None
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
            raise ExcelAutomationError(
                f"Excel could not recalculate and export the PU tab: {exc}"
            ) from exc
        finally:
            if workbook is not None:
                workbook.Close(SaveChanges=True)
            if excel is not None:
                excel.Quit()


def export_pu_csv_with_excel(
    workbook_path: Path,
    csv_path: Path,
    backend: ExcelBackend | None = None,
) -> ExcelExportResult:
    selected_backend = backend or Win32ExcelBackend()
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    selected_backend.recalculate_and_export(workbook_path, csv_path)
    return ExcelExportResult(csv_path=csv_path)
