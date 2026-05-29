from pathlib import Path
import sys
from types import ModuleType, SimpleNamespace

import pytest

from bex_payroll_import.errors import ExcelAutomationError
from bex_payroll_import.excel_exporter import (
    ExcelExportResult,
    Win32ExcelBackend,
    export_pu_csv_with_excel,
)


class FakeExcelBackend:
    def __init__(self) -> None:
        self.calls: list[tuple[Path, Path]] = []

    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        self.calls.append((workbook_path, csv_path))
        assert csv_path.parent.exists()
        csv_path.write_text(
            "CD0529143708,E100,DLPTO,REG,8,100,200,L,04/12/2026,0\n",
            encoding="utf-8",
        )


def test_export_pu_csv_with_excel_delegates_to_backend(tmp_path: Path) -> None:
    backend = FakeExcelBackend()
    workbook_path = tmp_path / "audit.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    csv_path = tmp_path / "nested" / "exports" / "PayrollImport.csv"

    result = export_pu_csv_with_excel(workbook_path, csv_path, backend=backend)

    assert result == ExcelExportResult(csv_path=csv_path)
    assert backend.calls == [(workbook_path, csv_path)]
    assert csv_path.exists()


class FakeCsvWorkbook:
    def SaveAs(self, _csv_path: str, FileFormat: int) -> None:
        assert FileFormat == 6

    def Close(self, SaveChanges: bool) -> None:
        assert SaveChanges is False


class FakeWorksheet:
    def __init__(self, excel: "FakeExcelApp") -> None:
        self.excel = excel

    def Copy(self) -> None:
        self.excel.ActiveWorkbook = FakeCsvWorkbook()


class FakeWorkbook:
    def __init__(
        self,
        excel: "FakeExcelApp",
        *,
        fail_save: bool = False,
        fail_close: bool = False,
    ) -> None:
        self.excel = excel
        self.fail_save = fail_save
        self.fail_close = fail_close

    def Save(self) -> None:
        if self.fail_save:
            raise RuntimeError("save failed")

    def Worksheets(self, name: str) -> FakeWorksheet:
        assert name == "PU"
        return FakeWorksheet(self.excel)

    def Close(self, SaveChanges: bool) -> None:
        assert SaveChanges is True
        if self.fail_close:
            raise RuntimeError("workbook close failed")


class FakeWorkbooks:
    def __init__(self, workbook: FakeWorkbook) -> None:
        self.workbook = workbook

    def Open(self, _workbook_path: str) -> FakeWorkbook:
        return self.workbook


class FakeExcelApp:
    def __init__(
        self,
        *,
        fail_save: bool = False,
        fail_close: bool = False,
        fail_quit: bool = False,
    ) -> None:
        self.ActiveWorkbook = None
        self.workbook = FakeWorkbook(
            self,
            fail_save=fail_save,
            fail_close=fail_close,
        )
        self.Workbooks = FakeWorkbooks(self.workbook)
        self.fail_quit = fail_quit

    def CalculateFullRebuild(self) -> None:
        pass

    def Quit(self) -> None:
        if self.fail_quit:
            raise RuntimeError("excel quit failed")


def install_fake_win32com(monkeypatch: pytest.MonkeyPatch, excel: FakeExcelApp) -> None:
    win32com = ModuleType("win32com")
    client = SimpleNamespace(DispatchEx=lambda _name: excel)
    win32com.client = client  # type: ignore[attr-defined]
    monkeypatch.setitem(sys.modules, "win32com", win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", client)


@pytest.mark.parametrize(
    ("excel", "message"),
    [
        (FakeExcelApp(fail_close=True), "workbook close failed"),
        (FakeExcelApp(fail_quit=True), "excel quit failed"),
    ],
)
def test_win32_backend_wraps_cleanup_failure(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    excel: FakeExcelApp,
    message: str,
) -> None:
    install_fake_win32com(monkeypatch, excel)

    with pytest.raises(ExcelAutomationError) as exc_info:
        Win32ExcelBackend().recalculate_and_export(
            tmp_path / "audit.xlsx",
            tmp_path / "PayrollImport.csv",
        )

    assert str(exc_info.value).startswith(
        "Excel could not recalculate and export the PU tab:"
    )
    assert message in str(exc_info.value)


def test_win32_backend_keeps_operation_failure_when_cleanup_also_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    install_fake_win32com(monkeypatch, FakeExcelApp(fail_save=True, fail_close=True))

    with pytest.raises(ExcelAutomationError) as exc_info:
        Win32ExcelBackend().recalculate_and_export(
            tmp_path / "audit.xlsx",
            tmp_path / "PayrollImport.csv",
        )

    assert "save failed" in str(exc_info.value)
    assert "workbook close failed" not in str(exc_info.value)
