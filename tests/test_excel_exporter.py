from __future__ import annotations

from pathlib import Path
import builtins
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
            "Batch code,Employee code,Department,Pay type,Hours\n"
            "CD0529143708,E100,DLPTO,REG,8\n",
            encoding="utf-8",
        )


class WideExcelBackend:
    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        project_row = ["OLD", "E100", "6", "R", "8", "25009", "011010", "L", "4/12/2026"] + [""] * 34
        project_row[21] = "0.00"
        expense_row = ["OLD", "E102", "1", "EXP REIMB", "", "", "", "0", "4/15/2026"] + [""] * 34
        expense_row[21] = "55.25"
        rows = [
            ["Batch code", "Employee code", "Department", "Pay type", "Hours", "Job", "Phase", "Cost type", "Date"] + [""] * 34,
            project_row,
            ["OLD", "E101", "1", "R", "8", "", "", "", "4/12/2026"] + [""] * 34,
            expense_row,
            ["OLD"] + [""] * 42,
        ]
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
            import csv

            csv.writer(csv_file).writerows(rows)


def test_export_pu_csv_with_excel_delegates_to_backend(tmp_path: Path) -> None:
    backend = FakeExcelBackend()
    workbook_path = tmp_path / "audit.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    csv_path = tmp_path / "nested" / "exports" / "PayrollImport.csv"

    result = export_pu_csv_with_excel(workbook_path, csv_path, backend=backend)

    assert result == ExcelExportResult(csv_path=csv_path)
    assert backend.calls == [(workbook_path, csv_path)]
    assert csv_path.exists()
    import csv

    with csv_path.open(newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    assert len(rows) == 1
    assert len(rows[0]) == 22
    assert rows[0][:5] == ["CD0529143708", "E100", "DLPTO", "REG", "8"]


def test_export_pu_csv_writes_only_employee_rows_and_exactly_22_columns(tmp_path: Path) -> None:
    workbook_path = tmp_path / "audit.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    csv_path = tmp_path / "PayrollImport.csv"

    export_pu_csv_with_excel(
        workbook_path,
        csv_path,
        backend=WideExcelBackend(),
        batch_code="CD0529143708",
    )

    import csv

    with csv_path.open(newline="", encoding="utf-8") as csv_file:
        rows = list(csv.reader(csv_file))
    assert len(rows) == 3
    assert {len(row) for row in rows} == {22}
    assert [row[0] for row in rows] == ["CD0529143708"] * 3
    assert [row[1] for row in rows] == ["E100", "E101", "E102"]
    assert rows[0][21] == ""
    assert rows[2][2:9] == ["1", "EXP REIMB", "", "", "", "", "4/15/2026"]
    assert rows[2][21] == "55.25"


class FakeCsvWorkbook:
    def __init__(self, *, fail_save_as: bool = False) -> None:
        self.fail_save_as = fail_save_as
        self.close_calls: list[bool] = []

    def SaveAs(self, _csv_path: str, FileFormat: int) -> None:
        assert FileFormat == 6
        if self.fail_save_as:
            raise RuntimeError("csv save failed")

    def Close(self, SaveChanges: bool) -> None:
        self.close_calls.append(SaveChanges)
        assert SaveChanges is False


class FakeWorksheet:
    def __init__(self, excel: "FakeExcelApp") -> None:
        self.excel = excel

    def Copy(self) -> None:
        self.excel.ActiveWorkbook = self.excel.csv_workbook


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
        csv_workbook: FakeCsvWorkbook | None = None,
    ) -> None:
        self.ActiveWorkbook = None
        self.csv_workbook = csv_workbook or FakeCsvWorkbook()
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


def test_win32_backend_closes_copied_workbook_when_save_as_fails(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    csv_workbook = FakeCsvWorkbook(fail_save_as=True)
    install_fake_win32com(monkeypatch, FakeExcelApp(csv_workbook=csv_workbook))

    with pytest.raises(ExcelAutomationError) as exc_info:
        Win32ExcelBackend().recalculate_and_export(
            tmp_path / "audit.xlsx",
            tmp_path / "PayrollImport.csv",
        )

    assert "csv save failed" in str(exc_info.value)
    assert csv_workbook.close_calls == [False]


def test_win32_backend_import_error_message(monkeypatch: pytest.MonkeyPatch) -> None:
    original_import = builtins.__import__

    def fake_import(name: str, *args: object, **kwargs: object) -> object:
        if name == "win32com.client":
            raise ImportError("missing win32com")
        return original_import(name, *args, **kwargs)

    monkeypatch.setattr(builtins, "__import__", fake_import)

    with pytest.raises(ExcelAutomationError) as exc_info:
        Win32ExcelBackend().recalculate_and_export(
            Path("audit.xlsx"),
            Path("PayrollImport.csv"),
        )

    assert str(exc_info.value) == "pywin32 is required on Windows for Excel export."
