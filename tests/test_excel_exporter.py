from pathlib import Path

from bex_payroll_import.excel_exporter import ExcelExportResult, export_pu_csv_with_excel


class FakeExcelBackend:
    def __init__(self) -> None:
        self.calls: list[tuple[Path, Path]] = []

    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        self.calls.append((workbook_path, csv_path))
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        csv_path.write_text(
            "CD0529143708,E100,DLPTO,REG,8,100,200,L,04/12/2026,0\n",
            encoding="utf-8",
        )


def test_export_pu_csv_with_excel_delegates_to_backend(tmp_path: Path) -> None:
    backend = FakeExcelBackend()
    workbook_path = tmp_path / "audit.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    csv_path = tmp_path / "PayrollImport.csv"

    result = export_pu_csv_with_excel(workbook_path, csv_path, backend=backend)

    assert result == ExcelExportResult(csv_path=csv_path)
    assert backend.calls == [(workbook_path, csv_path)]
    assert csv_path.exists()
