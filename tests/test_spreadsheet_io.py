from pathlib import Path

import pytest

from bex_payroll_import.spreadsheet_io import build_header_map, find_header_row
from tests.helpers import save_workbook


def test_find_header_row_by_required_columns(tmp_path: Path) -> None:
    workbook_path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            ["Report title"],
            ["EECode", "Lastname", "Firstname", "EarnHours"],
            ["100", "Doe", "Jane", 8],
        ],
    )

    header_row = find_header_row(workbook_path, ["EECode", "Lastname", "Firstname"])

    assert header_row.sheet_name == "Time Detail Repor"
    assert header_row.row_number == 2


def test_build_header_map_normalizes_case_and_spaces() -> None:
    headers = [" EECode ", "LastName", "Earn Hours"]

    assert build_header_map(headers) == {
        "eecode": 1,
        "lastname": 2,
        "earn hours": 3,
    }


def test_find_header_row_raises_for_missing_columns(tmp_path: Path) -> None:
    workbook_path = save_workbook(tmp_path / "bad.xlsx", "Sheet1", [["Name"], ["Jane"]])

    with pytest.raises(ValueError, match="Missing required columns"):
        find_header_row(workbook_path, ["EECode", "Lastname"])
