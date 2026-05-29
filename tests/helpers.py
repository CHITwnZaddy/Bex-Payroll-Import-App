from pathlib import Path

from openpyxl import Workbook


def save_workbook(path: Path, sheet_name: str, rows: list[list[object]]) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for row in rows:
        sheet.append(row)
    workbook.save(path)
    return path
