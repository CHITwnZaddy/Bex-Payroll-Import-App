# BEX Payroll Import App Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a small Windows desktop app that creates a Vista payroll import CSV from a Time Detail Report, an optional Expense Transaction file, and the existing Payroll Upload Worksheet template.

**Architecture:** Keep the workbook as the calculation source for phase 1. Python reads and normalizes source rows, writes them into a copied template workbook, asks desktop Excel to recalculate/export, and writes validation artifacts beside the selected Time Detail Report.

**Tech Stack:** Python 3, Tkinter, openpyxl, pytest, pywin32 on Windows, optional PyInstaller packaging.

---

## Files

Create these files:

- `pyproject.toml`: package metadata, dependencies, pytest config.
- `src/bex_payroll_import/__init__.py`: package marker and version.
- `src/bex_payroll_import/__main__.py`: app entrypoint.
- `src/bex_payroll_import/models.py`: dataclasses and typed constants.
- `src/bex_payroll_import/errors.py`: domain exception types.
- `src/bex_payroll_import/batch.py`: batch code and timestamp helpers.
- `src/bex_payroll_import/config.py`: saved template path load/save.
- `src/bex_payroll_import/paths.py`: output folder and output filename helpers.
- `src/bex_payroll_import/spreadsheet_io.py`: workbook open and header-map helpers.
- `src/bex_payroll_import/key_lookup.py`: `Key` tab employee lookup.
- `src/bex_payroll_import/normalizers.py`: TDR and expense row normalization.
- `src/bex_payroll_import/template_writer.py`: copied template creation, `TDR` writing, formula fill-down.
- `src/bex_payroll_import/excel_exporter.py`: Excel COM recalculation and `PU` CSV export.
- `src/bex_payroll_import/validation.py`: validation result types, validation checks, validation CSV writing.
- `src/bex_payroll_import/runner.py`: full generate workflow.
- `src/bex_payroll_import/ui.py`: Tkinter app.
- `tests/helpers.py`: shared test workbook builders.
- `tests/test_batch.py`
- `tests/test_paths.py`
- `tests/test_config.py`
- `tests/test_spreadsheet_io.py`
- `tests/test_key_lookup.py`
- `tests/test_normalizers.py`
- `tests/test_template_writer.py`
- `tests/test_validation.py`
- `tests/test_runner.py`

Modify these files:

- `README.md`: replace the starter text with operator-facing usage notes.
- `SETUP_GUIDE.md`: replace the starter text with Windows setup, first run, and test-run instructions.
- `.gitignore`: extend only if new build artifacts appear during implementation.

Do not commit real payroll workbooks, source reports, audit workbooks, generated CSV files, or local config.

---

### Task 1: Project setup

**Files:**

- Create: `pyproject.toml`
- Create: `src/bex_payroll_import/__init__.py`
- Create: `src/bex_payroll_import/__main__.py`
- Modify: `src/.gitkeep`
- Modify: `tests/.gitkeep`

- [ ] **Step 1: Write the package config**

Create `pyproject.toml`:

```toml
[build-system]
requires = ["setuptools>=69", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "bex-payroll-import"
version = "0.1.0"
description = "Windows desktop app for creating BEX Vista payroll import CSV files."
requires-python = ">=3.10"
dependencies = [
  "openpyxl>=3.1.2",
]

[project.optional-dependencies]
dev = [
  "pytest>=8.0",
]
windows = [
  "pywin32>=306",
  "pyinstaller>=6.0",
]

[project.scripts]
bex-payroll-import = "bex_payroll_import.__main__:main"

[tool.setuptools.packages.find]
where = ["src"]

[tool.pytest.ini_options]
testpaths = ["tests"]
pythonpath = ["src"]
```

- [ ] **Step 2: Add package marker**

Create `src/bex_payroll_import/__init__.py`:

```python
"""BEX payroll import app."""

__version__ = "0.1.0"
```

- [ ] **Step 3: Add the entrypoint**

Create `src/bex_payroll_import/__main__.py`:

```python
from bex_payroll_import.ui import main


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: Remove placeholder files after real package files exist**

Run:

```bash
rm -f src/.gitkeep tests/.gitkeep
```

Expected: command exits with no output.

- [ ] **Step 5: Install editable dev dependencies**

Run:

```bash
python3 -m pip install -e ".[dev]"
```

Expected: package installs without errors.

On Windows, run this instead:

```powershell
py -m pip install -e ".[dev,windows]"
```

- [ ] **Step 6: Run empty test suite**

Run:

```bash
python3 -m pytest -q
```

Expected: pytest exits with code `5` because no tests exist yet.

- [ ] **Step 7: Commit**

```bash
git add pyproject.toml src tests
git commit -m "chore: set up python package"
```

---

### Task 2: Core models and validation report primitives

**Files:**

- Create: `src/bex_payroll_import/models.py`
- Create: `src/bex_payroll_import/errors.py`
- Create: `src/bex_payroll_import/validation.py`
- Test: `tests/test_validation.py`

- [ ] **Step 1: Write failing tests for validation results**

Create `tests/test_validation.py`:

```python
from pathlib import Path

from bex_payroll_import.validation import ValidationMessage, ValidationReport, write_validation_csv


def test_validation_report_blocks_on_errors() -> None:
    report = ValidationReport()
    report.add_error("Missing column", "Add the Date column and rerun.", row_number=2)
    report.add_warning("Expense skipped", "User confirmed no expense file.")

    assert report.has_errors is True
    assert report.error_count == 1
    assert report.warning_count == 1


def test_write_validation_csv(tmp_path: Path) -> None:
    report = ValidationReport()
    report.messages.append(
        ValidationMessage(
            severity="ERROR",
            message="Missing column",
            suggested_fix="Add the Date column and rerun.",
            row_number=2,
            source="Time Detail Report",
        )
    )

    output = tmp_path / "ValidationErrors_CD0529143708.csv"
    write_validation_csv(report, output)

    text = output.read_text(encoding="utf-8")
    assert "severity,row_number,source,message,suggested_fix" in text
    assert "ERROR,2,Time Detail Report,Missing column,Add the Date column and rerun." in text
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_validation.py -q
```

Expected: FAIL with `ModuleNotFoundError` or import errors for `bex_payroll_import.validation`.

- [ ] **Step 3: Add model types**

Create `src/bex_payroll_import/models.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path


TDR_SHEET_NAME = "TDR"
PU_SHEET_NAME = "PU"
KEY_SHEET_NAME = "Key"

FINAL_OUTPUT_COLUMNS = [
    "Batch code",
    "Employee code",
    "Department",
    "Pay type",
    "Hours",
    "Job",
    "Phase",
    "Cost type",
    "Date",
    "$",
]


@dataclass(frozen=True)
class EmployeeIdentity:
    first_name: str
    last_name: str


@dataclass(frozen=True)
class NormalizedPayrollRow:
    batch_code: str
    employee_code: str
    department: str
    pay_type: str
    hours: Decimal
    job: str
    phase: str
    cost_type: str
    work_date: date
    dollars: Decimal | None
    source: str
    source_row_number: int


@dataclass(frozen=True)
class RunInputs:
    template_path: Path
    tdr_path: Path
    expense_path: Path | None
    expense_skipped: bool


@dataclass(frozen=True)
class RunOutputs:
    batch_code: str
    output_dir: Path
    audit_workbook_path: Path
    payroll_csv_path: Path | None
    validation_path: Path
    technical_log_path: Path | None


@dataclass(frozen=True)
class RunClock:
    now: datetime
```

- [ ] **Step 4: Add error types**

Create `src/bex_payroll_import/errors.py`:

```python
from __future__ import annotations


class PayrollImportError(Exception):
    """Base error for expected payroll import failures."""


class ValidationFailedError(PayrollImportError):
    """Raised when validation errors block CSV generation."""


class ExcelAutomationError(PayrollImportError):
    """Raised when desktop Excel cannot recalculate or export."""
```

- [ ] **Step 5: Add validation primitives**

Create `src/bex_payroll_import/validation.py`:

```python
from __future__ import annotations

import csv
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


Severity = Literal["ERROR", "WARNING"]


@dataclass(frozen=True)
class ValidationMessage:
    severity: Severity
    message: str
    suggested_fix: str
    row_number: int | None = None
    source: str = ""


@dataclass
class ValidationReport:
    messages: list[ValidationMessage] = field(default_factory=list)

    @property
    def has_errors(self) -> bool:
        return any(message.severity == "ERROR" for message in self.messages)

    @property
    def error_count(self) -> int:
        return sum(1 for message in self.messages if message.severity == "ERROR")

    @property
    def warning_count(self) -> int:
        return sum(1 for message in self.messages if message.severity == "WARNING")

    def add_error(
        self,
        message: str,
        suggested_fix: str,
        row_number: int | None = None,
        source: str = "",
    ) -> None:
        self.messages.append(
            ValidationMessage(
                severity="ERROR",
                message=message,
                suggested_fix=suggested_fix,
                row_number=row_number,
                source=source,
            )
        )

    def add_warning(
        self,
        message: str,
        suggested_fix: str,
        row_number: int | None = None,
        source: str = "",
    ) -> None:
        self.messages.append(
            ValidationMessage(
                severity="WARNING",
                message=message,
                suggested_fix=suggested_fix,
                row_number=row_number,
                source=source,
            )
        )


def write_validation_csv(report: ValidationReport, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=["severity", "row_number", "source", "message", "suggested_fix"],
        )
        writer.writeheader()
        for message in report.messages:
            writer.writerow(
                {
                    "severity": message.severity,
                    "row_number": message.row_number or "",
                    "source": message.source,
                    "message": message.message,
                    "suggested_fix": message.suggested_fix,
                }
            )
```

- [ ] **Step 6: Run test to verify it passes**

Run:

```bash
python3 -m pytest tests/test_validation.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add src/bex_payroll_import/models.py src/bex_payroll_import/errors.py src/bex_payroll_import/validation.py tests/test_validation.py
git commit -m "feat: add validation report primitives"
```

---

### Task 3: Config, batch code, and output paths

**Files:**

- Create: `src/bex_payroll_import/config.py`
- Create: `src/bex_payroll_import/batch.py`
- Create: `src/bex_payroll_import/paths.py`
- Test: `tests/test_config.py`
- Test: `tests/test_batch.py`
- Test: `tests/test_paths.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_batch.py`:

```python
from datetime import datetime

from bex_payroll_import.batch import build_batch_code


def test_build_batch_code_uses_cd_mm_dd_hh_mm_ss() -> None:
    now = datetime(2026, 5, 29, 14, 37, 8)

    assert build_batch_code(now) == "CD0529143708"
```

Create `tests/test_paths.py`:

```python
from datetime import datetime
from pathlib import Path

from bex_payroll_import.paths import build_output_paths


def test_build_output_paths_next_to_tdr(tmp_path: Path) -> None:
    tdr_path = tmp_path / "BEX Payroll - #05" / "tdr.xlsx"
    now = datetime(2026, 5, 29, 14, 37, 8)

    outputs = build_output_paths(tdr_path, "CD0529143708", now)

    assert outputs.output_dir == tmp_path / "BEX Payroll - #05" / "Payroll Import Outputs" / "2026-05-29"
    assert outputs.payroll_csv_path.name == "PayrollImport_CD0529143708_2026-05-29_14-37-08.csv"
    assert outputs.audit_workbook_path.name == "PayrollUpload_Audit_CD0529143708_2026-05-29_14-37-08.xlsx"
    assert outputs.validation_path.name == "ValidationReport_CD0529143708_2026-05-29_14-37-08.csv"
```

Create `tests/test_config.py`:

```python
from pathlib import Path

from bex_payroll_import.config import AppConfig, load_config, save_config


def test_save_and_load_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    template_path = tmp_path / "Payroll Upload Worksheet.xlsx"

    save_config(AppConfig(template_path=template_path), config_path)

    assert load_config(config_path) == AppConfig(template_path=template_path)


def test_load_missing_config_returns_empty_config(tmp_path: Path) -> None:
    assert load_config(tmp_path / "missing.json") == AppConfig(template_path=None)
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python3 -m pytest tests/test_batch.py tests/test_paths.py tests/test_config.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add batch helper**

Create `src/bex_payroll_import/batch.py`:

```python
from __future__ import annotations

from datetime import datetime


def build_batch_code(now: datetime) -> str:
    return f"CD{now:%m%d%H%M%S}"
```

- [ ] **Step 4: Add output path helper**

Create `src/bex_payroll_import/paths.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class PlannedOutputPaths:
    output_dir: Path
    audit_workbook_path: Path
    payroll_csv_path: Path
    validation_path: Path
    validation_errors_path: Path
    technical_log_path: Path


def build_output_paths(tdr_path: Path, batch_code: str, now: datetime) -> PlannedOutputPaths:
    timestamp = f"{now:%Y-%m-%d_%H-%M-%S}"
    output_dir = tdr_path.parent / "Payroll Import Outputs" / f"{now:%Y-%m-%d}"
    return PlannedOutputPaths(
        output_dir=output_dir,
        audit_workbook_path=output_dir / f"PayrollUpload_Audit_{batch_code}_{timestamp}.xlsx",
        payroll_csv_path=output_dir / f"PayrollImport_{batch_code}_{timestamp}.csv",
        validation_path=output_dir / f"ValidationReport_{batch_code}_{timestamp}.csv",
        validation_errors_path=output_dir / f"ValidationErrors_{batch_code}_{timestamp}.csv",
        technical_log_path=output_dir / f"TechnicalLog_{batch_code}_{timestamp}.log",
    )
```

- [ ] **Step 5: Add config persistence**

Create `src/bex_payroll_import/config.py`:

```python
from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path


APP_CONFIG_DIR = Path.home() / ".bex_payroll_import"
APP_CONFIG_PATH = APP_CONFIG_DIR / "config.json"


@dataclass(frozen=True)
class AppConfig:
    template_path: Path | None


def load_config(config_path: Path = APP_CONFIG_PATH) -> AppConfig:
    if not config_path.exists():
        return AppConfig(template_path=None)
    data = json.loads(config_path.read_text(encoding="utf-8"))
    raw_template_path = data.get("template_path")
    return AppConfig(template_path=Path(raw_template_path) if raw_template_path else None)


def save_config(config: AppConfig, config_path: Path = APP_CONFIG_PATH) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "template_path": str(config.template_path) if config.template_path else None,
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
```

- [ ] **Step 6: Run tests to verify they pass**

Run:

```bash
python3 -m pytest tests/test_batch.py tests/test_paths.py tests/test_config.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit**

```bash
git add src/bex_payroll_import/batch.py src/bex_payroll_import/paths.py src/bex_payroll_import/config.py tests/test_batch.py tests/test_paths.py tests/test_config.py
git commit -m "feat: add run config and output paths"
```

---

### Task 4: Spreadsheet header reading

**Files:**

- Create: `src/bex_payroll_import/spreadsheet_io.py`
- Test: `tests/helpers.py`
- Test: `tests/test_spreadsheet_io.py`

- [ ] **Step 1: Write failing tests and test helpers**

Create `tests/helpers.py`:

```python
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
```

Create `tests/test_spreadsheet_io.py`:

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```bash
python3 -m pytest tests/test_spreadsheet_io.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add spreadsheet helpers**

Create `src/bex_payroll_import/spreadsheet_io.py`:

```python
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook


@dataclass(frozen=True)
class HeaderRow:
    sheet_name: str
    row_number: int
    header_map: dict[str, int]


def normalize_header(value: object) -> str:
    return " ".join(str(value or "").strip().lower().split())


def build_header_map(headers: list[object]) -> dict[str, int]:
    return {
        normalize_header(header): index
        for index, header in enumerate(headers, start=1)
        if normalize_header(header)
    }


def find_header_row(workbook_path: Path, required_columns: list[str]) -> HeaderRow:
    workbook = load_workbook(workbook_path, read_only=True, data_only=False)
    normalized_required = [normalize_header(column) for column in required_columns]
    try:
        for sheet in workbook.worksheets:
            for row_number, row in enumerate(sheet.iter_rows(values_only=True), start=1):
                header_map = build_header_map(list(row))
                if all(column in header_map for column in normalized_required):
                    return HeaderRow(sheet.name, row_number, header_map)
        missing = ", ".join(required_columns)
        raise ValueError(f"Missing required columns: {missing}")
    finally:
        workbook.close()
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
python3 -m pytest tests/test_spreadsheet_io.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/spreadsheet_io.py tests/helpers.py tests/test_spreadsheet_io.py
git commit -m "feat: add spreadsheet header detection"
```

---

### Task 5: Key lookup and expense employee parsing

**Files:**

- Create: `src/bex_payroll_import/key_lookup.py`
- Test: `tests/test_key_lookup.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_key_lookup.py`:

```python
from pathlib import Path

import pytest
from openpyxl import Workbook

from bex_payroll_import.key_lookup import EmployeeLookup, load_employee_lookup, parse_expense_employee


def make_key_workbook(path: Path) -> Path:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Key"
    sheet.append(["Employee code", "Last Name", "First Name"])
    sheet.append(["E100", "Doe", "Jane"])
    workbook.save(path)
    return path


def test_parse_expense_employee_last_comma_first() -> None:
    identity = parse_expense_employee("DOE, JANE")

    assert identity.last_name == "Doe"
    assert identity.first_name == "Jane"


def test_parse_expense_employee_rejects_bad_format() -> None:
    with pytest.raises(ValueError, match="Expected employee format"):
        parse_expense_employee("Jane Doe")


def test_lookup_employee_code_from_key_tab(tmp_path: Path) -> None:
    lookup = load_employee_lookup(make_key_workbook(tmp_path / "template.xlsx"))

    assert lookup.resolve_code("jane", "doe") == "E100"


def test_lookup_raises_for_unmatched_employee() -> None:
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        lookup.resolve_code("John", "Doe")
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_key_lookup.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add key lookup implementation**

Create `src/bex_payroll_import/key_lookup.py`:

```python
from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from bex_payroll_import.models import EmployeeIdentity, KEY_SHEET_NAME
from bex_payroll_import.spreadsheet_io import build_header_map, normalize_header


class EmployeeLookup:
    def __init__(self, code_by_name: dict[tuple[str, str], str]) -> None:
        self._code_by_name = code_by_name

    def resolve_code(self, first_name: str, last_name: str) -> str:
        key = (normalize_name(first_name), normalize_name(last_name))
        try:
            return self._code_by_name[key]
        except KeyError as exc:
            raise KeyError(f"No employee code found for {first_name} {last_name}") from exc


def normalize_name(value: str) -> str:
    return " ".join(value.strip().lower().split())


def display_name(value: str) -> str:
    return " ".join(piece.capitalize() for piece in value.strip().split())


def parse_expense_employee(raw_employee: object) -> EmployeeIdentity:
    raw = str(raw_employee or "").strip()
    if "," not in raw:
        raise ValueError(f"Expected employee format 'LAST, FIRST', got '{raw}'")
    last_name, first_name = [part.strip() for part in raw.split(",", 1)]
    if not first_name or not last_name:
        raise ValueError(f"Expected employee format 'LAST, FIRST', got '{raw}'")
    return EmployeeIdentity(first_name=display_name(first_name), last_name=display_name(last_name))


def load_employee_lookup(template_path: Path) -> EmployeeLookup:
    workbook = load_workbook(template_path, read_only=True, data_only=False)
    try:
        if KEY_SHEET_NAME not in workbook.sheetnames:
            raise ValueError("Template is missing the Key tab.")
        sheet = workbook[KEY_SHEET_NAME]
        header_row = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))
        header_map = build_header_map(list(header_row))
        code_col = find_first_header(header_map, ["employee code", "eecode", "ee code"])
        last_col = find_first_header(header_map, ["last name", "lastname"])
        first_col = find_first_header(header_map, ["first name", "firstname"])

        code_by_name: dict[tuple[str, str], str] = {}
        for row in sheet.iter_rows(min_row=2, values_only=True):
            code = cell_value(row, code_col)
            first_name = cell_value(row, first_col)
            last_name = cell_value(row, last_col)
            if code and first_name and last_name:
                code_by_name[(normalize_name(first_name), normalize_name(last_name))] = str(code).strip()
        return EmployeeLookup(code_by_name)
    finally:
        workbook.close()


def find_first_header(header_map: dict[str, int], options: list[str]) -> int:
    for option in options:
        normalized = normalize_header(option)
        if normalized in header_map:
            return header_map[normalized]
    raise ValueError(f"Missing Key tab column. Tried: {', '.join(options)}")


def cell_value(row: tuple[object, ...], one_based_index: int) -> object:
    return row[one_based_index - 1] if one_based_index <= len(row) else None
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
python3 -m pytest tests/test_key_lookup.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/key_lookup.py tests/test_key_lookup.py
git commit -m "feat: add employee key lookup"
```

---

### Task 6: TDR and expense normalization

**Files:**

- Create: `src/bex_payroll_import/normalizers.py`
- Test: `tests/test_normalizers.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_normalizers.py`:

```python
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from bex_payroll_import.key_lookup import EmployeeLookup
from bex_payroll_import.normalizers import normalize_expense_file, normalize_tdr_file
from tests.helpers import save_workbook


def test_normalize_tdr_file_reads_required_values(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "tdr.xlsx",
        "Time Detail Repor",
        [
            ["EECode", "Lastname", "Firstname", "HomeDepartment", "Pay Class", "Badge", "Date", "InPunchTime", "OutPunchTime", "Department", "EarnCode", "EarnHours", "Dollars", "Job", "Phase", "Cost Type"],
            ["E100", "Doe", "Jane", "DLPTO", "REG", "1", date(2026, 4, 12), "", "", "OPS", "REG", 8, 0, "100", "200", "L"],
        ],
    )

    rows = normalize_tdr_file(path, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].batch_code == "CD0529143708"
    assert rows[0].employee_code == "E100"
    assert rows[0].hours == Decimal("8")
    assert rows[0].work_date == date(2026, 4, 12)


def test_normalize_expense_file_applies_defaults(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [
            ["Employee", "Expense Date", "Paid Amount"],
            ["DOE, JANE", datetime(2026, 4, 15, 8, 30), 123.45],
        ],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    rows = normalize_expense_file(path, lookup, "CD0529143708")

    assert len(rows) == 1
    assert rows[0].employee_code == "E100"
    assert rows[0].department == "DLPTO"
    assert rows[0].pay_type == "EXP REIM"
    assert rows[0].hours == Decimal("0")
    assert rows[0].job == ""
    assert rows[0].phase == ""
    assert rows[0].cost_type == "L"
    assert rows[0].dollars == Decimal("123.45")


def test_normalize_expense_file_raises_for_unmatched_employee(tmp_path: Path) -> None:
    path = save_workbook(
        tmp_path / "expense.xlsx",
        "Expense Transacti",
        [["Employee", "Expense Date", "Paid Amount"], ["DOE, JOHN", date(2026, 4, 15), 50]],
    )
    lookup = EmployeeLookup({("jane", "doe"): "E100"})

    with pytest.raises(KeyError, match="No employee code found"):
        normalize_expense_file(path, lookup, "CD0529143708")
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_normalizers.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add normalizer implementation**

Create `src/bex_payroll_import/normalizers.py`:

```python
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

from openpyxl import load_workbook

from bex_payroll_import.key_lookup import EmployeeLookup, parse_expense_employee
from bex_payroll_import.models import NormalizedPayrollRow
from bex_payroll_import.spreadsheet_io import find_header_row, normalize_header


TDR_REQUIRED_COLUMNS = ["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"]
EXPENSE_REQUIRED_COLUMNS = ["Employee", "Expense Date", "Paid Amount"]


def normalize_tdr_file(tdr_path: Path, batch_code: str) -> list[NormalizedPayrollRow]:
    header_row = find_header_row(tdr_path, TDR_REQUIRED_COLUMNS)
    workbook = load_workbook(tdr_path, read_only=True, data_only=True)
    rows: list[NormalizedPayrollRow] = []
    try:
        sheet = workbook[header_row.sheet_name]
        for row_number, row in enumerate(sheet.iter_rows(min_row=header_row.row_number + 1, values_only=True), start=header_row.row_number + 1):
            employee_code = text_value(row, header_row.header_map, "EECode")
            if not employee_code:
                continue
            rows.append(
                NormalizedPayrollRow(
                    batch_code=batch_code,
                    employee_code=employee_code,
                    department=text_value(row, header_row.header_map, "Department"),
                    pay_type=text_value(row, header_row.header_map, "EarnCode"),
                    hours=decimal_value(cell(row, header_row.header_map, "EarnHours")),
                    job=text_value(row, header_row.header_map, "Job"),
                    phase=text_value(row, header_row.header_map, "Phase"),
                    cost_type=text_value(row, header_row.header_map, "Cost Type"),
                    work_date=date_value(cell(row, header_row.header_map, "Date")),
                    dollars=decimal_or_none(cell(row, header_row.header_map, "Dollars")),
                    source="Time Detail Report",
                    source_row_number=row_number,
                )
            )
    finally:
        workbook.close()
    return rows


def normalize_expense_file(expense_path: Path, lookup: EmployeeLookup, batch_code: str) -> list[NormalizedPayrollRow]:
    header_row = find_header_row(expense_path, EXPENSE_REQUIRED_COLUMNS)
    workbook = load_workbook(expense_path, read_only=True, data_only=True)
    rows: list[NormalizedPayrollRow] = []
    try:
        sheet = workbook[header_row.sheet_name]
        for row_number, row in enumerate(sheet.iter_rows(min_row=header_row.row_number + 1, values_only=True), start=header_row.row_number + 1):
            raw_employee = cell(row, header_row.header_map, "Employee")
            if not raw_employee:
                continue
            identity = parse_expense_employee(raw_employee)
            rows.append(
                NormalizedPayrollRow(
                    batch_code=batch_code,
                    employee_code=lookup.resolve_code(identity.first_name, identity.last_name),
                    department="DLPTO",
                    pay_type="EXP REIM",
                    hours=Decimal("0"),
                    job="",
                    phase="",
                    cost_type="L",
                    work_date=date_value(cell(row, header_row.header_map, "Expense Date")),
                    dollars=decimal_value(cell(row, header_row.header_map, "Paid Amount")),
                    source="Expense Transaction",
                    source_row_number=row_number,
                )
            )
    finally:
        workbook.close()
    return rows


def cell(row: tuple[object, ...], header_map: dict[str, int], column_name: str) -> object:
    index = header_map.get(normalize_header(column_name))
    if index is None or index > len(row):
        return None
    return row[index - 1]


def text_value(row: tuple[object, ...], header_map: dict[str, int], column_name: str) -> str:
    value = cell(row, header_map, column_name)
    return "" if value is None else str(value).strip()


def decimal_value(value: object) -> Decimal:
    if value is None or value == "":
        return Decimal("0")
    return Decimal(str(value)).quantize(Decimal("0.01")).normalize()


def decimal_or_none(value: object) -> Decimal | None:
    if value is None or value == "":
        return None
    return decimal_value(value)


def date_value(value: object) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        return datetime.strptime(value.strip(), "%m/%d/%Y").date()
    raise ValueError(f"Expected date value, got {value!r}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
python3 -m pytest tests/test_normalizers.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/normalizers.py tests/test_normalizers.py
git commit -m "feat: normalize payroll source rows"
```

---

### Task 7: Template copy, TDR writing, and formula fill-down

**Files:**

- Create: `src/bex_payroll_import/template_writer.py`
- Test: `tests/test_template_writer.py`

- [ ] **Step 1: Write failing tests**

Create `tests/test_template_writer.py`:

```python
from datetime import date
from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook, load_workbook

from bex_payroll_import.models import NormalizedPayrollRow
from bex_payroll_import.template_writer import copy_template_to_audit, write_tdr_rows


def make_template(path: Path) -> Path:
    workbook = Workbook()
    pu = workbook.active
    pu.title = "PU"
    tdr = workbook.create_sheet("TDR")
    workbook.create_sheet("Key")
    tdr.append(["Batch code", "EECode", "Lastname", "Firstname", "HomeDepartment", "Pay Class", "Badge", "Date", "Text to Col", "InPunchTime", "OutPunchTime", "Department", "EarnCode", "Hours", "Dollars", "Employee Approved", "Supervisor Approved", "Tax Profile", "Home Department Desc", "Dist Department Desc", "Job", "Employee code", "Phase", "Cost Code", "Cost type", "Home Department", "Department", "Phase Adj", "Phase", "Pay type"])
    tdr.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=B2", "=W2", "", "=Y2", "", "=L2", "", "=W2", "=M2"])
    workbook.save(path)
    return path


def test_copy_template_to_audit_preserves_original(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    audit_path = tmp_path / "out" / "audit.xlsx"

    copy_template_to_audit(template_path, audit_path)

    assert template_path.exists()
    assert audit_path.exists()


def test_write_tdr_rows_adds_rows_and_copies_formula_columns(tmp_path: Path) -> None:
    audit_path = tmp_path / "audit.xlsx"
    make_template(audit_path)
    rows = [
        NormalizedPayrollRow("CD0529143708", "E100", "DLPTO", "REG", Decimal("8"), "100", "200", "L", date(2026, 4, 12), Decimal("0"), "Time Detail Report", 2),
        NormalizedPayrollRow("CD0529143708", "E101", "DLPTO", "EXP REIM", Decimal("0"), "", "", "L", date(2026, 4, 15), Decimal("55.25"), "Expense Transaction", 2),
    ]

    write_tdr_rows(audit_path, rows)

    workbook = load_workbook(audit_path, data_only=False)
    sheet = workbook["TDR"]
    assert sheet["A2"].value == "CD0529143708"
    assert sheet["B2"].value == "E100"
    assert sheet["H2"].value == date(2026, 4, 12)
    assert sheet["N3"].value == 0
    assert sheet["O3"].value == 55.25
    assert sheet["V3"].value == "=B3"
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_template_writer.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add template writer implementation**

Create `src/bex_payroll_import/template_writer.py`:

```python
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
    if TDR_SHEET_NAME not in workbook.sheetnames:
        raise ValueError("Template is missing the TDR tab.")
    sheet = workbook[TDR_SHEET_NAME]
    header_row = 1
    data_start_row = 2
    clear_existing_source_rows(sheet, data_start_row)

    formula_source_row = find_formula_source_row(sheet, data_start_row)
    for offset, payroll_row in enumerate(rows):
        target_row = data_start_row + offset
        write_tdr_source_cells(sheet, target_row, payroll_row)
        if formula_source_row:
            copy_formula_cells(sheet, formula_source_row, target_row)

    workbook.save(audit_workbook_path)


def clear_existing_source_rows(sheet, data_start_row: int) -> None:
    for row in range(data_start_row, sheet.max_row + 1):
        for column in range(1, 16):
            sheet.cell(row=row, column=column).value = None


def find_formula_source_row(sheet, data_start_row: int) -> int | None:
    for row in range(data_start_row, sheet.max_row + 1):
        for column in range(22, 31):
            value = sheet.cell(row=row, column=column).value
            if isinstance(value, str) and value.startswith("="):
                return row
    return None


def write_tdr_source_cells(sheet, row_number: int, payroll_row: NormalizedPayrollRow) -> None:
    sheet.cell(row=row_number, column=1).value = payroll_row.batch_code
    sheet.cell(row=row_number, column=2).value = payroll_row.employee_code
    sheet.cell(row=row_number, column=8).value = payroll_row.work_date
    sheet.cell(row=row_number, column=12).value = payroll_row.department
    sheet.cell(row=row_number, column=13).value = payroll_row.pay_type
    sheet.cell(row=row_number, column=14).value = float(payroll_row.hours)
    sheet.cell(row=row_number, column=15).value = float(payroll_row.dollars or 0)
    sheet.cell(row=row_number, column=21).value = payroll_row.job
    sheet.cell(row=row_number, column=23).value = payroll_row.phase
    sheet.cell(row=row_number, column=25).value = payroll_row.cost_type


def copy_formula_cells(sheet, formula_source_row: int, target_row: int) -> None:
    for column in range(22, 31):
        source_cell = sheet.cell(row=formula_source_row, column=column)
        target_cell = sheet.cell(row=target_row, column=column)
        if isinstance(source_cell.value, str) and source_cell.value.startswith("="):
            target_cell.value = Translator(source_cell.value, origin=source_cell.coordinate).translate_formula(target_cell.coordinate)
```

- [ ] **Step 4: Run tests to verify they pass**

Run:

```bash
python3 -m pytest tests/test_template_writer.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/template_writer.py tests/test_template_writer.py
git commit -m "feat: write payroll rows into template"
```

---

### Task 8: Excel recalculation and CSV export boundary

**Files:**

- Create: `src/bex_payroll_import/excel_exporter.py`
- Test: `tests/test_excel_exporter.py`

- [ ] **Step 1: Write failing tests for path and interface behavior**

Create `tests/test_excel_exporter.py`:

```python
from pathlib import Path

from bex_payroll_import.excel_exporter import ExcelExportResult, export_pu_csv_with_excel


class FakeExcelBackend:
    def __init__(self) -> None:
        self.calls: list[tuple[Path, Path]] = []

    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        self.calls.append((workbook_path, csv_path))
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        csv_path.write_text("CD0529143708,E100,DLPTO,REG,8,100,200,L,04/12/2026,0\n", encoding="utf-8")


def test_export_pu_csv_with_excel_delegates_to_backend(tmp_path: Path) -> None:
    backend = FakeExcelBackend()
    workbook_path = tmp_path / "audit.xlsx"
    workbook_path.write_text("fake", encoding="utf-8")
    csv_path = tmp_path / "PayrollImport.csv"

    result = export_pu_csv_with_excel(workbook_path, csv_path, backend=backend)

    assert result == ExcelExportResult(csv_path=csv_path)
    assert backend.calls == [(workbook_path, csv_path)]
    assert csv_path.exists()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_excel_exporter.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add Excel exporter implementation**

Create `src/bex_payroll_import/excel_exporter.py`:

```python
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
            raise ExcelAutomationError("pywin32 is required on Windows for Excel export.") from exc

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
            raise ExcelAutomationError(f"Excel could not recalculate and export the PU tab: {exc}") from exc
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
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
python3 -m pytest tests/test_excel_exporter.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/excel_exporter.py tests/test_excel_exporter.py
git commit -m "feat: add excel export boundary"
```

---

### Task 9: Post-recalculation CSV validation

**Files:**

- Modify: `src/bex_payroll_import/validation.py`
- Test: `tests/test_validation.py`

- [ ] **Step 1: Add failing tests for final CSV validation**

Append to `tests/test_validation.py`:

```python
from bex_payroll_import.validation import validate_final_csv


def test_validate_final_csv_blocks_header_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text("Batch code,Employee code,Department\nCD0529143708,E100,DLPTO\n", encoding="utf-8")
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert "Final CSV still contains a header row." in [message.message for message in report.messages]


def test_validate_final_csv_blocks_empty_file(tmp_path: Path) -> None:
    csv_path = tmp_path / "PayrollImport.csv"
    csv_path.write_text("", encoding="utf-8")
    report = ValidationReport()

    validate_final_csv(csv_path, report)

    assert report.has_errors is True
    assert "Final PU output has no valid rows." in [message.message for message in report.messages]
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_validation.py -q
```

Expected: FAIL because `validate_final_csv` does not exist.

- [ ] **Step 3: Add final CSV validation**

Append to `src/bex_payroll_import/validation.py`:

```python

def validate_final_csv(csv_path: Path, report: ValidationReport) -> None:
    if not csv_path.exists() or not csv_path.read_text(encoding="utf-8").strip():
        report.add_error(
            "Final PU output has no valid rows.",
            "Confirm the Time Detail Report has payroll rows and rerun.",
            source="PU CSV",
        )
        return

    first_line = csv_path.read_text(encoding="utf-8").splitlines()[0]
    first_cells = [cell.strip().lower() for cell in first_line.split(",")]
    if "batch code" in first_cells or "employee code" in first_cells:
        report.add_error(
            "Final CSV still contains a header row.",
            "Export the PU data rows without headers and rerun.",
            source="PU CSV",
        )
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
python3 -m pytest tests/test_validation.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/bex_payroll_import/validation.py tests/test_validation.py
git commit -m "feat: validate final payroll csv"
```

---

### Task 10: Full run orchestration

**Files:**

- Create: `src/bex_payroll_import/runner.py`
- Test: `tests/test_runner.py`

- [ ] **Step 1: Write failing orchestration tests**

Create `tests/test_runner.py`:

```python
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

from bex_payroll_import.models import RunInputs
from bex_payroll_import.runner import run_payroll_import
from tests.helpers import save_workbook


class FakeExcelBackend:
    def recalculate_and_export(self, workbook_path: Path, csv_path: Path) -> None:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        csv_path.write_text("CD0529143708,E100,DLPTO,REG,8,100,200,L,04/12/2026,0\n", encoding="utf-8")


def make_template(path: Path) -> Path:
    workbook = Workbook()
    pu = workbook.active
    pu.title = "PU"
    workbook.create_sheet("Key")
    tdr = workbook.create_sheet("TDR")
    tdr.append(["Batch code", "EECode", "Lastname", "Firstname", "HomeDepartment", "Pay Class", "Badge", "Date", "Text to Col", "InPunchTime", "OutPunchTime", "Department", "EarnCode", "Hours", "Dollars", "Employee Approved", "Supervisor Approved", "Tax Profile", "Home Department Desc", "Dist Department Desc", "Job", "Employee code", "Phase", "Cost Code", "Cost type", "Home Department", "Department", "Phase Adj", "Phase", "Pay type"])
    tdr.append(["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "=B2", "=W2", "", "=Y2", "", "=L2", "", "=W2", "=M2"])
    key = workbook["Key"]
    key.append(["Employee code", "Last Name", "First Name"])
    key.append(["E100", "Doe", "Jane"])
    workbook.save(path)
    return path


def test_run_payroll_import_success_with_expenses(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    tdr_path = save_workbook(
        tmp_path / "source" / "tdr.xlsx",
        "Time Detail Repor",
        [["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"], ["E100", "Doe", "Jane", "04/12/2026", "DLPTO", "REG", 8]],
    )
    expense_path = save_workbook(
        tmp_path / "source" / "expense.xlsx",
        "Expense Transacti",
        [["Employee", "Expense Date", "Paid Amount"], ["DOE, JANE", "04/15/2026", 55]],
    )

    outputs = run_payroll_import(
        RunInputs(template_path=template_path, tdr_path=tdr_path, expense_path=expense_path, expense_skipped=False),
        now=datetime(2026, 5, 29, 14, 37, 8),
        excel_backend=FakeExcelBackend(),
    )

    assert outputs.batch_code == "CD0529143708"
    assert outputs.payroll_csv_path is not None
    assert outputs.payroll_csv_path.exists()
    assert outputs.audit_workbook_path.exists()
    assert outputs.validation_path.exists()


def test_run_payroll_import_hard_stops_when_expense_missing_and_not_confirmed(tmp_path: Path) -> None:
    template_path = make_template(tmp_path / "template.xlsx")
    tdr_path = save_workbook(tmp_path / "source" / "tdr.xlsx", "Time Detail Repor", [["EECode", "Lastname", "Firstname", "Date", "Department", "EarnCode", "EarnHours"]])

    outputs = run_payroll_import(
        RunInputs(template_path=template_path, tdr_path=tdr_path, expense_path=None, expense_skipped=False),
        now=datetime(2026, 5, 29, 14, 37, 8),
        excel_backend=FakeExcelBackend(),
    )

    assert outputs.payroll_csv_path is None
    assert outputs.validation_path.name.startswith("ValidationErrors_")
    assert outputs.validation_path.exists()
```

- [ ] **Step 2: Run test to verify it fails**

Run:

```bash
python3 -m pytest tests/test_runner.py -q
```

Expected: FAIL with import errors.

- [ ] **Step 3: Add runner implementation**

Create `src/bex_payroll_import/runner.py`:

```python
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from bex_payroll_import.batch import build_batch_code
from bex_payroll_import.excel_exporter import ExcelBackend, export_pu_csv_with_excel
from bex_payroll_import.key_lookup import load_employee_lookup
from bex_payroll_import.models import RunInputs, RunOutputs
from bex_payroll_import.normalizers import normalize_expense_file, normalize_tdr_file
from bex_payroll_import.paths import build_output_paths
from bex_payroll_import.template_writer import copy_template_to_audit, write_tdr_rows
from bex_payroll_import.validation import ValidationReport, validate_final_csv, write_validation_csv


def run_payroll_import(
    inputs: RunInputs,
    now: datetime | None = None,
    excel_backend: ExcelBackend | None = None,
) -> RunOutputs:
    run_now = now or datetime.now()
    batch_code = build_batch_code(run_now)
    paths = build_output_paths(inputs.tdr_path, batch_code, run_now)
    paths.output_dir.mkdir(parents=True, exist_ok=True)
    report = ValidationReport()

    validate_inputs(inputs, report)
    if report.has_errors:
        write_validation_csv(report, paths.validation_errors_path)
        return RunOutputs(batch_code, paths.output_dir, paths.audit_workbook_path, None, paths.validation_errors_path, None)

    payroll_rows = []
    try:
        payroll_rows.extend(normalize_tdr_file(inputs.tdr_path, batch_code))
        if inputs.expense_path:
            lookup = load_employee_lookup(inputs.template_path)
            payroll_rows.extend(normalize_expense_file(inputs.expense_path, lookup, batch_code))
        else:
            report.add_warning(
                "Expense Transaction file intentionally skipped.",
                "Confirm this payroll import has no expense transactions.",
                source="Expense Transaction",
            )
    except Exception as exc:
        report.add_error(str(exc), "Correct the source file or template lookup data, then rerun.")

    if not payroll_rows:
        report.add_error(
            "Selected source file has no usable data rows.",
            "Confirm the Time Detail Report contains payroll rows and rerun.",
            source="Time Detail Report",
        )

    copy_template_to_audit(inputs.template_path, paths.audit_workbook_path)

    if report.has_errors:
        write_validation_csv(report, paths.validation_errors_path)
        return RunOutputs(batch_code, paths.output_dir, paths.audit_workbook_path, None, paths.validation_errors_path, None)

    write_tdr_rows(paths.audit_workbook_path, payroll_rows)
    export_pu_csv_with_excel(paths.audit_workbook_path, paths.payroll_csv_path, backend=excel_backend)
    validate_final_csv(paths.payroll_csv_path, report)

    validation_path = paths.validation_errors_path if report.has_errors else paths.validation_path
    write_validation_csv(report, validation_path)
    payroll_csv_path = None if report.has_errors else paths.payroll_csv_path
    return RunOutputs(batch_code, paths.output_dir, paths.audit_workbook_path, payroll_csv_path, validation_path, None)


def validate_inputs(inputs: RunInputs, report: ValidationReport) -> None:
    require_existing_file(inputs.template_path, "Payroll Upload Worksheet template", report)
    require_existing_file(inputs.tdr_path, "Time Detail Report", report)
    if inputs.expense_path:
        require_existing_file(inputs.expense_path, "Expense Transaction", report)
    elif not inputs.expense_skipped:
        report.add_error(
            "Expense Transaction decision is missing.",
            "Select the Expense Transaction file or confirm there is no Expense Transaction file for this payroll import.",
            source="Expense Transaction",
        )


def require_existing_file(path: Path, label: str, report: ValidationReport) -> None:
    if not path.exists() or not path.is_file():
        report.add_error(
            f"{label} file is missing.",
            f"Select a valid {label} file and rerun.",
            source=label,
        )
```

- [ ] **Step 4: Run test to verify it passes**

Run:

```bash
python3 -m pytest tests/test_runner.py -q
```

Expected: PASS.

- [ ] **Step 5: Run all non-Windows tests**

Run:

```bash
python3 -m pytest -q
```

Expected: PASS.

- [ ] **Step 6: Commit**

```bash
git add src/bex_payroll_import/runner.py tests/test_runner.py
git commit -m "feat: orchestrate payroll import run"
```

---

### Task 11: Tkinter desktop UI

**Files:**

- Create: `src/bex_payroll_import/ui.py`
- Modify: `src/bex_payroll_import/__main__.py`

- [ ] **Step 1: Add UI implementation**

Create `src/bex_payroll_import/ui.py`:

```python
from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from bex_payroll_import.config import AppConfig, load_config, save_config
from bex_payroll_import.models import RunInputs
from bex_payroll_import.runner import run_payroll_import


class PayrollImportApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BEX Payroll Import")
        self.geometry("720x360")
        self.resizable(False, False)

        self.config_state = load_config()
        self.template_path = tk.StringVar(value=str(self.config_state.template_path or ""))
        self.tdr_path = tk.StringVar(value="")
        self.expense_path = tk.StringVar(value="")
        self.expense_skipped = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="Select files to generate a payroll import.")

        self.build()
        self.refresh_generate_state()
        if not self.config_state.template_path or not self.config_state.template_path.exists():
            self.choose_template()

    def build(self) -> None:
        padding = {"padx": 12, "pady": 6}

        tk.Label(self, text="Payroll Upload Worksheet template").grid(row=0, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.template_path, width=70, state="readonly").grid(row=0, column=1, **padding)
        tk.Button(self, text="Change", command=self.choose_template).grid(row=0, column=2, **padding)

        tk.Label(self, text="Time Detail Report").grid(row=1, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.tdr_path, width=70, state="readonly").grid(row=1, column=1, **padding)
        tk.Button(self, text="Select", command=self.choose_tdr).grid(row=1, column=2, **padding)

        tk.Label(self, text="Expense Transaction").grid(row=2, column=0, sticky="w", **padding)
        tk.Entry(self, textvariable=self.expense_path, width=70, state="readonly").grid(row=2, column=1, **padding)
        tk.Button(self, text="Select", command=self.choose_expense).grid(row=2, column=2, **padding)

        tk.Checkbutton(
            self,
            text="I confirm there is no Expense Transaction file for this payroll import.",
            variable=self.expense_skipped,
            command=self.on_skip_expense_changed,
        ).grid(row=3, column=1, sticky="w", **padding)

        self.generate_button = tk.Button(self, text="Generate Payroll Import", command=self.generate)
        self.generate_button.grid(row=4, column=1, sticky="e", **padding)

        tk.Label(self, textvariable=self.status_text, wraplength=640, justify="left").grid(row=5, column=0, columnspan=3, sticky="w", **padding)

    def choose_template(self) -> None:
        path = filedialog.askopenfilename(title="Select Payroll Upload Worksheet", filetypes=[("Excel workbooks", "*.xlsx *.xlsm")])
        if path:
            template = Path(path)
            self.template_path.set(str(template))
            save_config(AppConfig(template_path=template))
            self.refresh_generate_state()

    def choose_tdr(self) -> None:
        path = filedialog.askopenfilename(title="Select Time Detail Report", filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.tdr_path.set(path)
            self.refresh_generate_state()

    def choose_expense(self) -> None:
        path = filedialog.askopenfilename(title="Select Expense Transaction", filetypes=[("Excel workbooks", "*.xlsx")])
        if path:
            self.expense_path.set(path)
            self.expense_skipped.set(False)
            self.refresh_generate_state()

    def on_skip_expense_changed(self) -> None:
        if self.expense_skipped.get():
            self.expense_path.set("")
        self.refresh_generate_state()

    def refresh_generate_state(self) -> None:
        ready = bool(self.template_path.get() and self.tdr_path.get() and (self.expense_path.get() or self.expense_skipped.get()))
        self.generate_button.config(state=tk.NORMAL if ready else tk.DISABLED)

    def generate(self) -> None:
        try:
            outputs = run_payroll_import(
                RunInputs(
                    template_path=Path(self.template_path.get()),
                    tdr_path=Path(self.tdr_path.get()),
                    expense_path=Path(self.expense_path.get()) if self.expense_path.get() else None,
                    expense_skipped=self.expense_skipped.get(),
                )
            )
            if outputs.payroll_csv_path:
                self.status_text.set(f"Payroll import created: {outputs.payroll_csv_path}")
                messagebox.showinfo("Payroll import created", f"Output folder:\n{outputs.output_dir}")
            else:
                self.status_text.set(f"Validation stopped the run: {outputs.validation_path}")
                messagebox.showerror("Validation stopped the run", f"Open this file for details:\n{outputs.validation_path}")
            open_folder(outputs.output_dir)
        except Exception as exc:
            messagebox.showerror("Payroll import failed", str(exc))
            self.status_text.set(f"Payroll import failed: {exc}")


def open_folder(path: Path) -> None:
    if os.name == "nt":
        os.startfile(path)


def main() -> None:
    app = PayrollImportApp()
    app.mainloop()
```

- [ ] **Step 2: Run import smoke test**

Run:

```bash
python3 -m bex_payroll_import
```

Expected on Mac: Tkinter window opens. Close it manually.

Expected on Windows: Tkinter window opens and file pickers work.

- [ ] **Step 3: Commit**

```bash
git add src/bex_payroll_import/ui.py src/bex_payroll_import/__main__.py
git commit -m "feat: add desktop app ui"
```

---

### Task 12: Documentation and packaging instructions

**Files:**

- Modify: `README.md`
- Modify: `SETUP_GUIDE.md`

- [ ] **Step 1: Replace README**

Replace `README.md` with:

```markdown
# BEX Payroll Import App

Small Windows desktop app for creating Vista payroll import CSV files from BEX payroll source reports.

## What it does

The app takes:

- Payroll Upload Worksheet template
- Time Detail Report
- Expense Transaction file, or a user confirmation that no expense file exists

It creates:

- payroll import CSV with no headers
- audit workbook
- validation report
- validation error file when a run is blocked

## Safety rules

The app copies the Payroll Upload Worksheet template for every run. It does not edit the saved template.

Real payroll workbooks and generated CSV files should stay out of git.

## Output folder

Outputs are created beside the selected Time Detail Report:

```text
<TDR folder>\Payroll Import Outputs\YYYY-MM-DD\
```

## Batch code

Every run gets a fresh batch code:

```text
CDMMDDHHMMSS
```

Example:

```text
CD0529143708
```

## Development

Install:

```bash
python3 -m pip install -e ".[dev]"
```

Run tests:

```bash
python3 -m pytest -q
```

Windows install with Excel automation support:

```powershell
py -m pip install -e ".[dev,windows]"
```
```

- [ ] **Step 2: Replace setup guide**

Replace `SETUP_GUIDE.md` with:

```markdown
# Setup guide

## Target machine

This app is built for one Windows user with:

- Windows
- local Microsoft Excel installed
- Python installed
- `openpyxl`
- `pywin32`

## Install dependencies

Open PowerShell in the project folder:

```powershell
py -m pip install -e ".[dev,windows]"
```

If `pywin32` needs to be installed separately:

```powershell
py -m pip install pywin32
```

## Start the app

```powershell
py -m bex_payroll_import
```

## First run

1. Select the Payroll Upload Worksheet template.
2. Select the Time Detail Report.
3. Select the Expense Transaction file, or check the box confirming there is no Expense Transaction file.
4. Click Generate Payroll Import.
5. Review the output folder.

## Success output

The app creates:

- `PayrollImport_<batch>_<timestamp>.csv`
- `PayrollUpload_Audit_<batch>_<timestamp>.xlsx`
- `ValidationReport_<batch>_<timestamp>.csv`

## Blocked run output

The app creates:

- `ValidationErrors_<batch>_<timestamp>.csv`
- audit workbook when the app reached workbook creation

Open the validation error file and correct the source file or template, then rerun.

## First live test

Use a known payroll period first.

Before importing into Vista, open the generated CSV and confirm:

- no header row
- one batch code on every row
- expected employee codes
- expense rows have `Hours=0`
- expense rows use `Department=DLPTO`
- expense rows use `Pay type=EXP REIM`
- expense rows use `Cost Type=L`
```

- [ ] **Step 3: Run test suite**

Run:

```bash
python3 -m pytest -q
```

Expected: PASS.

- [ ] **Step 4: Commit**

```bash
git add README.md SETUP_GUIDE.md
git commit -m "docs: add operator setup guide"
```

---

### Task 13: Windows packaging smoke path

**Files:**

- Modify: `SETUP_GUIDE.md`

- [ ] **Step 1: Build executable on Windows**

Run on Windows:

```powershell
py -m pip install -e ".[dev,windows]"
py -m PyInstaller --onefile --windowed --name BEXPayrollImport src\bex_payroll_import\__main__.py
```

Expected:

```text
dist\BEXPayrollImport.exe
```

- [ ] **Step 2: Launch packaged app**

Run on Windows:

```powershell
.\dist\BEXPayrollImport.exe
```

Expected: the desktop app opens.

- [ ] **Step 3: Add packaging notes to setup guide**

Append to `SETUP_GUIDE.md`:

```markdown

## Optional packaged app

Build the `.exe` on Windows:

```powershell
py -m PyInstaller --onefile --windowed --name BEXPayrollImport src\bex_payroll_import\__main__.py
```

Run:

```powershell
.\dist\BEXPayrollImport.exe
```

Give the user the `.exe` only after a successful test against a known payroll period.
```

- [ ] **Step 4: Commit**

```bash
git add SETUP_GUIDE.md
git commit -m "docs: add windows packaging notes"
```

---

## Manual acceptance test

Run this after implementation on the Windows machine with local Excel installed.

- [ ] Select the saved Payroll Upload Worksheet template.
- [ ] Select the Time Detail Report.
- [ ] Select the Expense Transaction file.
- [ ] Click Generate Payroll Import.
- [ ] Confirm the output folder exists beside the selected Time Detail Report.
- [ ] Confirm the audit workbook opens in Excel.
- [ ] Confirm the generated CSV has no headers.
- [ ] Confirm every row uses one `CDMMDDHHMMSS` batch code.
- [ ] Confirm expense rows use `Hours=0`.
- [ ] Confirm expense rows use `Department=DLPTO`.
- [ ] Confirm expense rows use `Pay type=EXP REIM`.
- [ ] Confirm expense rows use `Cost Type=L`.
- [ ] Confirm the saved template file timestamp did not change.

Failure-path checks:

- [ ] Generate stays disabled when the Expense Transaction file is blank and the confirmation box is unchecked.
- [ ] A skipped expense file creates a warning in the validation report.
- [ ] An unmatched expense employee creates a hard stop and no payroll import CSV.
- [ ] A missing source column creates a hard stop and names the missing column.

---

## Self-review notes

Spec coverage:

- Desktop app: Task 11.
- Saved template path: Task 3 and Task 11.
- Soft-required expense file: Task 10 and Task 11.
- Unique batch code: Task 3.
- Output folder beside TDR: Task 3.
- TDR and expense parsing: Tasks 4, 5, and 6.
- Template copy and formula preservation: Task 7.
- Excel recalculation/export: Task 8.
- Headerless CSV validation: Task 9.
- Validation hard stops and warnings: Tasks 2, 9, and 10.
- README and setup guide: Task 12.
- Windows packaging path: Task 13.

Known implementation risk:

- The Excel COM export must be tested on Windows. Mac tests can verify orchestration with a fake backend, but the real desktop Excel path needs the target operating system.
