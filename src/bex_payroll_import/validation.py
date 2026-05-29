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
