from __future__ import annotations


class PayrollImportError(Exception):
    """Base error for expected payroll import failures."""


class ValidationFailedError(PayrollImportError):
    """Raised when validation errors block CSV generation."""


class ExcelAutomationError(PayrollImportError):
    """Raised when desktop Excel cannot recalculate or export."""
