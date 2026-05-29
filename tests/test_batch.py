from datetime import datetime

from bex_payroll_import.batch import build_batch_code


def test_build_batch_code_uses_cd_mm_dd_hh_mm_ss() -> None:
    now = datetime(2026, 5, 29, 14, 37, 8)

    assert build_batch_code(now) == "CD0529143708"
