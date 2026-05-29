from __future__ import annotations

from datetime import datetime


def build_batch_code(now: datetime) -> str:
    return f"CD{now:%m%d%H%M%S}"
