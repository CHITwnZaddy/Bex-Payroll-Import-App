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
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return AppConfig(template_path=None)
    if not isinstance(data, dict):
        return AppConfig(template_path=None)
    raw_template_path = data.get("template_path")
    if raw_template_path is None:
        return AppConfig(template_path=None)
    if not isinstance(raw_template_path, str):
        return AppConfig(template_path=None)
    return AppConfig(template_path=Path(raw_template_path) if raw_template_path else None)


def save_config(config: AppConfig, config_path: Path = APP_CONFIG_PATH) -> None:
    config_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "template_path": str(config.template_path) if config.template_path else None,
    }
    config_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
