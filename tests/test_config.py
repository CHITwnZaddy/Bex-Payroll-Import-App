from pathlib import Path

from bex_payroll_import.config import AppConfig, load_config, save_config


def test_save_and_load_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    template_path = tmp_path / "Payroll Upload Worksheet.xlsx"

    save_config(AppConfig(template_path=template_path), config_path)

    assert load_config(config_path) == AppConfig(template_path=template_path)


def test_load_missing_config_returns_empty_config(tmp_path: Path) -> None:
    assert load_config(tmp_path / "missing.json") == AppConfig(template_path=None)


def test_load_corrupt_config_returns_empty_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text("{not-json", encoding="utf-8")

    assert load_config(config_path) == AppConfig(template_path=None)


def test_load_unexpected_json_shape_returns_empty_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text("[]", encoding="utf-8")

    assert load_config(config_path) == AppConfig(template_path=None)


def test_load_int_template_path_returns_empty_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text('{"template_path": 123}', encoding="utf-8")

    assert load_config(config_path) == AppConfig(template_path=None)


def test_load_list_template_path_returns_empty_config(tmp_path: Path) -> None:
    config_path = tmp_path / "config.json"
    config_path.write_text('{"template_path": ["bad"]}', encoding="utf-8")

    assert load_config(config_path) == AppConfig(template_path=None)
