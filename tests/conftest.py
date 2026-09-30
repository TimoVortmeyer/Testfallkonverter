from __future__ import annotations

import json
import logging
import shutil
from collections.abc import Iterator
from pathlib import Path

import pytest

from lunar_converter.cli import main
from lunar_converter.logging_setup import LOGGER_NAME, shutdown_logging
from tests.helpers import CONFIG_DIR, SCHEMA_PATH, RunConvert


@pytest.fixture(autouse=True)
def _close_log_handlers() -> Iterator[None]:
    yield
    shutdown_logging(logging.getLogger(LOGGER_NAME))


@pytest.fixture
def input_dir(tmp_path: Path) -> Path:
    path = tmp_path / "input"
    path.mkdir()
    return path


@pytest.fixture
def output_dir(tmp_path: Path) -> Path:
    return tmp_path / "output"


@pytest.fixture
def run_convert(input_dir: Path, output_dir: Path) -> RunConvert:
    def run(*extra: str, config_dir: Path | None = None) -> int:
        args = [
            "convert",
            "--input-dir",
            str(input_dir),
            "--output-dir",
            str(output_dir),
            "--schema-path",
            str(SCHEMA_PATH),
            "--config-dir",
            str(config_dir or CONFIG_DIR),
            *extra,
        ]
        return main(args)

    return run


@pytest.fixture
def two_profile_config(tmp_path: Path) -> Path:
    """Konfiguration mit zwei inhaltlich gleichen Profilen -> beide passen (mehrdeutig)."""
    config = tmp_path / "config_zwei_profile"
    profiles = config / "profiles"
    profiles.mkdir(parents=True)
    shutil.copy(CONFIG_DIR / "profiles" / "lunar_standard_v1.json", profiles / "lunar_standard_v1.json")
    data = json.loads((CONFIG_DIR / "profiles" / "lunar_standard_v1.json").read_text(encoding="utf-8"))
    data["id"] = "lunar_kopie_v1"
    data["name"] = "Kopie des Standardprofils"
    (profiles / "lunar_kopie_v1.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return config
