"""Gemeinsame Hilfsfunktionen und Pfade für Tests."""

from __future__ import annotations

import json
from collections.abc import Callable
from pathlib import Path
from typing import Any

RunConvert = Callable[..., int]

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = PROJECT_ROOT / "config"
SCHEMA_PATH = PROJECT_ROOT / "schema" / "testcase.schema.json"
REAL_SAMPLE = PROJECT_ROOT / "input" / "sample.docx"


def load_report(output_dir: Path) -> dict[str, Any]:
    return json.loads((output_dir / "conversion-report.json").read_text(encoding="utf-8"))


def load_testcase(output_dir: Path, folder: str) -> dict[str, Any]:
    return json.loads((output_dir / folder / "testcase.json").read_text(encoding="utf-8"))


def file_entry(report: dict[str, Any], file_name: str) -> dict[str, Any]:
    for entry in report["files"]:
        if Path(entry["input_file"]).name == file_name:
            return entry
    raise AssertionError(f"{file_name} fehlt im Report")


def error_codes(entry: dict[str, Any]) -> list[str]:
    return [error["code"] for error in entry["errors"]]


def warning_codes(entry: dict[str, Any]) -> list[str]:
    return [warning["code"] for warning in entry["warnings"]]


def final_entries(output_dir: Path) -> list[str]:
    """Alle Einträge im Output-Ordner außer Log und Report."""
    return sorted(p.name for p in output_dir.iterdir() if p.name not in {"conversion.log", "conversion-report.json"})
