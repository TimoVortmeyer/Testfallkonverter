"""Profil-Preflight für Word-Dateien mit CSV-Bericht."""

from __future__ import annotations

import csv
import json
import tempfile
from pathlib import Path
from typing import Any

from .doc_converter import convert_doc_to_docx
from .docx_reader import read_docx
from .exceptions import ConfigurationError, LunarConverterError
from .models import ProfileDefinition
from .profile_detector import detect_profile
from .profile_loader import load_profiles
from .source_discovery import discover_source_files, relative_display_path

CSV_FIELDS = (
    "datei",
    "status",
    "erkanntes_profil",
    "passende_profile",
    "profilpruefungen",
    "fehlercode",
    "fehler",
)


def run_preflight(input_dir: Path, csv_path: Path, config_dir: Path) -> tuple[int, int]:
    """Prüft alle Word-Dateien gegen alle Profile und schreibt eine CSV.

    Rückgabe: (Anzahl Dateien, Anzahl Dateien ohne genau einen Profiltreffer).
    """
    if not input_dir.is_dir():
        raise ConfigurationError(f"Eingabeordner '{input_dir}' existiert nicht oder ist kein Ordner.")
    profiles = load_profiles(config_dir)
    files = discover_source_files(input_dir)
    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("w", encoding="utf-8-sig", newline="") as output:
            writer = csv.DictWriter(output, fieldnames=CSV_FIELDS, delimiter=";")
            writer.writeheader()
            failed = 0
            for source in files:
                row = _check_file(source, input_dir, profiles)
                writer.writerow(row)
                if row["status"] != "matched":
                    failed += 1
    except OSError as exc:
        raise ConfigurationError(f"CSV-Bericht '{csv_path}' konnte nicht geschrieben werden.", details=str(exc)) from exc
    return len(files), failed


def _check_file(source: Path, input_dir: Path, profiles: list[ProfileDefinition]) -> dict[str, Any]:
    row: dict[str, Any] = {
        "datei": relative_display_path(source, input_dir),
        "status": "error",
        "erkanntes_profil": "",
        "passende_profile": "",
        "profilpruefungen": "[]",
        "fehlercode": "",
        "fehler": "",
    }
    try:
        with tempfile.TemporaryDirectory(prefix="lunar_preflight_") as temp_dir:
            docx_path = source
            if source.suffix.casefold() == ".doc":
                docx_path = convert_doc_to_docx(source, Path(temp_dir) / "doc-konvertierung")
            document = read_docx(docx_path)
        detection = detect_profile(document, profiles)
        row["status"] = detection.status
        row["erkanntes_profil"] = detection.profile.id if detection.profile else ""
        row["passende_profile"] = ";".join(detection.matching_profile_ids)
        row["profilpruefungen"] = json.dumps([check.to_dict() for check in detection.checks], ensure_ascii=False)
    except LunarConverterError as exc:
        row["fehlercode"] = exc.code
        row["fehler"] = exc.message
        if exc.details:
            row["fehler"] += f" Details: {exc.details}"
    except Exception as exc:
        row["fehlercode"] = "unexpected_error"
        row["fehler"] = f"{type(exc).__name__}: {exc}"
    return row