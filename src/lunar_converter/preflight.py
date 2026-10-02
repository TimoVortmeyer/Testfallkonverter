"""Profil-Preflight für Word-Dateien mit CSV-Bericht."""

from __future__ import annotations

import csv
import json
import logging
import tempfile
from pathlib import Path
from typing import Any

from .doc_converter import convert_doc_to_docx
from .docx_reader import read_docx
from .exceptions import ConfigurationError, LunarConverterError
from .logging_setup import attach_log_file, context_logger
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


def run_preflight(input_dir: Path, csv_path: Path, config_dir: Path, logger: logging.Logger) -> tuple[int, int]:
    """Prüft alle Word-Dateien gegen alle Profile und schreibt eine CSV.

    Rückgabe: (Anzahl Dateien, Anzahl Dateien ohne genau einen Profiltreffer).
    """
    if not input_dir.is_dir():
        raise ConfigurationError(f"Eingabeordner '{input_dir}' existiert nicht oder ist kein Ordner.")
    profiles = load_profiles(config_dir)
    files = discover_source_files(input_dir)
    log_path = csv_path.with_name(f"{csv_path.stem}.preflight.log")
    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        attach_log_file(logger, csv_path.parent, log_path.name)
        logger.info(
            "Start des Profil-Preflights: %d Datei(en) in '%s', Profile: %s.",
            len(files),
            input_dir,
            ", ".join(profile.id for profile in profiles),
        )
        if not files:
            logger.warning("Im Eingabeordner und seinen Unterordnern wurden keine .docx- oder .doc-Dateien gefunden.")
        with csv_path.open("w", encoding="utf-8-sig", newline="") as output:
            writer = csv.DictWriter(output, fieldnames=CSV_FIELDS, delimiter=";")
            writer.writeheader()
            failed = 0
            for source in files:
                row = _check_file(source, input_dir, profiles)
                writer.writerow(row)
                if row["status"] != "matched":
                    failed += 1
                _log_result(logger, row)
            logger.info(
                "Ende des Profil-Preflights: %d gesamt, %d eindeutig erkannt, %d ohne eindeutigen Treffer. CSV: %s",
                len(files),
                len(files) - failed,
                failed,
                csv_path,
            )
    except OSError as exc:
        raise ConfigurationError(f"CSV-Bericht '{csv_path}' konnte nicht geschrieben werden.", details=str(exc)) from exc
    return len(files), failed


def _log_result(logger: logging.Logger, row: dict[str, Any]) -> None:
    log = context_logger(logger, str(row["datei"]), str(row["erkanntes_profil"]))
    status = str(row["status"])
    if status == "matched":
        log.info("Profil erkannt.")
    elif status == "ambiguous_profile":
        log.warning("Mehrere Profile passen: %s.", row["passende_profile"])
    elif status == "no_matching_profile":
        log.warning("Kein Profil passt.")
    else:
        log.error("Datei konnte nicht geprüft werden (%s): %s", row["fehlercode"], row["fehler"])
    if row["profilpruefungen"] != "[]":
        checks = json.loads(str(row["profilpruefungen"]))
        for check in checks:
            check_log = context_logger(logger, str(row["datei"]), str(check["profile"]))
            check_log.debug("Profilprüfung: %s – %s", "passt" if check["matched"] else "passt nicht", " ".join(check["reasons"]))


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