"""Verantwortliche aus Deckblaettern als CSV exportieren."""

from __future__ import annotations

import csv
import logging
import re
import tempfile
from pathlib import Path
from typing import TypeAlias

from .doc_converter import convert_doc_to_docx
from .docx_reader import read_docx
from .exceptions import ConfigurationError, LunarConverterError
from .logging_setup import attach_log_file, context_logger
from .models import SourceDocument, SourceParagraph, SourceTable
from .source_discovery import discover_source_files, relative_display_path

_LABEL = re.compile(r"^Verantwortliche(?:r|s)?(?:\s+Team)?\s*[:\t]\s*(.+)$", re.IGNORECASE)
_TESTCASE = re.compile(r"^Testfall\s*[:\t]", re.IGNORECASE)
_PERSON = re.compile(r"^[^\W\d_]+(?:[-'][^\W\d_]+)?(?:\s+[^\W\d_]+(?:[-'][^\W\d_]+)?){1,3}$")
_CONTACT = re.compile(r"^(?:T|Tel|Telefon|F|Fax)\s*[:.]|\b[^\s@]+@[^\s@]+\.[^\s@]+\b", re.IGNORECASE)
ResponsibleEmailMapping: TypeAlias = dict[tuple[str, str], set[str]]


def responsible_mapping_key(relative_file: str, responsible: str) -> tuple[str, str]:
    normalized_path = "/".join(
        part for part in relative_file.strip().replace("\\", "/").split("/") if part
    ).casefold()
    normalized_name = " ".join(responsible.split()).casefold()
    return normalized_path, normalized_name


def load_responsible_email_mapping(csv_path: Path | None) -> ResponsibleEmailMapping:
    """Lädt optionale Name-/Datei-/E-Mail-Zuordnungen aus der angereicherten Verantwortlichen-CSV."""
    if csv_path is None:
        return {}
    try:
        with csv_path.open("r", encoding="utf-8-sig", newline="") as source:
            reader = csv.DictReader(source, delimiter=";")
            headers = {
                header.strip().casefold(): header
                for header in (reader.fieldnames or [])
                if header and header.strip()
            }
            required_headers = {"verantwortlicher", "datei", "email"}
            missing_headers = sorted(required_headers - headers.keys())
            if missing_headers:
                raise ConfigurationError(
                    f"CSV '{csv_path}' benötigt die Spalten: "
                    + ", ".join(sorted(required_headers))
                    + f". Fehlend: {', '.join(missing_headers)}."
                )
            mapping: ResponsibleEmailMapping = {}
            for line_number, row in enumerate(reader, start=2):
                if row is None:
                    continue
                name = (row.get(headers["verantwortlicher"]) or "").strip()
                relative_file = (row.get(headers["datei"]) or "").strip()
                email = (row.get(headers["email"]) or "").strip()
                if not name or not relative_file:
                    raise ConfigurationError(
                        f"CSV '{csv_path}', Zeile {line_number}: "
                        "'verantwortlicher' und 'datei' dürfen nicht leer sein."
                    )
                if email:
                    mapping.setdefault(responsible_mapping_key(relative_file, name), set()).add(email)
            return mapping
    except ConfigurationError:
        raise
    except (OSError, UnicodeDecodeError, csv.Error) as exc:
        raise ConfigurationError(
            f"Mapping-CSV '{csv_path}' ist nicht lesbar oder ungültig.", details=str(exc)
        ) from exc


def _contact_names(lines: list[str]) -> list[str]:
    return [
        line for index, line in enumerate(lines)
        if _PERSON.fullmatch(line) and index + 1 < len(lines) and _CONTACT.search(lines[index + 1])
    ]


def cover_responsibles(document: SourceDocument) -> list[str]:
    """Liest beschriftete Verantwortliche und Namen mit Kontaktdaten vom Deckblatt."""
    found: list[str] = []
    saw_testcase = False
    contact_lines: list[str] = []
    for block in document.blocks:
        last_cover_table = isinstance(block, SourceTable) and saw_testcase
        if isinstance(block, SourceTable):
            paragraphs = [paragraph for row in block.rows for cell in row.cells for paragraph in cell.paragraphs]
        else:
            paragraphs = [block]
        for paragraph in paragraphs:
            for line in paragraph.text.splitlines():
                text = line.strip()
                match = _LABEL.match(text)
                if match:
                    value = match.group(1).strip()
                    if value and value not in found:
                        found.append(value)
                if _TESTCASE.match(text):
                    saw_testcase = True
                elif saw_testcase and text:
                    contact_lines.append(text)
        if last_cover_table:
            break
    for name in _contact_names(contact_lines):
        if name not in found:
            found.append(name)
    return found


def export_responsibles(input_dir: Path, csv_path: Path, logger: logging.Logger) -> tuple[int, int]:
    """Schreibt je Dokument und Verantwortlichem eine CSV-Zeile; meldet unvollstaendige Dateien."""
    if not input_dir.is_dir():
        raise ConfigurationError(f"Eingabeordner '{input_dir}' existiert nicht oder ist kein Ordner.")
    files = discover_source_files(input_dir)
    log_path = csv_path.with_name(f"{csv_path.stem}.verantwortliche.log")
    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        if csv_path.exists() or log_path.exists():
            raise ConfigurationError(f"CSV oder Logdatei existiert bereits: '{csv_path}', '{log_path}'.")
        attach_log_file(logger, csv_path.parent, log_path.name)
    except OSError as exc:
        raise ConfigurationError(f"Logdatei '{log_path}' konnte nicht angelegt werden.", details=str(exc)) from exc
    logger.info("Start der Verantwortlichen-Erfassung: %d Datei(en) in '%s'.", len(files), input_dir)
    if not files:
        logger.warning("Im Eingabeordner und seinen Unterordnern wurden keine .docx- oder .doc-Dateien gefunden.")
    rows: list[tuple[str, str]] = []
    failures = 0
    for source in files:
        relative_path = relative_display_path(source, input_dir)
        log = context_logger(logger, relative_path)
        try:
            with tempfile.TemporaryDirectory(prefix="lunar_verantwortliche_") as temp_dir:
                docx_path = source
                if source.suffix.casefold() == ".doc":
                    docx_path = convert_doc_to_docx(source, Path(temp_dir) / "doc-konvertierung")
                names = cover_responsibles(read_docx(docx_path))
            if not names:
                raise ValueError("Kein Verantwortlicher auf dem Deckblatt gefunden")
            rows.extend((name, relative_path) for name in names)
            log.info("Verantwortliche gefunden: %s.", "; ".join(names))
        except (LunarConverterError, OSError, ValueError) as exc:
            failures += 1
            log.error("Verantwortliche konnten nicht erfasst werden: %s", exc)
    try:
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        with csv_path.open("x", encoding="utf-8-sig", newline="") as output:
            writer = csv.writer(output, delimiter=";")
            writer.writerow(("verantwortlicher", "datei"))
            writer.writerows(rows)
    except OSError as exc:
        raise ConfigurationError(f"CSV '{csv_path}' konnte nicht angelegt werden (existiert sie bereits?).", details=str(exc)) from exc
    logger.info("Ende der Verantwortlichen-Erfassung: %d gesamt, %d ohne Ergebnis. CSV: %s", len(files), failures, csv_path)
    return len(files), failures