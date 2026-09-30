"""Dateisystemfunktionen: Output-Preflight, Ordnernamen und atomarer Export."""

from __future__ import annotations

import os
import re
import shutil
import uuid
from pathlib import Path

from .exceptions import ExportError, OutputDirectoryError

_INVALID_CHARS_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WHITESPACE_RE = re.compile(r"\s+")
_WINDOWS_RESERVED = frozenset(
    {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)), *(f"LPT{i}" for i in range(1, 10))}
)


def sanitize_folder_name(stem: str) -> str:
    """Bildet einen gültigen Ordnernamen aus dem Dateinamen ohne Erweiterung.

    Wie im POC werden ungültige Zeichen und Leerraum durch ``_`` ersetzt.
    """
    safe = _INVALID_CHARS_RE.sub("_", stem.strip())
    safe = _WHITESPACE_RE.sub("_", safe)
    safe = safe.strip("._ ")
    if not safe:
        return "dokument"
    if safe.split(".")[0].upper() in _WINDOWS_RESERVED:
        safe = f"{safe}_"
    return safe


def prepare_output_directory(output_dir: Path) -> None:
    """Preflight: Der Output-Basisordner wird angelegt oder muss leer sein."""
    if output_dir.exists():
        if not output_dir.is_dir():
            raise OutputDirectoryError(f"Output-Pfad '{output_dir}' existiert, ist aber kein Ordner.")
        entries = sorted(entry.name for entry in output_dir.iterdir())
        if entries:
            shown = ", ".join(entries[:5]) + (" …" if len(entries) > 5 else "")
            raise OutputDirectoryError(
                f"Output-Ordner '{output_dir}' ist nicht leer ({len(entries)} Einträge: {shown}). "
                "Der Lauf wird abgebrochen, es wird nichts überschrieben. Bitte einen leeren oder neuen Ordner angeben."
            )
        return
    try:
        output_dir.mkdir(parents=True)
    except OSError as exc:
        raise OutputDirectoryError(f"Output-Ordner '{output_dir}' konnte nicht angelegt werden.", details=str(exc)) from exc


def publish_directory(staged_dir: Path, final_dir: Path) -> None:
    """Überträgt einen vollständig vorbereiteten Testfallordner atomar in den Output-Ordner.

    Der Inhalt wird zunächst in einen versteckten Staging-Ordner neben dem Ziel
    kopiert (gleiches Laufwerk) und dann per ``os.rename`` in einem Schritt
    sichtbar gemacht. Bei Fehlern bleibt kein unvollständiger Zielordner zurück.
    """
    if final_dir.exists():
        raise ExportError(f"Zielordner '{final_dir}' existiert bereits; es wird nichts überschrieben.", code="output_name_conflict")
    staging = final_dir.parent / f".{final_dir.name}.tmp-{uuid.uuid4().hex[:8]}"
    try:
        shutil.copytree(staged_dir, staging)
        os.rename(staging, final_dir)
    except OSError as exc:
        raise ExportError(f"Testfallordner '{final_dir.name}' konnte nicht geschrieben werden.", details=str(exc)) from exc
    finally:
        if staging.exists():
            shutil.rmtree(staging, ignore_errors=True)
