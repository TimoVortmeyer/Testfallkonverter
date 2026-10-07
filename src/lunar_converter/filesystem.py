"""Dateisystemfunktionen: Output-Preflight, Ordnernamen und atomarer Export."""

from __future__ import annotations

import os
import re
import shutil
import stat
import sys
import uuid
from pathlib import Path
from typing import Literal

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


def prepare_output_directory(
    output_dir: Path,
    protected_dir: Path | None = None,
    *,
    existing_policy: Literal["prompt", "clear", "preserve"] = "prompt",
) -> None:
    """Output vorbereiten; bestehende Inhalte nur nach ausdrücklicher Bestätigung löschen."""
    if protected_dir is not None:
        resolved_output = output_dir.resolve()
        resolved_protected = protected_dir.resolve()
        if resolved_output == resolved_protected or resolved_protected.is_relative_to(resolved_output):
            raise OutputDirectoryError(
                f"Output-Ordner '{output_dir}' ist der Eingabeordner oder ein übergeordneter Ordner. "
                "Dieser Pfad darf nicht zum Löschen freigegeben werden."
            )
    if output_dir.exists():
        if not output_dir.is_dir():
            raise OutputDirectoryError(f"Output-Pfad '{output_dir}' existiert, ist aber kein Ordner.")
        entries = sorted(entry.name for entry in output_dir.iterdir())
        if entries:
            if existing_policy == "preserve":
                return
            if existing_policy == "clear":
                try:
                    shutil.rmtree(output_dir, onerror=_remove_readonly)
                    output_dir.mkdir(parents=True)
                except OSError as exc:
                    raise OutputDirectoryError(
                        f"Output-Ordner '{output_dir}' konnte nicht vollständig gelöscht und neu angelegt werden.",
                        details=str(exc),
                    ) from exc
                return
            shown = ", ".join(entries[:5]) + (" …" if len(entries) > 5 else "")
            print(
                f"WARNUNG: Output-Ordner '{output_dir}' ist nicht leer "
                f"({len(entries)} Einträge: {shown}). Bei Bestätigung wird der gesamte Ordner gelöscht.",
                file=sys.stderr,
            )
            try:
                answer = input("Soll der Output-Ordner einschließlich aller Inhalte gelöscht werden? [ja/N]: ")
            except EOFError:
                answer = ""
            if answer.strip().casefold() not in {"j", "ja"}:
                raise OutputDirectoryError(f"Lauf abgebrochen; Output-Ordner '{output_dir}' wurde nicht gelöscht.")
            try:
                shutil.rmtree(output_dir, onerror=_remove_readonly)
                output_dir.mkdir(parents=True)
            except OSError as exc:
                raise OutputDirectoryError(
                    f"Output-Ordner '{output_dir}' konnte nicht vollständig gelöscht und neu angelegt werden.",
                    details=str(exc),
                ) from exc
        return
    try:
        output_dir.mkdir(parents=True)
    except OSError as exc:
        raise OutputDirectoryError(f"Output-Ordner '{output_dir}' konnte nicht angelegt werden.", details=str(exc)) from exc


def _remove_readonly(function, path, exc_info) -> None:
    error = exc_info[1]
    if not isinstance(error, PermissionError):
        raise error
    os.chmod(path, stat.S_IWRITE)
    function(path)


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
