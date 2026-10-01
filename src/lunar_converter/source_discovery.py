"""Ermittlung der zu verarbeitenden Quelldateien."""

from __future__ import annotations

import os
from pathlib import Path

SUPPORTED_SUFFIXES: frozenset[str] = frozenset({".docx", ".doc"})
TEMPORARY_PREFIX = "~$"


def is_temporary_word_file(path: Path) -> bool:
    return path.name.startswith(TEMPORARY_PREFIX)


def discover_source_files(input_dir: Path, exclude_dir: Path | None = None) -> list[Path]:
    """Liefert alle ``.docx``/``.doc``-Dateien im Eingabeordner und allen Unterordnern.

    Sortierung nach relativem Pfad (ohne Groß-/Kleinschreibung). Temporäre Word-Dateien
    (``~$…``), symbolische Verzeichnisverknüpfungen und ``exclude_dir`` (z. B. ein im
    Eingabeordner liegender Output-Ordner) werden ignoriert.
    """
    excluded = exclude_dir.resolve() if exclude_dir is not None else None
    files: list[Path] = []
    for root, dirs, names in os.walk(input_dir, followlinks=False):
        root_path = Path(root)
        dirs[:] = [name for name in dirs if excluded is None or (root_path / name).resolve() != excluded]
        for name in names:
            path = root_path / name
            if path.suffix.lower() in SUPPORTED_SUFFIXES and not is_temporary_word_file(path) and path.is_file():
                files.append(path)
    return sorted(files, key=lambda path: _sort_key(path, input_dir))


def relative_display_path(path: Path, input_dir: Path) -> str:
    """Pfad relativ zum Eingabeordner für Logs und Meldungen, z. B. ``Unterordner/Datei.docx``."""
    try:
        return path.relative_to(input_dir).as_posix()
    except ValueError:
        return path.name


def _sort_key(path: Path, input_dir: Path) -> tuple[tuple[str, ...], str]:
    relative = path.relative_to(input_dir)
    return tuple(part.casefold() for part in relative.parts), relative.as_posix()
