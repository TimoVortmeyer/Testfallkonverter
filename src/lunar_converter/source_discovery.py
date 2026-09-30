"""Ermittlung der zu verarbeitenden Quelldateien."""

from __future__ import annotations

from pathlib import Path

SUPPORTED_SUFFIXES: frozenset[str] = frozenset({".docx", ".doc"})
TEMPORARY_PREFIX = "~$"


def is_temporary_word_file(path: Path) -> bool:
    return path.name.startswith(TEMPORARY_PREFIX)


def discover_source_files(input_dir: Path) -> list[Path]:
    """Liefert alle ``.docx``/``.doc``-Dateien direkt im Eingabeordner, alphabetisch nach Dateiname.

    Temporäre Word-Dateien (``~$…``) und Unterordner werden ignoriert.
    """
    files = [
        path
        for path in input_dir.iterdir()
        if path.is_file() and path.suffix.lower() in SUPPORTED_SUFFIXES and not is_temporary_word_file(path)
    ]
    return sorted(files, key=lambda path: (path.name.casefold(), path.name))
