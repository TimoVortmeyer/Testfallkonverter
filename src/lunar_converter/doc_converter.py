"""Konvertierung von ``.doc`` nach ``.docx`` über Microsoft Word.

Ruft das mitgelieferte PowerShell-Skript ``resources/convert_doc_to_docx.ps1``
ohne Shell auf. Die Ausführungsrichtlinie von PowerShell wird dabei nicht umgangen.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from importlib.resources import as_file, files
from pathlib import Path

from .exceptions import ConfigurationError, DocConversionError, OutputDirectoryError
from .filesystem import prepare_output_directory
from .source_discovery import discover_source_files

POWERSHELL_EXECUTABLES: tuple[str, ...] = ("pwsh", "powershell")
WORD_SCRIPT_NAME = "convert_doc_to_docx.ps1"
CONVERSION_TIMEOUT_SECONDS = 300


def find_powershell() -> Path:
    for executable in POWERSHELL_EXECUTABLES:
        found = shutil.which(executable)
        if found:
            return Path(found)
    raise DocConversionError(
        "PowerShell wurde nicht gefunden; .doc-Dateien können nicht mit Word umgewandelt werden.",
        details="Weder 'pwsh' noch 'powershell' ist im PATH verfügbar.",
    )


def build_conversion_command(
    powershell: Path,
    script: Path,
    input_dir: Path,
    output_dir: Path,
    *,
    recursive: bool = False,
    show_progress: bool = False,
    skip_existing: bool = False,
) -> list[str]:
    command = [
        str(powershell),
        "-NoProfile",
        "-NonInteractive",
        "-File",
        str(script),
        "-InputDir",
        str(input_dir),
        "-OutputDir",
        str(output_dir),
    ]
    if recursive:
        command.append("-Recurse")
    if show_progress:
        command.append("-ShowProgress")
    if skip_existing:
        command.append("-SkipExisting")
    return command


def convert_doc_to_docx(source: Path, work_dir: Path) -> Path:
    """Wandelt ``source`` in ``work_dir`` mit Word um und liefert den Pfad der erzeugten ``.docx``."""
    powershell = find_powershell()
    input_dir = work_dir / "eingabe"
    output_dir = work_dir / "ausgabe"
    input_dir.mkdir(parents=True, exist_ok=True)
    # Fester ASCII-Name vermeidet Probleme mit Sonderzeichen im Dateinamen.
    local_source = input_dir / "quelle.doc"
    shutil.copy2(source, local_source)
    expected_output = output_dir / "quelle.docx"

    with as_file(files("lunar_converter").joinpath("resources").joinpath(WORD_SCRIPT_NAME)) as script:
        command = build_conversion_command(powershell, script, input_dir, output_dir)
        try:
            completed = subprocess.run(
                command,
                check=False,
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=CONVERSION_TIMEOUT_SECONDS,
                env={**os.environ, "LUNAR_UTF8_OUTPUT": "1"},
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise DocConversionError(
                f"Word-Konvertierung von '{source.name}' konnte nicht ausgeführt werden.",
                details=f"Befehl: {command}\n{type(exc).__name__}: {exc}",
            ) from exc

    details = (
        f"Befehl: {command}\nExit-Code: {completed.returncode}\n"
        f"stdout: {completed.stdout.strip()}\nstderr: {completed.stderr.strip()}"
    )
    if completed.returncode != 0:
        raise DocConversionError(f"Word-Konvertierung von '{source.name}' ist fehlgeschlagen.", details=details)
    if not expected_output.is_file() or expected_output.stat().st_size == 0:
        raise DocConversionError(f"Word hat für '{source.name}' keine DOCX-Datei erzeugt.", details=details)
    return expected_output


def prepare_docx_directory(input_dir: Path, output_dir: Path) -> tuple[int, int]:
    """Bereitet rekursiv alle Word-Dateien in einer gespiegelten DOCX-Struktur vor.

    Rückgabe: (Anzahl Quelldateien, PowerShell-Exitcode).
    """
    if not input_dir.is_dir():
        raise ConfigurationError(f"Eingabeordner '{input_dir}' existiert nicht oder ist kein Ordner.")
    source_root = input_dir.resolve()
    target_root = output_dir.resolve()
    if (
        source_root == target_root
        or source_root.is_relative_to(target_root)
        or target_root.is_relative_to(source_root)
    ):
        raise ConfigurationError("Eingabe- und Ausgabeordner dürfen nicht identisch sein oder ineinander liegen.")

    source_files = discover_source_files(input_dir)
    if not source_files:
        return 0, 0
    powershell = find_powershell()
    skip_existing = False
    if output_dir.exists() and output_dir.is_dir() and any(output_dir.iterdir()):
        answer = _ask_existing_output_action(output_dir)
        if answer == "delete":
            prepare_output_directory(output_dir, protected_dir=input_dir, existing_policy="clear")
        else:
            prepare_output_directory(output_dir, protected_dir=input_dir, existing_policy="preserve")
            skip_existing = True
    else:
        prepare_output_directory(output_dir, protected_dir=input_dir)

    with as_file(files("lunar_converter").joinpath("resources").joinpath(WORD_SCRIPT_NAME)) as script:
        command = build_conversion_command(
            powershell,
            script,
            input_dir,
            output_dir,
            recursive=True,
            show_progress=True,
            skip_existing=skip_existing,
        )
        try:
            completed = subprocess.run(
                command,
                check=False,
                capture_output=False,
                text=True,
                timeout=CONVERSION_TIMEOUT_SECONDS * len(source_files),
                env={**os.environ, "LUNAR_UTF8_OUTPUT": "1"},
            )
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise DocConversionError(
                "Die Stapelvorbereitung der Word-Dateien konnte nicht ausgeführt werden.",
                details=f"Befehl: {command}\n{type(exc).__name__}: {exc}",
            ) from exc
    return len(source_files), completed.returncode


def _ask_existing_output_action(output_dir: Path) -> str:
    try:
        answer = input(
            f"Output-Ordner '{output_dir}' ist nicht leer. "
            "[L] Löschen und neu starten / [N] Nicht löschen, vorhandene DOCX überspringen / [A] Abbrechen: "
        )
    except EOFError:
        answer = ""
    normalized = answer.strip().casefold()
    if normalized in {"l", "löschen", "loeschen", "j", "ja"}:
        return "delete"
    if normalized in {"n", "nein", "nicht löschen", "nicht loeschen", "beibehalten"}:
        return "preserve"
    raise OutputDirectoryError(
        f"DOCX-Vorbereitung abgebrochen; Output-Ordner '{output_dir}' wurde nicht verändert."
    )
