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

from .exceptions import DocConversionError

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


def build_conversion_command(powershell: Path, script: Path, input_dir: Path, output_dir: Path) -> list[str]:
    return [
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
