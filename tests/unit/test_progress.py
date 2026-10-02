from __future__ import annotations

from io import StringIO

from lunar_converter.progress import TerminalProgress


def test_progress_zeigt_gesamt_und_restzeit() -> None:
    now = [0.0]
    output = StringIO()
    progress = TerminalProgress("Preflight", 5, stream=output, clock=lambda: now[0])

    progress.update(0, current="fall.docx", status="Prüfe")
    now[0] = 4.0
    progress.update(2, current="fall.docx", status="matched")

    rendered = output.getvalue()
    assert "2/5 (40%)" in rendered
    assert "Laufzeit 00:04" in rendered
    assert "Gesamt ~00:10" in rendered
    assert "Rest ~00:06" in rendered


def test_progress_meldet_leeren_eingabeordner() -> None:
    output = StringIO()
    progress = TerminalProgress("Konvertierung", 0, stream=output)

    progress.finish()

    assert output.getvalue() == "Konvertierung: keine Dateien gefunden.\n"