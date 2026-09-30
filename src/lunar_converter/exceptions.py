"""Fehlerklassen der Konvertierung.

Jeder fachliche Fehler trägt einen stabilen Fehlercode, eine verständliche
Ursache, optional einen Feldpfad sowie technische Details. Daraus wird der
Eintrag im Conversion-Report erzeugt.
"""

from __future__ import annotations

from .models import Issue


class LunarConverterError(Exception):
    """Basisfehler der Konvertierung."""

    code: str = "conversion_error"

    def __init__(self, message: str, *, field: str | None = None, details: str | None = None, code: str | None = None) -> None:
        super().__init__(message)
        self.message = message
        self.field = field
        self.details = details
        if code is not None:
            self.code = code

    def to_issues(self) -> list[Issue]:
        return [Issue(code=self.code, message=self.message, field=self.field, details=self.details)]


class ConfigurationError(LunarConverterError):
    """Globale Konfiguration (Profile, Schema, Parameter) ist ungültig. Führt zum Abbruch des gesamten Laufs."""

    code = "configuration_error"


class OutputDirectoryError(LunarConverterError):
    """Der Output-Basisordner ist nicht verwendbar (z. B. nicht leer). Führt zum Abbruch des gesamten Laufs."""

    code = "output_dir_not_empty"


class InputFileError(LunarConverterError):
    """Die Quelldatei kann nicht gelesen werden."""

    code = "input_file_error"


class DocConversionError(LunarConverterError):
    """Die DOC-zu-DOCX-Konvertierung über Microsoft Word ist fehlgeschlagen."""

    code = "doc_conversion_failed"


class ProfileDetectionError(LunarConverterError):
    """Kein oder mehr als ein Profil passt zum Dokument."""


class ExportError(LunarConverterError):
    """Der finale Export des Testfallordners ist fehlgeschlagen."""

    code = "export_failed"


class ValidationFailedError(LunarConverterError):
    """Eine oder mehrere Validierungen sind fehlgeschlagen."""

    code = "validation_failed"

    def __init__(self, issues: list[Issue]) -> None:
        if not issues:
            raise ValueError("ValidationFailedError benötigt mindestens einen Befund.")
        first = issues[0]
        super().__init__(first.message, field=first.field, details=first.details, code=first.code)
        self.issues = list(issues)

    def to_issues(self) -> list[Issue]:
        return list(self.issues)
