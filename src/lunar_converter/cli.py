"""Kommandozeile: ``python -m lunar_converter convert``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .conversion_service import ConversionOptions, run_conversion
from .exceptions import LunarConverterError
from .logging_setup import configure_logging, shutdown_logging

EXIT_OK = 0
EXIT_FILE_ERRORS = 1
EXIT_GLOBAL_ERROR = 2


def project_root() -> Path:
    """Projektwurzel (enthält ``config/`` und ``schema/``) bei Ausführung aus dem Repository."""
    return Path(__file__).resolve().parents[2]


def _default_path(relative: str) -> Path:
    candidate = project_root() / relative
    return candidate if candidate.exists() else Path.cwd() / relative


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m lunar_converter",
        description="Konvertiert LUNAR-Testdokumente (.docx/.doc) in Import-JSON für den vorhandenen Jira/Xray-Importprozess.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True, metavar="BEFEHL")

    convert = subparsers.add_parser(
        "convert",
        help="Konvertiert alle .docx- und .doc-Dateien eines Eingabeordners.",
        description=(
            "Verarbeitet alle .docx- und .doc-Dateien des Eingabeordners einschließlich aller Unterordner, "
            "sortiert nach relativem Pfad. .doc-Dateien werden "
            "automatisch mit Microsoft Word nach .docx umgewandelt. Je "
            "Datei wird das passende Profil automatisch erkannt und ein Testfallordner mit testcase.json und "
            "screenshots/ erzeugt. "
            "Exit-Codes: 0 = alle Dateien erfolgreich, 1 = mindestens eine Datei fehlgeschlagen, "
            "2 = globaler Fehler (z. B. Output-Ordner nicht leer, ungültige Konfiguration)."
        ),
    )
    convert.add_argument("--input-dir", required=True, type=Path, help="Eingabeordner mit .docx-/.doc-Dateien (inkl. Unterordner).")
    convert.add_argument("--output-dir", required=True, type=Path, help="Output-Basisordner; muss leer sein oder wird angelegt.")
    convert.add_argument(
        "--profile",
        default=None,
        metavar="PROFIL_ID",
        help="Nur dieses Profil prüfen (Debugging/Migration). Ohne Angabe: automatische Erkennung über alle Profile.",
    )
    convert.add_argument("--log-level", choices=["DEBUG", "INFO", "WARNING", "ERROR"], default="INFO", help="Log-Level (Standard: INFO).")
    convert.add_argument(
        "--dry-run",
        action="store_true",
        help="Nur analysieren und validieren; keine Testfallordner, kein testcase.json, keine Screenshots schreiben.",
    )
    convert.add_argument("--fail-fast", action="store_true", help="Beim ersten Dateifehler abbrechen.")
    convert.add_argument("--schema-path", type=Path, default=None, help="JSON-Schema (Standard: schema/testcase.schema.json).")
    convert.add_argument("--config-dir", type=Path, default=None, help="Konfigurationsordner mit profiles/ (Standard: config/).")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    logger = configure_logging(args.log_level)
    options = ConversionOptions(
        input_dir=args.input_dir,
        output_dir=args.output_dir,
        schema_path=args.schema_path or _default_path("schema/testcase.schema.json"),
        config_dir=args.config_dir or _default_path("config"),
        profile_id=args.profile,
        dry_run=args.dry_run,
        fail_fast=args.fail_fast,
    )
    try:
        report = run_conversion(options, logger)
    except LunarConverterError as exc:
        logger.error("Lauf abgebrochen (%s): %s", exc.code, exc.message)
        if exc.details:
            logger.error("Details: %s", exc.details)
        return EXIT_GLOBAL_ERROR
    finally:
        shutdown_logging(logger)

    summary = (
        f"Ergebnis: {len(report.files)} Datei(en), {report.count('success')} erfolgreich, "
        f"{report.count('failed')} fehlgeschlagen, {report.count('skipped')} übersprungen."
    )
    print(summary)
    print(f"Report: {options.output_dir / 'conversion-report.json'}")
    if report.has_failures:
        for item in report.files:
            for issue in item.errors:
                field = f" [{issue.field}]" if issue.field else ""
                print(f"  FEHLER {item.input_file}: {issue.code}{field} – {issue.message}", file=sys.stderr)
        return EXIT_FILE_ERRORS
    return EXIT_OK
