"""Orchestrierung der Batch-Konvertierung.

Pro Quelldatei: DOC-Konvertierung (Word) -> DOCX-Extraktion -> Profilerkennung ->
semantisches Modell -> Bildexport und -zuordnung -> Rendering -> Validierung ->
atomarer Export. Fehler einer Datei werden isoliert und im Report festgehalten.
"""

from __future__ import annotations

import logging
import tempfile
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .doc_converter import convert_doc_to_docx
from .docx_reader import read_docx
from .exceptions import ConfigurationError, ExportError, LunarConverterError, ProfileDetectionError
from .filesystem import prepare_output_directory, publish_directory, sanitize_folder_name
from .image_assignment import assign_images
from .image_extractor import export_images
from .logging_setup import ContextLogger, attach_log_file, context_logger
from .lunar_parser import parse_test_case
from .models import Issue, ProfileDefinition
from .profile_detector import DetectionResult, detect_profile
from .profile_loader import load_profiles
from .reporting import BatchReport, FileReport, now_iso, write_report
from .source_discovery import discover_source_files, relative_display_path
from .target_renderer import XrayImportRenderer, write_testcase_json
from .validator import load_schema, validate_payload

TESTCASE_FILE_NAME = "testcase.json"
SCREENSHOTS_DIR_NAME = "screenshots"


@dataclass(frozen=True)
class ConversionOptions:
    input_dir: Path
    output_dir: Path
    schema_path: Path
    config_dir: Path
    profile_id: str | None = None
    dry_run: bool = False
    fail_fast: bool = False


def run_conversion(options: ConversionOptions, logger: logging.Logger) -> BatchReport:
    """Führt den gesamten Lauf aus.

    Globale Probleme (Eingabeordner, Profile, Schema, nicht leerer Output-Ordner)
    lösen vor jeder Dateiverarbeitung eine ``LunarConverterError`` aus.
    """
    if not options.input_dir.is_dir():
        raise ConfigurationError(f"Eingabeordner '{options.input_dir}' existiert nicht oder ist kein Ordner.")
    profiles = _select_profiles(load_profiles(options.config_dir), options.profile_id)
    schema = load_schema(options.schema_path)
    prepare_output_directory(options.output_dir)
    attach_log_file(logger, options.output_dir)

    report = BatchReport(
        input_dir=str(options.input_dir),
        output_dir=str(options.output_dir),
        dry_run=options.dry_run,
        profile_override=options.profile_id,
    )
    converter = _FileConverter(options, profiles, schema, logger)
    try:
        files = discover_source_files(options.input_dir, exclude_dir=options.output_dir)
        logger.info(
            "Start der Konvertierung: %d Datei(en) in '%s' (inkl. Unterordner), Profile: %s%s.",
            len(files),
            options.input_dir,
            ", ".join(profile.id for profile in profiles),
            " (Dry Run)" if options.dry_run else "",
        )
        if not files:
            logger.warning("Im Eingabeordner und seinen Unterordnern wurden keine .docx- oder .doc-Dateien gefunden.")
        # Ordnername (casefold) -> Quelldatei, die ihn belegt; der Output bleibt flach.
        used_folder_names: dict[str, str] = {}
        for position, source in enumerate(files):
            file_report = converter.convert(source, used_folder_names)
            report.files.append(file_report)
            if file_report.status == "failed" and options.fail_fast:
                for remaining in files[position + 1 :]:
                    report.files.append(_skipped_report(remaining))
                    context_logger(logger, relative_display_path(remaining, options.input_dir)).warning(
                        "Übersprungen wegen --fail-fast."
                    )
                logger.error("Abbruch nach erstem Dateifehler (--fail-fast).")
                break
    finally:
        report.finished_at = now_iso()
        report_path = write_report(report, options.output_dir)
        logger.info(
            "Ende der Konvertierung: %d gesamt, %d erfolgreich, %d fehlgeschlagen, %d übersprungen. Report: %s",
            len(report.files),
            report.count("success"),
            report.count("failed"),
            report.count("skipped"),
            report_path,
        )
    return report


def _select_profiles(profiles: list[ProfileDefinition], profile_id: str | None) -> list[ProfileDefinition]:
    if profile_id is None:
        return profiles
    selected = [profile for profile in profiles if profile.id == profile_id]
    if not selected:
        available = ", ".join(profile.id for profile in profiles)
        raise ConfigurationError(f"Profil '{profile_id}' ist nicht vorhanden. Verfügbare Profile: {available}.")
    return selected


def _skipped_report(source: Path) -> FileReport:
    return FileReport(
        input_file=str(source),
        status="skipped",
        warnings=[Issue(code="skipped_fail_fast", message="Nicht verarbeitet, da --fail-fast nach einem Fehler abgebrochen hat.")],
    )


class _FileConverter:
    def __init__(
        self,
        options: ConversionOptions,
        profiles: Sequence[ProfileDefinition],
        schema: dict[str, Any],
        logger: logging.Logger,
    ) -> None:
        self._options = options
        self._profiles = tuple(profiles)
        self._schema = schema
        self._logger = logger
        self._renderer = XrayImportRenderer()

    def convert(self, source: Path, used_folder_names: dict[str, str]) -> FileReport:
        report = FileReport(input_file=str(source))
        context_logger(self._logger, self._label(source)).info("Verarbeitung gestartet.")
        try:
            with tempfile.TemporaryDirectory(prefix="lunar_", ignore_cleanup_errors=True) as temp:
                self._convert(source, Path(temp), report, used_folder_names)
            report.status = "success"
        except LunarConverterError as exc:
            self._fail(report, source, exc.to_issues())
        except Exception as exc:  # Fehlerisolierung: unerwartete Fehler dürfen den Batch nicht abbrechen.
            issue = Issue(
                code="unexpected_error",
                message="Unerwarteter Fehler bei der Verarbeitung.",
                details=f"{type(exc).__name__}: {exc}",
            )
            self._fail(report, source, [issue])
        return report

    def _label(self, source: Path) -> str:
        return relative_display_path(source, self._options.input_dir)

    def _convert(self, source: Path, work_dir: Path, report: FileReport, used_folder_names: dict[str, str]) -> None:
        log = context_logger(self._logger, self._label(source))
        docx_path = source
        if source.suffix.lower() == ".doc":
            docx_path = convert_doc_to_docx(source, work_dir / "doc-konvertierung")
            log.info("DOC-Datei wurde mit Microsoft Word nach DOCX umgewandelt.")
        document = read_docx(docx_path)
        report.ignored_header_footer_image_count = document.ignored_header_footer_images
        if document.ignored_header_footer_images:
            log.info(
                "%d Bild(er) in Kopf-/Fußzeilen werden nicht extrahiert und nicht verankert.",
                document.ignored_header_footer_images,
            )

        detection = detect_profile(document, self._profiles)
        report.checked_profiles = detection.checks
        report.profile_detection_status = detection.status
        for check in detection.checks:
            log.debug("Profilprüfung '%s': %s – %s", check.profile, "passt" if check.matched else "passt nicht", " ".join(check.reasons))
        profile = _require_single_profile(detection)
        report.detected_profile = profile.id
        log = context_logger(self._logger, self._label(source), profile.id)
        log.info("Profil erkannt: %s (%s).", profile.id, profile.name)

        test_case = parse_test_case(document, profile)
        folder_name = sanitize_folder_name(source.stem)
        if folder_name.casefold() in used_folder_names:
            raise ExportError(
                f"Der Ordnername '{folder_name}' wird in diesem Lauf bereits von "
                f"'{used_folder_names[folder_name.casefold()]}' verwendet (gleicher Dateiname, ggf. in anderem Unterordner).",
                code="output_name_conflict",
            )

        package_dir = work_dir / "paket" / folder_name
        screenshots_dir = package_dir / SCREENSHOTS_DIR_NAME
        exported = export_images(document.images, screenshots_dir)
        assignment = assign_images(test_case, exported.exported)
        for warning in (*test_case.warnings, *exported.warnings, *assignment.warnings):
            report.warnings.append(warning)
            log.warning("%s", warning.message)
            if warning.details:
                log.debug("Details: %s", warning.details)

        payload = self._renderer.render(test_case, assignment)
        validate_payload(payload, self._schema, screenshots_dir)
        write_testcase_json(package_dir / TESTCASE_FILE_NAME, payload)

        final_dir = self._options.output_dir / folder_name
        if self._options.dry_run:
            report.planned_output_directory = str(final_dir)
            log.info(
                "Dry Run: Export wäre nach '%s' erfolgt (%d Schritt(e), %d Bild(er)); es wurde nichts geschrieben.",
                final_dir,
                len(payload["steps"]),
                len(exported.exported),
            )
        else:
            publish_directory(package_dir, final_dir)
            report.output_directory = str(final_dir)
            log.info(
                "Export erfolgreich nach '%s' (%d Schritt(e), %d Bild(er)).",
                final_dir,
                len(payload["steps"]),
                len(exported.exported),
            )
        used_folder_names[folder_name.casefold()] = self._label(source)
        report.step_count = len(payload["steps"])
        report.exported_image_count = len(exported.exported)

    def _fail(self, report: FileReport, source: Path, issues: list[Issue]) -> None:
        report.status = "failed"
        report.output_directory = None
        report.planned_output_directory = None
        report.errors.extend(issues)
        log: ContextLogger = context_logger(self._logger, self._label(source), report.detected_profile)
        for issue in issues:
            lines = [f"Datei: {source}", f"Profil: {report.detected_profile or '-'}", f"Fehler: {issue.code}"]
            if issue.field:
                lines.append(f"Feld: {issue.field}")
            lines.append(f"Ursache: {issue.message}")
            if issue.details:
                lines.append(f"Details: {issue.details}")
            log.error("\n".join(lines))
        log.debug("Stacktrace:", exc_info=True)


def _require_single_profile(detection: DetectionResult) -> ProfileDefinition:
    if detection.status == "matched" and detection.profile is not None:
        return detection.profile
    details = " | ".join(f"{check.profile}: {' '.join(check.reasons)}" for check in detection.checks)
    if detection.status == "ambiguous_profile":
        raise ProfileDetectionError(
            "Mehrere Profile passen zum Dokument: "
            + ", ".join(detection.matching_profile_ids)
            + ". Es wird kein Profil willkürlich gewählt.",
            code="ambiguous_profile",
            details=details,
        )
    raise ProfileDetectionError("Kein hinterlegtes Profil passt zum Dokument.", code="no_matching_profile", details=details)
