"""Logging-Konfiguration mit Kontext (Quelldatei, Profil) je Logzeile."""

from __future__ import annotations

import logging
from collections.abc import Mapping
from pathlib import Path
from typing import Any

LOGGER_NAME = "lunar_converter"
LOG_FILE_NAME = "conversion.log"
LOG_FORMAT = "%(asctime)s %(levelname)-7s [Datei: %(source_file)s] [Profil: %(profile)s] %(message)s"
DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
_NO_CONTEXT = "-"


class _ContextDefaultsFilter(logging.Filter):
    """Ergänzt fehlende Kontextfelder, damit das Format für jede Logzeile funktioniert."""

    def filter(self, record: logging.LogRecord) -> bool:
        if not hasattr(record, "source_file"):
            record.source_file = _NO_CONTEXT
        if not hasattr(record, "profile"):
            record.profile = _NO_CONTEXT
        return True


class ContextLogger(logging.LoggerAdapter[logging.Logger]):
    """Logger mit fest gebundener Quelldatei und Profil."""

    def process(self, msg: Any, kwargs: Any) -> tuple[Any, Any]:
        extra: dict[str, Any] = dict(self.extra or {})
        extra.update(kwargs.get("extra") or {})
        kwargs["extra"] = extra
        return msg, kwargs


def context_logger(logger: logging.Logger, source_file: str | None = None, profile: str | None = None) -> ContextLogger:
    context: Mapping[str, Any] = {"source_file": source_file or _NO_CONTEXT, "profile": profile or _NO_CONTEXT}
    return ContextLogger(logger, context)


def _make_handler(handler: logging.Handler, level: int) -> logging.Handler:
    handler.setLevel(level)
    handler.setFormatter(logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT))
    handler.addFilter(_ContextDefaultsFilter())
    return handler


def configure_logging(log_level: str = "INFO") -> logging.Logger:
    """Richtet den Paket-Logger mit Konsolenausgabe (stderr) ein; vorhandene Handler werden ersetzt."""
    level = logging.getLevelName(log_level.upper())
    if not isinstance(level, int):
        level = logging.INFO
    logger = logging.getLogger(LOGGER_NAME)
    shutdown_logging(logger)
    logger.setLevel(level)
    logger.addHandler(_make_handler(logging.StreamHandler(), level))
    return logger


def attach_log_file(logger: logging.Logger, output_dir: Path) -> Path:
    """Schreibt ab sofort zusätzlich nach ``<output_dir>/conversion.log`` (UTF-8)."""
    path = output_dir / LOG_FILE_NAME
    logger.addHandler(_make_handler(logging.FileHandler(path, encoding="utf-8"), logger.level))
    return path


def shutdown_logging(logger: logging.Logger) -> None:
    """Schließt alle Handler (wichtig unter Windows, damit Logdateien freigegeben werden)."""
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()
