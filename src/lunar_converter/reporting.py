"""Conversion-Report (``conversion-report.json``)."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Literal

from .models import Issue
from .profile_detector import DetectionStatus, ProfileCheckResult

REPORT_FILE_NAME = "conversion-report.json"
FileStatus = Literal["success", "failed", "skipped"]


def now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


@dataclass
class FileReport:
    input_file: str
    status: FileStatus = "failed"
    profile_detection_status: DetectionStatus | None = None
    detected_profile: str | None = None
    checked_profiles: list[ProfileCheckResult] = field(default_factory=list)
    output_directory: str | None = None
    planned_output_directory: str | None = None
    step_count: int = 0
    exported_image_count: int = 0
    ignored_header_footer_image_count: int = 0
    warnings: list[Issue] = field(default_factory=list)
    errors: list[Issue] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "input_file": self.input_file,
            "status": self.status,
            "profile_detection_status": self.profile_detection_status,
            "detected_profile": self.detected_profile,
            "checked_profiles": [check.to_dict() for check in self.checked_profiles],
            "output_directory": self.output_directory,
            "planned_output_directory": self.planned_output_directory,
            "step_count": self.step_count,
            "exported_image_count": self.exported_image_count,
            "ignored_header_footer_image_count": self.ignored_header_footer_image_count,
            "warnings": [issue.to_dict() for issue in self.warnings],
            "errors": [issue.to_dict() for issue in self.errors],
        }


@dataclass
class BatchReport:
    input_dir: str
    output_dir: str
    dry_run: bool = False
    profile_override: str | None = None
    started_at: str = field(default_factory=now_iso)
    finished_at: str | None = None
    files: list[FileReport] = field(default_factory=list)

    def count(self, status: FileStatus) -> int:
        return sum(1 for item in self.files if item.status == status)

    @property
    def has_failures(self) -> bool:
        return self.count("failed") > 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "started_at": self.started_at,
            "finished_at": self.finished_at,
            "input_dir": self.input_dir,
            "output_dir": self.output_dir,
            "dry_run": self.dry_run,
            "profile_override": self.profile_override,
            "summary": {
                "total": len(self.files),
                "success": self.count("success"),
                "failed": self.count("failed"),
                "skipped": self.count("skipped"),
            },
            "files": [item.to_dict() for item in self.files],
        }


def write_report(report: BatchReport, output_dir: Path) -> Path:
    """Schreibt den Report über eine temporäre Datei, damit nie ein halber Report entsteht."""
    path = output_dir / REPORT_FILE_NAME
    temporary = output_dir / f".{REPORT_FILE_NAME}.tmp"
    temporary.write_text(json.dumps(report.to_dict(), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    os.replace(temporary, path)
    return path
