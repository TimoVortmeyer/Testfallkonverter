"""Automatische Profilerkennung anhand des Dokumentinhalts (nie anhand des Dateinamens).

Ein Profil passt genau dann, wenn
1. alle ``required_markers`` im normalisierten Dokumenttext vorkommen und
2. eine Tabelle alle Pflichtspalten der Schritttabelle abdeckt.

Befüllte Schrittzeilen sind optional.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from typing import Any, Literal

from .models import ProfileDefinition, SourceDocument
from .step_table import count_data_rows, find_step_tables
from .text_normalizer import contains_marker

DetectionStatus = Literal["matched", "no_matching_profile", "ambiguous_profile"]


@dataclass
class ProfileCheckResult:
    profile: str
    matched: bool
    reasons: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"profile": self.profile, "matched": self.matched, "reasons": list(self.reasons)}


@dataclass
class DetectionResult:
    status: DetectionStatus
    profile: ProfileDefinition | None
    checks: list[ProfileCheckResult]

    @property
    def matching_profile_ids(self) -> list[str]:
        return [check.profile for check in self.checks if check.matched]


def check_profile(document: SourceDocument, profile: ProfileDefinition) -> ProfileCheckResult:
    reasons: list[str] = []
    normalized_text = document.normalized_text()
    missing_markers = [marker for marker in profile.required_markers if not contains_marker(normalized_text, marker)]
    if missing_markers:
        reasons.append("Marker fehlen: " + ", ".join(f"'{marker}'" for marker in missing_markers))
    else:
        reasons.append("Alle Marker gefunden.")

    layouts = find_step_tables(document, profile)
    if not layouts:
        reasons.append(
            "Keine Tabelle mit allen Pflichtspalten gefunden: " + ", ".join(f"'{c}'" for c in profile.required_columns)
        )
        return ProfileCheckResult(profile=profile.id, matched=False, reasons=reasons)

    table_numbers = ", ".join(str(layout.table.table_index) for layout in layouts)
    reasons.append(f"Schritttabelle gefunden (Tabelle {table_numbers}).")
    data_rows = sum(count_data_rows(layout) for layout in layouts)
    if data_rows == 0:
        reasons.append("Keine fachlich befüllte Schrittzeile vorhanden (Testschritte sind optional).")
    else:
        reasons.append(f"{data_rows} fachlich befüllte Schrittzeile(n) gefunden.")
    return ProfileCheckResult(profile=profile.id, matched=not missing_markers, reasons=reasons)


def detect_profile(document: SourceDocument, profiles: Sequence[ProfileDefinition]) -> DetectionResult:
    """Prüft das Dokument gegen alle übergebenen Profile. Mehrdeutigkeit wird nie willkürlich aufgelöst."""
    checks = [check_profile(document, profile) for profile in profiles]
    matching = [profile for profile, check in zip(profiles, checks) if check.matched]
    if len(matching) == 1:
        return DetectionResult(status="matched", profile=matching[0], checks=checks)
    if not matching:
        return DetectionResult(status="no_matching_profile", profile=None, checks=checks)
    return DetectionResult(status="ambiguous_profile", profile=None, checks=checks)
