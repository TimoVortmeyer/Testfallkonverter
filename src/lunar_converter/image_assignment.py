"""Zuordnung exportierter Bilder zu Testfall bzw. Testschritten.

Regeln:
* Bild in der Tabelle mit zentralen Informationen -> globaler Screenshot, Anker in ``description``.
* Bild in Aktion/Erwartetem/Tatsächlichem Ergebnis eines Schritts -> Schritt-Anhang mit Anker.
* Alles andere -> globaler Screenshot ohne Anker, Warnung.
Ein Schritt-Anhang erscheint nie zusätzlich in den globalen Screenshots.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Literal

from .models import Issue
from .semantic_model import TestCase

AssignmentTarget = Literal["testcase_description", "step", "testcase_unassigned"]


@dataclass(frozen=True)
class AssignedImage:
    image_id: int
    file_name: str
    target: AssignmentTarget
    anchored: bool
    field_key: str | None = None
    step_index: int | None = None


@dataclass
class ImageAssignmentResult:
    assignments: list[AssignedImage] = field(default_factory=list)
    warnings: list[Issue] = field(default_factory=list)

    @property
    def anchors(self) -> dict[int, str]:
        return {item.image_id: item.file_name for item in self.assignments if item.anchored}

    @property
    def global_screenshots(self) -> list[str]:
        return sorted(item.file_name for item in self.assignments if item.target != "step")

    def step_attachments(self, step_index: int) -> list[str]:
        return sorted(item.file_name for item in self.assignments if item.target == "step" and item.step_index == step_index)


def assign_images(test_case: TestCase, exported: Mapping[int, str]) -> ImageAssignmentResult:
    """Ordnet ausschließlich erfolgreich exportierte Bilder zu."""
    result = ImageAssignmentResult()
    seen: set[int] = set()

    def add(image_id: int, target: AssignmentTarget, *, field_key: str | None = None, step_index: int | None = None) -> None:
        if image_id not in exported or image_id in seen:
            return
        seen.add(image_id)
        result.assignments.append(
            AssignedImage(
                image_id=image_id,
                file_name=exported[image_id],
                target=target,
                anchored=target != "testcase_unassigned",
                field_key=field_key,
                step_index=step_index,
            )
        )

    if test_case.info_table is not None:
        for image_id in test_case.info_table.image_ids():
            add(image_id, "testcase_description")
    for step in test_case.steps:
        for key, rich in step.rich_fields().items():
            for image_id in rich.image_ids():
                add(image_id, "step", field_key=key, step_index=step.index)
    for unassigned in test_case.unassigned_images:
        if unassigned.image_id in exported and unassigned.image_id not in seen:
            add(unassigned.image_id, "testcase_unassigned")
            result.warnings.append(
                Issue(
                    code="image_assignment_unclear",
                    message=(
                        f"Bild {exported[unassigned.image_id]} ({unassigned.location}) konnte nicht eindeutig zugeordnet werden "
                        f"({unassigned.reason}); es wird als Testfall-Screenshot ohne Anker übernommen."
                    ),
                )
            )
    result.assignments.sort(key=lambda item: item.file_name)
    return result
