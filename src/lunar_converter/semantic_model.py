"""Semantisches, toolneutrales Testfallmodell.

Das Modell kennt weder Word-Strukturen noch das Ziel-JSON. Bildpositionen
werden als ``ImageMarker`` innerhalb von ``RichText`` festgehalten, damit
erst der Ziel-Renderer entscheidet, ob und wie ein Anker geschrieben wird.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .models import Issue
from .text_normalizer import is_blank

TESTCASE_NAME_KEY = "testfallname"

STEP_NUMBER_KEY = "schritt_nummer"
STEP_SYSTEM_KEY = "system"
STEP_ACTION_KEY = "action"
STEP_DATA_KEY = "data"
STEP_EXPECTED_KEY = "expected_result"
STEP_ACTUAL_KEY = "actual_result"
STEP_COLUMN_KEYS: tuple[str, ...] = (
    STEP_NUMBER_KEY,
    STEP_SYSTEM_KEY,
    STEP_ACTION_KEY,
    STEP_DATA_KEY,
    STEP_EXPECTED_KEY,
    STEP_ACTUAL_KEY,
)
# Spalten, in denen Bilder eindeutig einem Textfeld des Schritts zugeordnet werden dürfen.
STEP_IMAGE_KEYS: tuple[str, ...] = (STEP_ACTION_KEY, STEP_EXPECTED_KEY)
# Einzige in Profilen erlaubte Alias-Schlüssel.
PROFILE_ALIAS_KEYS: tuple[str, ...] = (TESTCASE_NAME_KEY, *STEP_COLUMN_KEYS)


@dataclass(frozen=True)
class ImageMarker:
    """Position eines Bildes innerhalb eines Textfeldes."""

    image_id: int


@dataclass(frozen=True)
class CheckboxMarker:
    """Zustand einer Checkbox, deren Zieldarstellung der Renderer bestimmt."""

    checked: bool


Inline = str | ImageMarker | CheckboxMarker


@dataclass
class RichText:
    """Text mit Bildpositionen; jede Zeile entspricht einem Quellabsatz."""

    lines: list[list[Inline]] = field(default_factory=list)

    @classmethod
    def from_text(cls, text: str) -> RichText:
        return cls(lines=[[text]]) if text else cls()

    def image_ids(self) -> list[int]:
        return [item.image_id for line in self.lines for item in line if isinstance(item, ImageMarker)]

    def plain_text(self) -> str:
        return "\n".join("".join(item for item in line if isinstance(item, str)) for line in self.lines)

    def is_empty(self) -> bool:
        has_checkbox = any(isinstance(item, CheckboxMarker) for line in self.lines for item in line)
        return not self.image_ids() and not has_checkbox and is_blank(self.plain_text())


@dataclass
class TestStep:
    __test__ = False  # kein pytest-Testfall

    index: int
    source_location: str
    step_number: str = ""
    system: str = ""
    data: str = ""
    action: RichText = field(default_factory=RichText)
    expected_result: RichText = field(default_factory=RichText)
    actual_result: RichText = field(default_factory=RichText)

    def rich_fields(self) -> dict[str, RichText]:
        return {
            STEP_ACTION_KEY: self.action,
            STEP_EXPECTED_KEY: self.expected_result,
        }


@dataclass(frozen=True)
class UnassignedImage:
    """Bild, das weder der Info-Tabelle noch einem Schritt sicher zugeordnet werden kann."""

    image_id: int
    reason: str
    location: str = ""


@dataclass
class InfoTable:
    """Tabelle mit zentralen Testfallinformationen; jede Zelle als ``RichText``."""

    rows: list[list[RichText]]
    source_location: str = ""

    def image_ids(self) -> list[int]:
        return [image_id for row in self.rows for cell in row for image_id in cell.image_ids()]


@dataclass
class TestCase:
    __test__ = False  # kein pytest-Testfall

    source_file: Path
    profile_id: str
    name: str
    steps: list[TestStep]
    labels: list[str] = field(default_factory=list)
    responsible_names: list[str] = field(default_factory=list)
    reporter_email: str | None = None
    # Geschäftsprozess, Geschäftsprozessszenario, optionale Unterprozesse (Deckblatt).
    process_path: list[str] = field(default_factory=list)
    # Optionale Testfallbeschreibung vom Deckblatt, einzeilig.
    title: str = ""
    info_table: InfoTable | None = None
    unassigned_images: list[UnassignedImage] = field(default_factory=list)
    ignored_image_ids: set[int] = field(default_factory=set)
    warnings: list[Issue] = field(default_factory=list)
