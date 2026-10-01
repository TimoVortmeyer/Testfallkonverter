"""Technisches Quellmodell, Profildefinition und gemeinsame Befundstruktur.

Das technische Quellmodell bildet die DOCX-Struktur ab (Absätze, Tabellen,
Zellen, Textsegmente, Bildreferenzen in Dokumentreihenfolge) und enthält
bewusst keine fachliche Interpretation.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal

from .text_normalizer import normalize_for_matching


@dataclass(frozen=True)
class Issue:
    """Fehler oder Warnung für Log und Conversion-Report."""

    code: str
    message: str
    field: str | None = None
    details: str | None = None

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {"code": self.code, "message": self.message}
        if self.field is not None:
            result["field"] = self.field
        if self.details is not None:
            result["details"] = self.details
        return result


@dataclass(frozen=True)
class ProfileDefinition:
    """Schlankes, deklaratives Dokumentprofil."""

    id: str
    name: str
    required_markers: tuple[str, ...]
    required_columns: tuple[str, ...]
    field_aliases: dict[str, tuple[str, ...]]
    # Leer: letzte Tabelle vor der Schritttabelle gilt als Info-Tabelle.
    info_table_labels: tuple[str, ...] = ()
    source_path: Path | None = None

    def aliases_for(self, key: str) -> tuple[str, ...]:
        return self.field_aliases.get(key, ())


ImageKind = Literal["drawingml", "vml"]


@dataclass(frozen=True)
class CheckboxRef:
    """Zustand einer Word-Checkbox im Dokumentfluss."""

    checked: bool


@dataclass(frozen=True)
class ImageRef:
    """Eine sichtbare Bildplatzierung im Dokumentkörper.

    ``image_id`` ist die 1-basierte Position in natürlicher Dokumentreihenfolge.
    Wird dieselbe Mediendatei mehrfach platziert, entsteht je Platzierung eine
    eigene ``ImageRef``.
    """

    image_id: int
    relationship_id: str
    kind: ImageKind
    floating: bool
    location: str
    part_name: str | None = None
    blob: bytes | None = field(default=None, repr=False)
    external: bool = False


Segment = str | ImageRef | CheckboxRef


@dataclass
class SourceParagraph:
    segments: list[Segment]
    style_id: str | None = None
    is_heading: bool = False

    @property
    def text(self) -> str:
        return "".join(segment for segment in self.segments if isinstance(segment, str))

    @property
    def images(self) -> list[ImageRef]:
        return [segment for segment in self.segments if isinstance(segment, ImageRef)]


VMerge = Literal["restart", "continue"]


@dataclass
class SourceCell:
    paragraphs: list[SourceParagraph]
    grid_column: int
    grid_span: int = 1
    vmerge: VMerge | None = None

    @property
    def text(self) -> str:
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    @property
    def images(self) -> list[ImageRef]:
        return [image for paragraph in self.paragraphs for image in paragraph.images]

    def covers(self, grid_column: int) -> bool:
        return self.grid_column <= grid_column < self.grid_column + self.grid_span


@dataclass
class SourceRow:
    cells: list[SourceCell]
    is_header: bool = False

    def cell_at(self, grid_column: int) -> SourceCell | None:
        for cell in self.cells:
            if cell.covers(grid_column):
                return cell
        return None


@dataclass
class SourceTable:
    table_index: int
    rows: list[SourceRow]


SourceBlock = SourceParagraph | SourceTable


@dataclass
class SourceDocument:
    """Technisches Abbild des Dokumentkörpers (ohne Kopf- und Fußzeilen)."""

    source_path: Path
    blocks: list[SourceBlock]
    images: list[ImageRef]
    # Nur zur Protokollierung: Bilder aus Kopf-/Fußzeilen werden nie extrahiert.
    ignored_header_footer_images: int = 0

    @property
    def tables(self) -> list[SourceTable]:
        return [block for block in self.blocks if isinstance(block, SourceTable)]

    def iter_paragraphs(self) -> list[SourceParagraph]:
        paragraphs: list[SourceParagraph] = []
        for block in self.blocks:
            if isinstance(block, SourceParagraph):
                paragraphs.append(block)
            else:
                for row in block.rows:
                    for cell in row.cells:
                        paragraphs.extend(cell.paragraphs)
        return paragraphs

    def normalized_text(self) -> str:
        return " ".join(
            normalized for normalized in (normalize_for_matching(p.text) for p in self.iter_paragraphs()) if normalized
        )
