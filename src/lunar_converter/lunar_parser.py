"""Fachliche Extraktion: technisches Quellmodell -> semantisches Testfallmodell.

Aufbau eines Testdokuments:

* Deckblatt: Absatzfolge mit ``Testfall: <Name>`` (Alias ``testfallname``). Davor
  stehen die nummerierten Prozesszeilen (z. B. ``03.02 Einkaufsverwaltung``,
  ``03.02.001 ...``, optional Unterprozesse) und danach optional eine
  Testfallbeschreibung bis zur Testfall-Zeile.
* Tabelle mit zentralen Informationen: letzte Tabelle vor der Schritttabelle
  (ohne Deckblatt-Tabelle); wird vollständig übernommen.
* Schritttabelle: Spalten über die ``field_aliases`` des Profils.

Bilder, die nicht sicher zugeordnet werden können, landen in
``TestCase.unassigned_images`` (Sicherheitsprinzip: Testfall-Anhang statt
Schritt-Anhang).
"""

from __future__ import annotations

import re
from collections.abc import Iterator, Sequence
from dataclasses import dataclass

from .models import CheckboxRef, ImageRef, Issue, ProfileDefinition, Segment, SourceCell, SourceDocument, SourceParagraph, SourceRow, SourceTable
from .responsibles import cover_responsibles
from .semantic_model import (
    STEP_ACTION_KEY,
    STEP_ACTUAL_KEY,
    STEP_DATA_KEY,
    STEP_EXPECTED_KEY,
    STEP_NUMBER_KEY,
    STEP_SYSTEM_KEY,
    TESTCASE_NAME_KEY,
    CheckboxMarker,
    ImageMarker,
    InfoTable,
    Inline,
    RichText,
    TestCase,
    TestStep,
    UnassignedImage,
)
from .step_table import StepTableLayout, classify_row, find_step_tables
from .text_normalizer import clean_line, clean_multiline, is_blank, normalize_heading, normalize_newlines

_LABEL_SEPARATOR_RE = re.compile(r"[:\t]")
# Prozessnummer wie "03.02", "03.02.001" oder "03.02.001.01", gefolgt von einer Bezeichnung.
_PROCESS_LINE_RE = re.compile(r"^\d{2}(?:\.\d+)+\s+\S")
_DEFAULT_UNASSIGNED_REASON = "Bild liegt außerhalb eines eindeutig zuordenbaren Textfeldes"
_FLOATING_REASON = "Frei positioniertes Bild; die Position zum Text ist nicht eindeutig"
_NAME_IMAGE_REASON = "Bild im Testfallnamen; die Summary kann keine Bilder aufnehmen"

Lines = list[list[Segment]]


@dataclass
class _Cover:
    name_lines: Lines
    process_path: list[str]
    title: str
    table: SourceTable | None


def parse_test_case(
    document: SourceDocument,
    profile: ProfileDefinition,
    layouts: Sequence[StepTableLayout] | None = None,
) -> TestCase:
    """Extrahiert Deckblatt, Info-Tabelle und Testschritte gemäß Profil."""
    return _TestCaseParser(document, profile, list(layouts) if layouts is not None else find_step_tables(document, profile)).parse()


class _TestCaseParser:
    def __init__(self, document: SourceDocument, profile: ProfileDefinition, layouts: list[StepTableLayout]) -> None:
        self._document = document
        self._profile = profile
        self._layouts = layouts
        self._name_aliases = {normalize_heading(alias) for alias in profile.aliases_for(TESTCASE_NAME_KEY)} - {""}
        self._image_reasons: dict[int, str] = {}
        self._warnings: list[Issue] = []

    def parse(self) -> TestCase:
        cover = self._find_cover()
        info_table = self._build_info_table(cover.table if cover else None)
        steps = self._collect_steps()
        name = self._plain_text(cover.name_lines, reason=_NAME_IMAGE_REASON, single_line=True) if cover else ""

        assigned: set[int] = set(info_table.image_ids()) if info_table else set()
        for step in steps:
            for rich in step.rich_fields().values():
                assigned.update(rich.image_ids())
        unassigned = [
            UnassignedImage(
                image_id=image.image_id,
                reason=self._image_reasons.get(image.image_id, _DEFAULT_UNASSIGNED_REASON),
                location=image.location,
            )
            for image in self._document.images
            if image.image_id not in assigned
        ]
        return TestCase(
            source_file=self._document.source_path,
            profile_id=self._profile.id,
            name=name,
            steps=steps,
            labels=_cover_labels(cover.table) if cover else [],
            responsible_names=cover_responsibles(self._document),
            process_path=cover.process_path if cover else [],
            title=cover.title if cover else "",
            info_table=info_table,
            unassigned_images=unassigned,
            warnings=self._warnings,
        )

    # --- Deckblatt ------------------------------------------------------------------

    def _containers(self) -> Iterator[tuple[list[SourceParagraph], SourceTable | None]]:
        """Absatzfolgen au\u00dferhalb der Schritttabellen: Flie\u00dftext zwischen Tabellen und einzelne Tabellenzellen."""
        step_tables = {id(layout.table) for layout in self._layouts}
        container: list[SourceParagraph] = []
        for block in self._document.blocks:
            if isinstance(block, SourceParagraph):
                container.append(block)
                continue
            if container:
                yield container, None
            container = []
            if id(block) in step_tables:
                continue
            for row in block.rows:
                for cell in row.cells:
                    yield cell.paragraphs, block
        if container:
            yield container, None

    def _find_cover(self) -> _Cover | None:
        for paragraphs, table in self._containers():
            for index in range(len(paragraphs)):
                name_lines = self._name_value(paragraphs, index)
                if name_lines is None:
                    continue
                process_path, title = _cover_header(paragraphs[:index])
                return _Cover(name_lines=name_lines, process_path=process_path, title=title, table=table)
        return None

    def _name_value(self, paragraphs: list[SourceParagraph], index: int) -> Lines | None:
        """Liefert den Testfallnamen aus ``Testfall: <Name>`` bzw. aus dem Absatz nach einem reinen Bezeichner."""
        paragraph = paragraphs[index]
        if not paragraph.images and normalize_heading(paragraph.text) in self._name_aliases:
            for following in paragraphs[index + 1 :]:
                if is_blank(following.text):
                    continue
                # Eine Folgezeile "Bezeichner: Wert" ist ein anderes Feld, nicht der Testfallname.
                return None if _LABEL_SEPARATOR_RE.search(following.text) else [following.segments]
            return None
        prefix = ""
        for position, segment in enumerate(paragraph.segments):
            if isinstance(segment, ImageRef):
                return None
            match = _LABEL_SEPARATOR_RE.search(segment)
            if match is None:
                prefix += segment
                continue
            if normalize_heading(prefix + segment[: match.start()]) not in self._name_aliases:
                return None
            rest = segment[match.end() :]
            value: list[Segment] = ([rest] if rest else []) + list(paragraph.segments[position + 1 :])
            return [value] if _lines_have_content([value]) else None
        return None

    # --- Tabelle mit zentralen Informationen -------------------------------------------

    def _build_info_table(self, cover_table: SourceTable | None) -> InfoTable | None:
        """Tabelle vor der ersten Schritttabelle (ohne Deckblatt-Tabelle).

        Mit ``info_table.required_labels`` im Profil: erste Tabelle, die alle Bezeichner als Zelltext enthält.
        Sonst: letzte Tabelle vor der Schritttabelle.
        """
        if not self._layouts:
            return None
        step_tables = {id(layout.table) for layout in self._layouts}
        candidates: list[SourceTable] = []
        for block in self._document.blocks:
            if not isinstance(block, SourceTable):
                continue
            if id(block) in step_tables:
                break
            if block is not cover_table:
                candidates.append(block)
        labels = {normalize_heading(label) for label in self._profile.info_table_labels}
        if labels:
            candidates = [table for table in candidates if labels <= _cell_headings(table)][:1]
            if not candidates:
                self._warnings.append(
                    Issue(
                        code="info_table_not_found",
                        message="Keine Tabelle mit den Bezeichnern "
                        + ", ".join(f"'{label}'" for label in self._profile.info_table_labels)
                        + " vor dem Testablauf gefunden; die Description enthält keine Tabelle.",
                    )
                )
        if not candidates:
            return None
        table = candidates[-1]
        rows: list[list[RichText]] = []
        for row in table.rows:
            cells = [
                RichText() if cell.vmerge == "continue" else self._rich_text([p.segments for p in cell.paragraphs])
                for cell in row.cells
            ]
            if any(not cell.is_empty() for cell in cells):
                rows.append(cells)
        return InfoTable(rows=rows, source_location=f"Tabelle {table.table_index}") if rows else None

    # --- Testschritte ---------------------------------------------------------------

    def _collect_steps(self) -> list[TestStep]:
        steps: list[TestStep] = []
        for layout in self._layouts:
            merge_origins: dict[int, SourceCell] = {}
            for row_index, row in layout.data_rows():
                location = f"Tabelle {layout.table.table_index}, Zeile {row_index + 1}"
                kind = classify_row(layout, row)
                if kind == "layout":
                    self._warnings.append(
                        Issue(
                            code="layout_row_ignored",
                            message=f"{location}: Zeile mit über mehrere Spalten verbundener Zelle wird als Layout-Zeile ignoriert.",
                        )
                    )
                if kind == "data":
                    steps.append(self._build_step(layout, row, len(steps), location, merge_origins))
                for cell in row.cells:
                    if cell.vmerge != "continue":
                        for grid_column in range(cell.grid_column, cell.grid_column + cell.grid_span):
                            merge_origins[grid_column] = cell
        return steps

    def _build_step(
        self,
        layout: StepTableLayout,
        row: SourceRow,
        index: int,
        location: str,
        merge_origins: dict[int, SourceCell],
    ) -> TestStep:
        step = TestStep(index=index, source_location=location)
        for key, grid_column in layout.columns.items():
            cell = row.cell_at(grid_column)
            if cell is None:
                continue
            if cell.vmerge == "continue":
                # Vertikal verbundene Zelle: Text der Ursprungszelle übernehmen, Bilder nicht duplizieren.
                origin = merge_origins.get(grid_column)
                lines: Lines = [[origin.text]] if origin is not None else []
            else:
                lines = [paragraph.segments for paragraph in cell.paragraphs]
            if key == STEP_NUMBER_KEY:
                step.step_number = self._plain_text(lines, reason=_DEFAULT_UNASSIGNED_REASON, single_line=True)
            elif key == STEP_SYSTEM_KEY:
                step.system = self._plain_text(lines, reason=_DEFAULT_UNASSIGNED_REASON, single_line=False)
            elif key == STEP_DATA_KEY:
                step.data = self._plain_text(lines, reason=_DEFAULT_UNASSIGNED_REASON, single_line=False)
            elif key == STEP_ACTION_KEY:
                step.action = self._rich_text(lines)
            elif key == STEP_EXPECTED_KEY:
                step.expected_result = self._rich_text(lines)
            elif key == STEP_ACTUAL_KEY:
                step.actual_result = self._rich_text(lines)
        return step

    # --- Hilfsfunktionen ---------------------------------------------------------------

    def _rich_text(self, lines: Lines) -> RichText:
        converted: list[list[Inline]] = []
        for line in lines:
            items: list[Inline] = []
            for segment in line:
                if isinstance(segment, str):
                    items.append(segment)
                elif isinstance(segment, CheckboxRef):
                    items.append(CheckboxMarker(segment.checked))
                elif segment.floating:
                    self._image_reasons.setdefault(segment.image_id, _FLOATING_REASON)
                else:
                    items.append(ImageMarker(segment.image_id))
            converted.append(items)
        while converted and _inline_line_blank(converted[0]):
            converted.pop(0)
        while converted and _inline_line_blank(converted[-1]):
            converted.pop()
        return RichText(lines=converted)

    def _plain_text(self, lines: Lines, *, reason: str, single_line: bool) -> str:
        texts: list[str] = []
        for line in lines:
            for segment in line:
                if isinstance(segment, ImageRef):
                    self._image_reasons.setdefault(segment.image_id, reason)
            texts.append("".join(segment for segment in line if isinstance(segment, str)))
        if single_line:
            return clean_line(" ".join(texts))
        return clean_multiline("\n".join(texts))


def _cover_header(paragraphs: list[SourceParagraph]) -> tuple[list[str], str]:
    """Prozesspfad und Beschreibung vor der Testfall-Zeile aus dem Deckblatt lesen."""
    texts = [clean_line(normalize_newlines(paragraph.text).replace("\n", " ")) for paragraph in paragraphs]
    process_indices = [index for index, text in enumerate(texts) if _PROCESS_LINE_RE.match(text)]
    if not process_indices:
        return [], ""
    process_path = [texts[index] for index in process_indices]
    trailing_levels = [text for text in texts[process_indices[-1] + 1 :] if text]
    process_path.extend(trailing_levels)
    title = " ".join(trailing_levels)
    return process_path, title


def _cover_labels(table: SourceTable | None) -> list[str]:
    if table is None or not table.rows or not table.rows[0].cells:
        return []
    for paragraph in table.rows[0].cells[0].paragraphs:
        label = clean_line(paragraph.text)
        if label:
            return [label]
    return []


def _cell_headings(table: SourceTable) -> set[str]:
    return {normalize_heading(cell.text) for row in table.rows for cell in row.cells}


def _lines_have_content(lines: Lines) -> bool:
    return any(isinstance(segment, (ImageRef, CheckboxRef)) or not is_blank(segment) for line in lines for segment in line)


def _inline_line_blank(line: list[Inline]) -> bool:
    return all(isinstance(item, str) and is_blank(item) for item in line)
