"""Erkennung und Zeilenklassifikation der Testschritt-Tabelle.

Wird gemeinsam von der Profilerkennung und dem fachlichen Parser genutzt, damit
beide identische Regeln für Spaltenköpfe und befüllte Schrittzeilen verwenden.
"""

from __future__ import annotations

from collections.abc import Iterator
from dataclasses import dataclass
from typing import Literal

from .models import ProfileDefinition, SourceCell, SourceDocument, SourceRow, SourceTable
from .semantic_model import (
    STEP_ACTION_KEY,
    STEP_ACTUAL_KEY,
    STEP_COLUMN_KEYS,
    STEP_DATA_KEY,
    STEP_EXPECTED_KEY,
    STEP_SYSTEM_KEY,
)
from .text_normalizer import is_blank, normalize_heading

# Kopfzeile wird in den ersten Zeilen gesucht, um z. B. Titelzeilen über dem Tabellenkopf zu erlauben.
MAX_HEADER_SCAN_ROWS = 3
CONTENT_KEYS: tuple[str, ...] = (STEP_SYSTEM_KEY, STEP_ACTION_KEY, STEP_DATA_KEY, STEP_EXPECTED_KEY, STEP_ACTUAL_KEY)

RowKind = Literal["header", "empty", "layout", "data"]


@dataclass(frozen=True)
class StepTableLayout:
    table: SourceTable
    header_row_index: int
    columns: dict[str, int]
    header_texts: dict[str, str]

    def data_rows(self) -> Iterator[tuple[int, SourceRow]]:
        for index, row in enumerate(self.table.rows):
            if index > self.header_row_index:
                yield index, row


def column_aliases(profile: ProfileDefinition) -> dict[str, set[str]]:
    return {key: {normalize_heading(alias) for alias in profile.aliases_for(key)} for key in STEP_COLUMN_KEYS}


def required_column_variants(profile: ProfileDefinition) -> list[tuple[str, set[str]]]:
    """Liefert je Pflichtspalte die zulässigen normalisierten Bezeichnungen (Spaltenname plus Aliasgruppe)."""
    aliases = column_aliases(profile)
    variants: list[tuple[str, set[str]]] = []
    for column in profile.required_columns:
        normalized = normalize_heading(column)
        accepted = {normalized}
        for group in aliases.values():
            if normalized in group:
                accepted |= group
        variants.append((column, accepted))
    return variants


def find_step_tables(document: SourceDocument, profile: ProfileDefinition) -> list[StepTableLayout]:
    """Findet alle Tabellen, deren Kopfzeile alle Pflichtspalten des Profils abdeckt."""
    variants = required_column_variants(profile)
    aliases = column_aliases(profile)
    layouts: list[StepTableLayout] = []
    for table in document.tables:
        for row_index, row in enumerate(table.rows[:MAX_HEADER_SCAN_ROWS]):
            header_names = [(cell.grid_column, normalize_heading(cell.text)) for cell in row.cells]
            present = {name for _, name in header_names if name}
            if not all(accepted & present for _, accepted in variants):
                continue
            columns: dict[str, int] = {}
            header_texts: dict[str, str] = {}
            for key in STEP_COLUMN_KEYS:
                for grid_column, name in header_names:
                    if name and name in aliases[key]:
                        columns[key] = grid_column
                        header_texts[key] = name
                        break
            layouts.append(StepTableLayout(table=table, header_row_index=row_index, columns=columns, header_texts=header_texts))
            break
    return layouts


def cell_has_content(cell: SourceCell | None) -> bool:
    return cell is not None and (bool(cell.images) or not is_blank(cell.text))


def classify_row(layout: StepTableLayout, row: SourceRow) -> RowKind:
    """Ordnet eine Zeile unterhalb der Kopfzeile ein (wiederholter Kopf, leer, Layout-Zeile, fachliche Schrittzeile)."""
    if row.is_header or _repeats_header(layout, row):
        return "header"
    content_cells: list[SourceCell] = []
    spans_multiple = False
    for key in CONTENT_KEYS:
        if key not in layout.columns:
            continue
        cell = row.cell_at(layout.columns[key])
        if cell is None or cell.vmerge == "continue":
            continue
        if any(existing is cell for existing in content_cells):
            spans_multiple = True
            continue
        content_cells.append(cell)
    has_content = any(cell_has_content(cell) for cell in content_cells)
    if not has_content:
        return "empty"
    return "layout" if spans_multiple else "data"


def count_data_rows(layout: StepTableLayout) -> int:
    return sum(1 for _, row in layout.data_rows() if classify_row(layout, row) == "data")


def _repeats_header(layout: StepTableLayout, row: SourceRow) -> bool:
    if not layout.header_texts:
        return False
    for key, expected in layout.header_texts.items():
        cell = row.cell_at(layout.columns[key])
        if cell is None or normalize_heading(cell.text) != expected:
            return False
    return True
