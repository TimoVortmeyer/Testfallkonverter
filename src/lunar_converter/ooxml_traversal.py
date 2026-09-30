"""OOXML-nahe Traversierung des Dokumentkörpers.

``python-docx`` liefert Absätze und Tabellen, aber nicht die exakte Reihenfolge
von Text und Bildankern innerhalb eines Absatzes. Dieses Modul durchläuft daher
die XML-Elemente von ``word/document.xml`` direkt und erzeugt das technische
Quellmodell inklusive Bildplatzierungen (DrawingML ``a:blip`` und VML
``v:imagedata``) in natürlicher Dokumentreihenfolge: oben nach unten, in
Tabellen zeilenweise und innerhalb einer Zeile von links nach rechts.

Kopf- und Fußzeilen werden bewusst nicht durchlaufen.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from typing import Any

from .models import ImageKind, ImageRef, Segment, SourceBlock, SourceCell, SourceParagraph, SourceRow, SourceTable, VMerge

NAMESPACES: dict[str, str] = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "v": "urn:schemas-microsoft-com:vml",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}


def qn(prefixed: str) -> str:
    prefix, local = prefixed.split(":", 1)
    return f"{{{NAMESPACES[prefix]}}}{local}"


W_P = qn("w:p")
W_TBL = qn("w:tbl")
W_TR = qn("w:tr")
W_TC = qn("w:tc")
W_SDT = qn("w:sdt")
W_SDT_CONTENT = qn("w:sdtContent")
W_CUSTOM_XML = qn("w:customXml")
W_T = qn("w:t")
W_TAB = qn("w:tab")
W_PTAB = qn("w:ptab")
W_BR = qn("w:br")
W_CR = qn("w:cr")
W_NO_BREAK_HYPHEN = qn("w:noBreakHyphen")
W_DRAWING = qn("w:drawing")
W_PICT = qn("w:pict")
W_OBJECT = qn("w:object")
W_TXBX_CONTENT = qn("w:txbxContent")
W_VAL = qn("w:val")
A_BLIP = qn("a:blip")
V_IMAGEDATA = qn("v:imagedata")
WP_ANCHOR = qn("wp:anchor")
MC_ALTERNATE_CONTENT = qn("mc:AlternateContent")
MC_CHOICE = qn("mc:Choice")
MC_FALLBACK = qn("mc:Fallback")
R_EMBED = qn("r:embed")
R_LINK = qn("r:link")
R_ID = qn("r:id")

# Elemente ohne sichtbaren Inhalt oder mit gelöschtem/verstecktem Inhalt.
_SKIPPED_INLINE_TAGS = frozenset(
    qn(tag)
    for tag in (
        "w:pPr",
        "w:rPr",
        "w:sectPr",
        "w:del",
        "w:delText",
        "w:moveFrom",
        "w:instrText",
        "w:delInstrText",
        "w:sdtPr",
        "w:sdtEndPr",
        "w:fldChar",
        "w:softHyphen",
        "w:commentReference",
        "w:footnoteReference",
        "w:endnoteReference",
        "mc:Fallback",
    )
)
_VML_ABSOLUTE_RE = re.compile(r"position\s*:\s*absolute", re.IGNORECASE)


@dataclass(frozen=True)
class ResolvedImage:
    """Ergebnis der Auflösung einer Relationship-ID auf eine Mediendatei."""

    part_name: str | None
    blob: bytes | None
    external: bool = False


ImageResolver = Callable[[str, bool], ResolvedImage]
StyleClassifier = Callable[[str | None], bool]


class BodyTraversal:
    """Erzeugt das technische Quellmodell aus dem ``w:body``-Element."""

    def __init__(self, resolve_image: ImageResolver, is_heading_style: StyleClassifier) -> None:
        self._resolve_image = resolve_image
        self._is_heading_style = is_heading_style
        self._images: list[ImageRef] = []
        self._table_count = 0
        self._paragraph_count = 0

    @property
    def images(self) -> list[ImageRef]:
        return list(self._images)

    def traverse(self, body: Any) -> list[SourceBlock]:
        blocks: list[SourceBlock] = []
        for element in _iter_content(body, (W_P, W_TBL)):
            if element.tag == W_P:
                self._paragraph_count += 1
                blocks.append(self._paragraph(element, f"Absatz {self._paragraph_count}"))
            else:
                self._table_count += 1
                blocks.append(self._table(element, self._table_count))
        return blocks

    def _table(self, tbl: Any, table_index: int) -> SourceTable:
        rows: list[SourceRow] = []
        for row_index, tr in enumerate(_iter_content(tbl, (W_TR,)), start=1):
            rows.append(self._row(tr, f"Tabelle {table_index}, Zeile {row_index}"))
        return SourceTable(table_index=table_index, rows=rows)

    def _row(self, tr: Any, location: str) -> SourceRow:
        grid_column = _int_attribute(tr.find(f"{qn('w:trPr')}/{qn('w:gridBefore')}"), default=0)
        is_header = tr.find(f"{qn('w:trPr')}/{qn('w:tblHeader')}") is not None
        cells: list[SourceCell] = []
        for cell_index, tc in enumerate(_iter_content(tr, (W_TC,)), start=1):
            span = max(1, _int_attribute(tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}"), default=1))
            cell_location = f"{location}, Zelle {cell_index}"
            cells.append(
                SourceCell(
                    paragraphs=self._cell_paragraphs(tc, cell_location),
                    grid_column=grid_column,
                    grid_span=span,
                    vmerge=_vmerge(tc),
                )
            )
            grid_column += span
        return SourceRow(cells=cells, is_header=is_header)

    def _cell_paragraphs(self, container: Any, location: str) -> list[SourceParagraph]:
        # Verschachtelte Tabellen werden in Lesereihenfolge in die Zelle eingeflacht.
        paragraphs: list[SourceParagraph] = []
        for element in _iter_content(container, (W_P, W_TBL)):
            if element.tag == W_P:
                paragraphs.append(self._paragraph(element, location))
            else:
                for tr in _iter_content(element, (W_TR,)):
                    for tc in _iter_content(tr, (W_TC,)):
                        paragraphs.extend(self._cell_paragraphs(tc, location))
        return paragraphs

    def _paragraph(self, p: Any, location: str) -> SourceParagraph:
        style_element = p.find(f"{qn('w:pPr')}/{qn('w:pStyle')}")
        style_id = style_element.get(W_VAL) if style_element is not None else None
        has_outline = p.find(f"{qn('w:pPr')}/{qn('w:outlineLvl')}") is not None
        segments: list[Segment] = []
        self._walk_inline(p, segments, location, floating=False)
        return SourceParagraph(
            segments=_merge_text_segments(segments),
            style_id=style_id,
            is_heading=has_outline or self._is_heading_style(style_id),
        )

    def _walk_inline(self, element: Any, segments: list[Segment], location: str, *, floating: bool) -> None:
        for child in element:
            tag = child.tag
            if not isinstance(tag, str) or tag in _SKIPPED_INLINE_TAGS:
                continue
            if tag == W_T:
                if child.text:
                    segments.append(child.text)
            elif tag in (W_TAB, W_PTAB):
                segments.append("\t")
            elif tag in (W_BR, W_CR):
                segments.append("\n")
            elif tag == W_NO_BREAK_HYPHEN:
                segments.append("-")
            elif tag == MC_ALTERNATE_CONTENT:
                self._walk_alternate_content(child, segments, location, floating=floating)
            elif tag in (W_DRAWING, W_PICT, W_OBJECT):
                self._walk_graphic(child, segments, location, floating=floating)
            else:
                self._walk_inline(child, segments, location, floating=floating)

    def _walk_alternate_content(self, element: Any, segments: list[Segment], location: str, *, floating: bool) -> None:
        # Nur die erste Variante auswerten, sonst würden Bilder aus mc:Fallback doppelt gezählt.
        choice = element.find(MC_CHOICE)
        chosen = choice if choice is not None else element.find(MC_FALLBACK)
        if chosen is not None:
            self._walk_inline(chosen, segments, location, floating=floating)

    def _walk_graphic(self, element: Any, segments: list[Segment], location: str, *, floating: bool) -> None:
        for child in element:
            tag = child.tag
            if not isinstance(tag, str) or tag == MC_FALLBACK:
                continue
            child_floating = floating or tag == WP_ANCHOR or _is_absolute_vml(child)
            if tag == A_BLIP:
                self._add_image(child.get(R_EMBED) or child.get(R_LINK), "drawingml", child.get(R_EMBED) is None, segments, location, child_floating)
            elif tag == V_IMAGEDATA:
                self._add_image(child.get(R_ID), "vml", False, segments, location, child_floating)
            elif tag == W_TXBX_CONTENT:
                for paragraph in _iter_content(child, (W_P,)):
                    segments.append("\n")
                    self._walk_inline(paragraph, segments, location, floating=child_floating)
            elif tag == MC_ALTERNATE_CONTENT:
                choice = child.find(MC_CHOICE)
                chosen = choice if choice is not None else child.find(MC_FALLBACK)
                if chosen is not None:
                    self._walk_graphic(chosen, segments, location, floating=child_floating)
            else:
                self._walk_graphic(child, segments, location, floating=child_floating)

    def _add_image(
        self,
        relationship_id: str | None,
        kind: ImageKind,
        linked: bool,
        segments: list[Segment],
        location: str,
        floating: bool,
    ) -> None:
        if not relationship_id:
            return
        resolved = self._resolve_image(relationship_id, linked)
        image = ImageRef(
            image_id=len(self._images) + 1,
            relationship_id=relationship_id,
            kind=kind,
            floating=floating,
            location=location,
            part_name=resolved.part_name,
            blob=resolved.blob,
            external=resolved.external,
        )
        self._images.append(image)
        segments.append(image)


def _iter_content(element: Any, wanted: tuple[str, ...]) -> Iterator[Any]:
    """Liefert gewünschte Kindelemente und löst Inhaltssteuerelemente (``w:sdt``, ``w:customXml``) transparent auf."""
    for child in element:
        tag = child.tag
        if tag in wanted:
            yield child
        elif tag == W_SDT:
            content = child.find(W_SDT_CONTENT)
            if content is not None:
                yield from _iter_content(content, wanted)
        elif tag == W_CUSTOM_XML:
            yield from _iter_content(child, wanted)


def _merge_text_segments(segments: list[Segment]) -> list[Segment]:
    merged: list[Segment] = []
    for segment in segments:
        if isinstance(segment, str) and merged and isinstance(merged[-1], str):
            merged[-1] = merged[-1] + segment
        else:
            merged.append(segment)
    return merged


def _int_attribute(element: Any, *, default: int) -> int:
    if element is None:
        return default
    try:
        return int(element.get(W_VAL, default))
    except (TypeError, ValueError):
        return default


def _vmerge(tc: Any) -> VMerge | None:
    element = tc.find(f"{qn('w:tcPr')}/{qn('w:vMerge')}")
    if element is None:
        return None
    return "restart" if element.get(W_VAL) == "restart" else "continue"


def _is_absolute_vml(element: Any) -> bool:
    tag = element.tag
    if not isinstance(tag, str) or not tag.startswith(f"{{{NAMESPACES['v']}}}"):
        return False
    return bool(_VML_ABSOLUTE_RE.search(element.get("style") or ""))
