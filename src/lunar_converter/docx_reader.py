"""Technische DOCX-Extraktion.

``python-docx`` öffnet das Paket, löst Relationships und Formatvorlagen auf und
liefert die Mediendateien. Die eigentliche Reihenfolge von Text und Bildern
ermittelt ``ooxml_traversal`` direkt auf XML-Ebene.

Ausgewertet wird nur der Dokumentkörper. Bilder in Kopf- und Fußzeilen werden
nie extrahiert und erhalten nie einen Anker; sie werden nur gezählt.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import docx
from docx.document import Document as DocxDocument
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .exceptions import InputFileError
from .models import SourceDocument
from .ooxml_traversal import BodyTraversal, EmbeddedWordDocumentDetected, NestedTableDetected, ResolvedImage, StyleClassifier


def read_docx(path: Path) -> SourceDocument:
    """Liest eine DOCX-Datei in das technische Quellmodell."""
    document = _open_document(path)
    traversal = BodyTraversal(
        resolve_image=lambda rel_id, linked: _resolve_image(document, rel_id, linked),
        is_heading_style=_heading_style_classifier(document),
        numbering=_numbering_element(document),
        styles=document.styles.element,
        is_embedded_word=lambda relationship_id, prog_id: _is_embedded_word(document, relationship_id, prog_id),
    )
    try:
        blocks = traversal.traverse(document.element.body)
    except EmbeddedWordDocumentDetected as exc:
        raise InputFileError(
            f"Eingebettetes Word-Dokument kann nicht verarbeitet werden: {exc.location}.",
            field=exc.location,
            details=f"OLE ProgID: {exc.prog_id}; Verarbeitungsphase: DOCX-Objekterkennung.",
            code="embedded_word_document_not_processed",
        ) from exc
    except NestedTableDetected as exc:
        location = "; ".join(exc.locations)
        raise InputFileError(
            f"Verschachtelte Word-Tabelle wird nicht unterstützt: {location}.",
            field=location,
            details=f"{len(exc.locations)} innere Tabelle(n); Verarbeitungsphase: DOCX-Strukturerkennung.",
            code="nested_table_not_supported",
        ) from exc
    except Exception as exc:
        raise InputFileError(
            "Der Dokumentinhalt konnte nicht ausgewertet werden (Phase: DOCX-Auswertung, Dokumentkörper).",
            details=f"{type(exc).__name__}: {exc}",
            code="docx_content_error",
        ) from exc
    return SourceDocument(
        source_path=path,
        blocks=blocks,
        images=traversal.images,
        ignored_header_footer_images=_count_header_footer_images(document),
        warnings=traversal.warnings,
    )


def _numbering_element(document: DocxDocument) -> Any | None:
    # python-docx kann einen fehlenden Nummerierungsteil nicht anlegen (NotImplementedError).
    try:
        return document.part.part_related_by(RT.NUMBERING).element
    except KeyError:
        return None


def _is_embedded_word(document: DocxDocument, relationship_id: str | None, prog_id: str | None) -> bool:
    if prog_id and prog_id.casefold().startswith("word.document"):
        return True
    if not relationship_id:
        return False
    relationship = document.part.rels.get(relationship_id)
    if relationship is None or relationship.is_external:
        return False
    target = relationship.target_part
    content_type = getattr(target, "content_type", "").casefold()
    part_name = str(getattr(target, "partname", "")).casefold()
    return (
        "wordprocessingml.document" in content_type
        or content_type == "application/msword"
        or part_name.endswith((".doc", ".docx"))
    )


def _open_document(path: Path) -> DocxDocument:
    if not path.is_file():
        raise InputFileError(f"Eingabedatei '{path}' wurde nicht gefunden.")
    if path.suffix.lower() != ".docx":
        raise InputFileError(f"Nur .docx-Dateien können direkt gelesen werden: '{path.name}'.")
    try:
        return docx.Document(str(path))
    except Exception as exc:
        # python-docx meldet defekte Pakete über unterschiedliche Ausnahmetypen.
        raise InputFileError(
            "Die DOCX-Datei ist beschädigt oder kein gültiges Word-Dokument.",
            details=f"{type(exc).__name__}: {exc}",
        ) from exc


def _count_header_footer_images(document: DocxDocument) -> int:
    """Zählt Bildplatzierungen in Kopf- und Fußzeilen, ohne sie aufzulösen oder zu exportieren."""
    count = 0
    for relationship in document.part.rels.values():
        if relationship.is_external or relationship.reltype not in (RT.HEADER, RT.FOOTER):
            continue
        traversal = BodyTraversal(
            resolve_image=lambda rel_id, linked: ResolvedImage(part_name=None, blob=None),
            is_heading_style=lambda style_id: False,
            reject_nested_tables=False,
        )
        traversal.traverse(relationship.target_part.element)
        count += len(traversal.images)
    return count


def _resolve_image(document: DocxDocument, relationship_id: str, linked: bool) -> ResolvedImage:
    relationship: Any = document.part.rels.get(relationship_id)
    if relationship is None:
        return ResolvedImage(part_name=None, blob=None)
    if linked or relationship.is_external:
        return ResolvedImage(part_name=str(relationship.target_ref), blob=None, external=True)
    target = relationship.target_part
    return ResolvedImage(part_name=str(target.partname), blob=target.blob)


def _heading_style_classifier(document: DocxDocument) -> StyleClassifier:
    heading_ids: set[str] = set()
    for style in document.styles:
        style_id = getattr(style, "style_id", None)
        name = (getattr(style, "name", None) or "").casefold()
        if style_id and (name.startswith("heading") or name in {"title", "subtitle"}):
            heading_ids.add(style_id)

    def is_heading(style_id: str | None) -> bool:
        if not style_id:
            return False
        lowered = style_id.casefold()
        return style_id in heading_ids or lowered.startswith(("heading", "berschrift", "überschrift"))

    return is_heading
