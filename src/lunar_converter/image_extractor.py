"""Export der sichtbaren Bildplatzierungen als nummerierte PNG-Dateien.

Nummerierung je Dokument ab ``0001.png`` in natürlicher Dokumentreihenfolge.
Nicht unterstützte, externe oder unlesbare Bilder werden übersprungen und als
Warnung gemeldet; sie erhalten keine Nummer und keinen Anker.
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import PurePosixPath, Path

from .image_normalizer import SUPPORTED_EXTENSIONS, ImageNormalizationError, to_png
from .models import ImageRef, Issue


@dataclass
class ImageExportResult:
    exported: dict[int, str] = field(default_factory=dict)
    warnings: list[Issue] = field(default_factory=list)


def image_file_name(number: int) -> str:
    return f"{number:04d}.png"


def export_images(images: Sequence[ImageRef], target_dir: Path) -> ImageExportResult:
    """Schreibt alle unterstützten Bilder nach ``target_dir`` und liefert die Zuordnung ``image_id -> Dateiname``."""
    target_dir.mkdir(parents=True, exist_ok=True)
    result = ImageExportResult()
    for image in images:
        issue = _check_exportable(image)
        if issue is not None:
            result.warnings.append(issue)
            continue
        assert image.blob is not None
        try:
            png_bytes = to_png(image.blob)
        except ImageNormalizationError as exc:
            result.warnings.append(
                Issue(
                    code="image_unreadable",
                    message=f"Bild {image.image_id} ({image.location}) ist nicht lesbar und wird nicht exportiert.",
                    details=f"{image.part_name}: {exc}",
                )
            )
            continue
        file_name = image_file_name(len(result.exported) + 1)
        (target_dir / file_name).write_bytes(png_bytes)
        result.exported[image.image_id] = file_name
    return result


def _check_exportable(image: ImageRef) -> Issue | None:
    if image.external:
        return Issue(
            code="external_image_not_supported",
            message=f"Bild {image.image_id} ({image.location}) ist nur verknüpft, nicht eingebettet, und wird nicht exportiert.",
            details=image.part_name,
        )
    if image.blob is None or image.part_name is None:
        return Issue(
            code="image_missing",
            message=f"Bild {image.image_id} ({image.location}) verweist auf eine fehlende Mediendatei.",
            details=f"Relationship-ID: {image.relationship_id}",
        )
    extension = PurePosixPath(image.part_name).suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        return Issue(
            code="unsupported_image_format",
            message=(
                f"Bild {image.image_id} ({image.location}) hat das nicht unterstützte Format '{extension or '?'}' "
                "und wird nicht exportiert."
            ),
            details=image.part_name,
        )
    return None
