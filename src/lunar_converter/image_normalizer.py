"""Bildprüfung und PNG-Normalisierung mit Pillow."""

from __future__ import annotations

import warnings
from io import BytesIO

from PIL import Image, UnidentifiedImageError

SUPPORTED_EXTENSIONS: frozenset[str] = frozenset({".png", ".jpg", ".jpeg"})
_SUPPORTED_FORMATS: frozenset[str] = frozenset({"PNG", "JPEG"})


class ImageNormalizationError(Exception):
    """Bilddaten sind nicht lesbar oder kein unterstütztes Format."""


def to_png(data: bytes) -> bytes:
    """Prüft PNG/JPEG-Bilddaten und liefert PNG-Bytes.

    Gültige PNG-Dateien werden unverändert übernommen, JPEG-Dateien werden
    nach PNG konvertiert.
    """
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("error", Image.DecompressionBombWarning)
            with Image.open(BytesIO(data)) as probe:
                image_format = probe.format
                probe.verify()
            if image_format not in _SUPPORTED_FORMATS:
                raise ImageNormalizationError(f"Bildinhalt hat das nicht unterstützte Format '{image_format}'.")
            if image_format == "PNG":
                return data
            with Image.open(BytesIO(data)) as image:
                image.load()
                converted = image.convert("RGB") if image.mode not in {"RGB", "L"} else image
                output = BytesIO()
                converted.save(output, format="PNG")
                return output.getvalue()
    except ImageNormalizationError:
        raise
    except (UnidentifiedImageError, OSError, SyntaxError, ValueError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise ImageNormalizationError(f"Bilddaten sind nicht lesbar: {type(exc).__name__}: {exc}") from exc
