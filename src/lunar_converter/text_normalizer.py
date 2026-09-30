"""Zentrale, wiederverwendbare Textnormalisierung.

Es gibt zwei klar getrennte Anwendungsfälle:

* ``normalize_for_matching``/``normalize_heading``: aggressive Normalisierung
  ausschließlich für Vergleiche (Marker, Spaltenköpfe, Feldbezeichner).
* ``clean_line``/``clean_multiline``: behutsame Bereinigung von Word-Layout-
  Artefakten für Inhalte, die exportiert werden. Inhalte werden dabei nicht
  fachlich verändert.
"""

from __future__ import annotations

import re
import unicodedata

_DASH_TRANSLATION = str.maketrans({char: "-" for char in "\u2010\u2011\u2012\u2013\u2014\u2015\u2212"})
_INVISIBLE_RE = re.compile("[\u200b\u200c\u200d\u2060\ufeff\u00ad]")
_HORIZONTAL_SPACE_RE = re.compile("[ \t\u00a0\u2007\u202f\u3000]+")
_SLASH_RE = re.compile(r"\s*/\s*")
_DASH_SPACING_RE = re.compile(r"\s*-\s*")
_TRAILING_PUNCTUATION_RE = re.compile(r"[\s.:;,!?]+$")


def normalize_newlines(value: str) -> str:
    """Vereinheitlicht Zeilenumbrüche auf ``\\n``."""
    return value.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")


def normalize_for_matching(value: str | None) -> str:
    """Normalisiert Text für robuste Vergleiche.

    Ignoriert Groß-/Kleinschreibung, Zeilenumbrüche, Tabs, mehrfache
    Leerzeichen, Bindestrichvarianten sowie Leerzeichen um ``/`` und ``-``.
    """
    if not value:
        return ""
    text = unicodedata.normalize("NFC", value)
    text = normalize_newlines(text).replace("\n", " ")
    text = _INVISIBLE_RE.sub("", text)
    text = text.translate(_DASH_TRANSLATION)
    text = _HORIZONTAL_SPACE_RE.sub(" ", text)
    text = _SLASH_RE.sub("/", text)
    text = _DASH_SPACING_RE.sub("-", text)
    return text.casefold().strip()


def normalize_heading(value: str | None) -> str:
    """Wie ``normalize_for_matching``, ignoriert zusätzlich Satzzeichen am Ende."""
    return _TRAILING_PUNCTUATION_RE.sub("", normalize_for_matching(value))


def contains_marker(normalized_text: str, marker: str) -> bool:
    """Prüft, ob ein Marker im bereits normalisierten Dokumenttext vorkommt."""
    normalized_marker = normalize_heading(marker)
    return bool(normalized_marker) and normalized_marker in normalized_text


def clean_line(value: str) -> str:
    """Entfernt Layout-Artefakte einer einzelnen Zeile (unsichtbare Zeichen, Mehrfach-Leerzeichen, Tabs)."""
    text = _INVISIBLE_RE.sub("", unicodedata.normalize("NFC", value))
    return _HORIZONTAL_SPACE_RE.sub(" ", text).strip()


def clean_multiline(value: str) -> str:
    """Bereinigt mehrzeiligen Text zeilenweise und fasst Leerzeilenfolgen zu einer Leerzeile zusammen."""
    lines: list[str] = []
    for raw_line in normalize_newlines(value).split("\n"):
        line = clean_line(raw_line)
        if not line and (not lines or not lines[-1]):
            continue
        lines.append(line)
    while lines and not lines[-1]:
        lines.pop()
    return "\n".join(lines)


def is_blank(value: str | None) -> bool:
    """Liefert ``True``, wenn der Text nach Entfernen von Layout-Artefakten leer ist."""
    return not clean_multiline(value or "")
