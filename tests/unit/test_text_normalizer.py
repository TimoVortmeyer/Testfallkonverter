from __future__ import annotations

import pytest

from lunar_converter.text_normalizer import clean_multiline, contains_marker, normalize_for_matching, normalize_heading


@pytest.mark.parametrize(
    ("left", "right"),
    [
        ("System / Komponente", "System/Komponente"),
        ("Schritt\u2013Nr.", "schritt-nr"),
        ("  Kurzbeschreibung:  ", "KURZBESCHREIBUNG"),
        ("Datum\tder  letzten\nÄnderung", "datum der letzten änderung"),
        ("Geschäftsprozess \u2011 Szenario", "Geschäftsprozess-Szenario"),
    ],
)
def test_normalize_heading_gleicht_varianten_an(left: str, right: str) -> None:
    assert normalize_heading(left) == normalize_heading(right)


def test_marker_suche_im_normalisierten_text() -> None:
    text = normalize_for_matching("Kapitel 3\nTESTABLAUF")
    assert contains_marker(text, "Testablauf")
    assert not contains_marker(text, "Kurzbeschreibung")
    assert not contains_marker(text, "   ")


def test_clean_multiline_entfernt_nur_layoutartefakte() -> None:
    raw = "\n  Zeile\u00a0 1\t mit Tab \n\n\n\u200bZeile 2  \n\n"
    assert clean_multiline(raw) == "Zeile 1 mit Tab\n\nZeile 2"
