"""Optionaler Test mit dem lokal vorhandenen echten Beispieldokument ``input/sample.docx``.

Das Dokument wird nicht in das Repository übernommen; ohne Datei wird der Test übersprungen.
"""

from __future__ import annotations

import pytest

from lunar_converter.docx_reader import read_docx
from lunar_converter.lunar_parser import parse_test_case
from lunar_converter.profile_detector import detect_profile
from lunar_converter.profile_loader import load_profiles
from lunar_converter.semantic_model import ImageMarker
from tests.helpers import CONFIG_DIR, REAL_SAMPLE

pytestmark = pytest.mark.skipif(not REAL_SAMPLE.is_file(), reason="input/sample.docx ist lokal nicht vorhanden")


def test_echtes_beispiel_wird_erkannt_und_extrahiert() -> None:
    document = read_docx(REAL_SAMPLE)
    detection = detect_profile(document, load_profiles(CONFIG_DIR))
    assert detection.status == "matched"
    assert detection.profile is not None

    test_case = parse_test_case(document, detection.profile)

    assert test_case.name.startswith("TFB_")
    assert test_case.process_path and test_case.process_path[0][:2].isdigit()
    assert test_case.info_table is not None
    assert len(test_case.steps) >= 1
    assert all(step.system for step in test_case.steps)
    anchored = [item for step in test_case.steps for line in step.action.lines for item in line if isinstance(item, ImageMarker)]
    assert len(anchored) + len(test_case.unassigned_images) == len(document.images)
