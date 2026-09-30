"""JSON-Schema- und Referenzvalidierung."""

from __future__ import annotations

import copy
from pathlib import Path
from typing import Any

import pytest

from lunar_converter.exceptions import ValidationFailedError
from lunar_converter.validator import check_image_references, check_schema, load_schema, validate_payload
from tests.helpers import SCHEMA_PATH


def _payload() -> dict[str, Any]:
    return {
        "summary": "TF_1",
        "description": "h2. Kurzbeschreibung\nText\n!0001.png!",
        "labels": [],
        "components": [],
        "custom_fields": {},
        "screenshots": ["0001.png"],
        "steps": [
            {
                "system": "SAP",
                "action": "Klick !0002.png!",
                "data": "",
                "expected_result": "OK",
                "tester": "",
                "attachments": ["0002.png"],
                "screenshots": [],
            }
        ],
    }


@pytest.fixture
def screenshots(tmp_path: Path) -> Path:
    for name in ("0001.png", "0002.png"):
        (tmp_path / name).write_bytes(b"png")
    return tmp_path


def test_gueltiges_payload_besteht_schema() -> None:
    assert check_schema(_payload(), load_schema(SCHEMA_PATH)) == []


@pytest.mark.parametrize(
    ("mutate", "expected_field"),
    [
        (lambda p: p.pop("summary"), "<root>"),
        (lambda p: p.update(steps=[]), "steps"),
        (lambda p: p.update(unbekannt=1), "<root>"),
        (lambda p: p["steps"][0].pop("expected_result"), "steps[0]"),
        (lambda p: p["steps"][0].update(system=""), "steps[0].system"),
        (lambda p: p.update(screenshots=["../boese.png"]), "screenshots[0]"),
        (lambda p: p["steps"][0].update(attachments=["bild.gif"]), "steps[0].attachments[0]"),
    ],
)
def test_schemaverletzungen_werden_erkannt(mutate: Any, expected_field: str) -> None:
    payload = copy.deepcopy(_payload())
    mutate(payload)
    issues = check_schema(payload, load_schema(SCHEMA_PATH))
    assert issues
    assert all(issue.code == "schema_validation_failed" for issue in issues)
    assert expected_field in [issue.field for issue in issues]


def test_validate_payload_erfolgreich(screenshots: Path) -> None:
    validate_payload(_payload(), load_schema(SCHEMA_PATH), screenshots)


def test_fehlende_bilddatei_wird_erkannt(screenshots: Path) -> None:
    (screenshots / "0002.png").unlink()
    issues = check_image_references(_payload(), screenshots)
    assert [(issue.code, issue.field) for issue in issues] == [("missing_image_reference", "steps[0].attachments[0]")]


def test_anker_ohne_referenz_wird_erkannt(screenshots: Path) -> None:
    payload = _payload()
    payload["steps"][0]["attachments"] = []
    issues = check_image_references(payload, screenshots)
    assert [issue.field for issue in issues] == ["steps[0].action"]


def test_pflichtfelder_vor_schema(screenshots: Path) -> None:
    payload = _payload()
    payload["summary"] = ""
    with pytest.raises(ValidationFailedError) as info:
        validate_payload(payload, load_schema(SCHEMA_PATH), screenshots)
    assert info.value.code == "required_field_missing"
    assert info.value.field == "summary"
