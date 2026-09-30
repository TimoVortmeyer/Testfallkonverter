"""Validierung des Ziel-JSON vor dem finalen Export.

Reihenfolge: fachliche Pflichtfelder -> JSON-Schema -> Bildreferenzen.
Jede Stufe sammelt alle Befunde, bevor sie einen Fehler auslöst.
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator
from jsonschema.exceptions import SchemaError

from .exceptions import ConfigurationError, ValidationFailedError
from .models import Issue

_STEP_REQUIRED_FIELDS: tuple[tuple[str, str], ...] = (
    ("system", "System"),
    ("action", "Beschreibung des Testschritts"),
    ("expected_result", "Erwartetes Ergebnis"),
)
_ANCHOR_RE = re.compile(r"!([^!\s/\\]+\.png)!")


def load_schema(schema_path: Path) -> dict[str, Any]:
    try:
        schema = json.loads(schema_path.read_text(encoding="utf-8"))
        Draft202012Validator.check_schema(schema)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError, SchemaError) as exc:
        raise ConfigurationError(f"JSON-Schema '{schema_path}' ist nicht lesbar oder ungültig.", details=str(exc)) from exc
    if not isinstance(schema, dict):
        raise ConfigurationError(f"JSON-Schema '{schema_path}' muss ein JSON-Objekt sein.")
    return schema


def check_required_fields(payload: Mapping[str, Any]) -> list[Issue]:
    issues: list[Issue] = []
    if not _non_empty(payload.get("summary")):
        issues.append(
            Issue(code="required_field_missing", field="summary", message="Testfallname fehlt oder ist leer (Summary).")
        )
    steps = payload.get("steps")
    if not isinstance(steps, list) or not steps:
        issues.append(Issue(code="required_field_missing", field="steps", message="Es wurde kein Testschritt erkannt."))
        return issues
    for index, step in enumerate(steps):
        if not isinstance(step, Mapping):
            issues.append(Issue(code="required_field_missing", field=f"steps[{index}]", message=f"Testschritt {index + 1} ist ungültig."))
            continue
        for key, label in _STEP_REQUIRED_FIELDS:
            if not _non_empty(step.get(key)):
                issues.append(
                    Issue(
                        code="required_field_missing",
                        field=f"steps[{index}].{key}",
                        message=f"{label} im {index + 1}. Testschritt ist leer.",
                    )
                )
    return issues


def check_schema(payload: Mapping[str, Any], schema: Mapping[str, Any]) -> list[Issue]:
    validator = Draft202012Validator(schema)
    issues: list[Issue] = []
    for error in sorted(validator.iter_errors(payload), key=lambda err: [str(p) for p in err.absolute_path]):
        issues.append(
            Issue(
                code="schema_validation_failed",
                field=_json_path(list(error.absolute_path)),
                message="Das erzeugte JSON entspricht nicht dem JSON-Schema.",
                details=error.message,
            )
        )
    return issues


def check_image_references(payload: Mapping[str, Any], screenshots_dir: Path) -> list[Issue]:
    """Prüft, dass jede referenzierte Datei existiert und jeder Anker auf eine referenzierte Datei zeigt."""
    issues: list[Issue] = []
    global_refs = [str(name) for name in payload.get("screenshots", [])]
    for index, name in enumerate(global_refs):
        issues.extend(_check_file(name, screenshots_dir, f"screenshots[{index}]"))
    issues.extend(_check_anchors(str(payload.get("description", "")), set(global_refs), "description"))

    for step_index, step in enumerate(payload.get("steps", [])):
        step_refs: list[str] = []
        for list_key in ("attachments", "screenshots"):
            for index, name in enumerate(step.get(list_key, [])):
                step_refs.append(str(name))
                issues.extend(_check_file(str(name), screenshots_dir, f"steps[{step_index}].{list_key}[{index}]"))
        for text_key in ("action", "expected_result"):
            issues.extend(_check_anchors(str(step.get(text_key, "")), set(step_refs), f"steps[{step_index}].{text_key}"))
    return issues


def validate_payload(payload: Mapping[str, Any], schema: Mapping[str, Any], screenshots_dir: Path) -> None:
    for check in (
        lambda: check_required_fields(payload),
        lambda: check_schema(payload, schema),
        lambda: check_image_references(payload, screenshots_dir),
    ):
        issues = check()
        if issues:
            raise ValidationFailedError(issues)


def _check_file(name: str, screenshots_dir: Path, field: str) -> list[Issue]:
    candidate = screenshots_dir / name
    inside = Path(name).name == name and name not in {"", ".", ".."}
    if inside and candidate.is_file():
        return []
    return [
        Issue(
            code="missing_image_reference",
            field=field,
            message=f"Referenzierte Bilddatei '{name}' fehlt im Ordner 'screenshots/'.",
        )
    ]


def _check_anchors(text: str, allowed: set[str], field: str) -> list[Issue]:
    return [
        Issue(
            code="missing_image_reference",
            field=field,
            message=f"Bildanker '!{name}!' verweist auf keine referenzierte Bilddatei dieses Feldes.",
        )
        for name in _ANCHOR_RE.findall(text)
        if name not in allowed
    ]


def _non_empty(value: Any) -> bool:
    return isinstance(value, str) and bool(value.strip())


def _json_path(parts: list[Any]) -> str:
    path = ""
    for part in parts:
        path += f"[{part}]" if isinstance(part, int) else (f".{part}" if path else str(part))
    return path or "<root>"


def validate_references(payload: dict, screenshots_dir: str | Path) -> None:
    screenshots_dir = Path(screenshots_dir)
    referenced = set()
    for image_name in payload.get("screenshots", []):
        referenced.add(image_name)
    for step in payload.get("steps", []):
        for attachment in step.get("attachments", []):
            referenced.add(attachment)
    missing = sorted(name for name in referenced if not (screenshots_dir / name).exists())
    if missing:
        raise ValidationError(f"Referenzierte Bilddateien fehlen im Screenshot-Ordner: {', '.join(missing)}")
