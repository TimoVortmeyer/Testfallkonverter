"""Ziel-Renderer: semantisches Testfallmodell -> ``testcase.json`` des Jira/Xray-Importers.

Dies ist die einzige Stelle, an der das Ziel-JSON zusammengesetzt wird.
Anpassungen am Zielformat (z. B. Custom-Field-Mappings) erfolgen hier,
ohne die Extraktion zu verändern.
"""

from __future__ import annotations

import json
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from .image_assignment import ImageAssignmentResult
from .semantic_model import CheckboxMarker, InfoTable, RichText, TestCase, TestStep
from .text_normalizer import clean_multiline

ACTUAL_RESULT_HEADING = "h3. Tatsächliches Ergebnis"
EMPTY_EXPECTED_RESULT = "-"
UNDEFINED_SYSTEM = "nicht definiert"
PROCESS_PATH_FIELD = "customfield_15909"
PROCESS_PATH_SEPARATOR = "/"
_JIRA_EMOTICON_RE = re.compile(r"\((?:/?|[ynix!+*?\-]|on|off|flag|flagoff|warning|thumbsup|thumbsdown|heart|star)\)", re.IGNORECASE)
_CHECKED_CHECKBOX_TOKEN = "\x00LUNAR_CHECKED_CHECKBOX\x00"
_UNCHECKED_CHECKBOX_TOKEN = "\x00LUNAR_UNCHECKED_CHECKBOX\x00"


def escape_jira_emoticons(text: str) -> str:
    """Verhindert, dass Jira bekannte Emoticon-Kürzel in Symbole umwandelt."""
    return _JIRA_EMOTICON_RE.sub(lambda match: "\\" + match.group(0), text)


def wiki_anchor(file_name: str) -> str:
    return f"!{file_name}!"


def render_rich_text(rich: RichText, anchors: Mapping[int, str], *, render_checkboxes: bool = False) -> str:
    """Setzt Text und Jira-Wiki-Bildanker zusammen. Bilder ohne Anker werden ausgelassen."""
    lines: list[str] = []
    for line in rich.lines:
        buffer = ""
        needs_gap = False
        for item in line:
            if isinstance(item, str):
                # Nach Anker oder ausgelassenem Bild Wörter nicht zusammenkleben.
                if needs_gap and item[:1].isalnum() and buffer and not buffer[-1].isspace():
                    buffer += " "
                buffer += item
                needs_gap = False
                continue
            if isinstance(item, CheckboxMarker):
                if render_checkboxes:
                    buffer += _CHECKED_CHECKBOX_TOKEN if item.checked else _UNCHECKED_CHECKBOX_TOKEN
                else:
                    buffer += "☒" if item.checked else "☐"
                needs_gap = False
                continue
            needs_gap = True
            file_name = anchors.get(item.image_id)
            if file_name is None:
                continue
            if buffer and not buffer[-1].isspace():
                buffer += " "
            buffer += wiki_anchor(file_name)
        if not buffer.strip() and any(not isinstance(item, str) for item in line):
            # Absatz bestand nur aus nicht exportierten Bildern: keine Leerzeile erzeugen.
            continue
        buffer = escape_jira_emoticons(buffer)
        buffer = buffer.replace(_CHECKED_CHECKBOX_TOKEN, "(/)").replace(_UNCHECKED_CHECKBOX_TOKEN, "(x)")
        lines.append(buffer)
    return clean_multiline("\n".join(lines))


def render_wiki_table(table: InfoTable, anchors: Mapping[int, str]) -> str:
    """Jira-Wiki-Tabelle; Zeilenumbrüche in Zellen werden zu ``\\\\``, Pipes werden maskiert."""
    lines: list[str] = []
    for row in table.rows:
        cells = [_wiki_cell(render_rich_text(cell, anchors, render_checkboxes=True)) for cell in row]
        lines.append("|" + "|".join(cells) + "|")
    return "\n".join(lines)


def _wiki_cell(text: str) -> str:
    if not text:
        return " "
    return " \\\\ ".join(line.replace("|", "\\|") for line in text.split("\n") if line)


class XrayImportRenderer:
    """Erzeugt das Import-JSON. Methoden können für abweichende Zielformate überschrieben werden."""

    def render(self, test_case: TestCase, images: ImageAssignmentResult, source_name: str) -> dict[str, Any]:
        anchors = images.anchors
        return {
            "summary": self.render_summary(source_name),
            "description": self.render_description(test_case, anchors),
            "labels": self.render_labels(test_case),
            "components": self.render_components(test_case),
            "custom_fields": self.render_custom_fields(test_case),
            "screenshots": images.global_screenshots,
            "steps": [self.render_step(step, anchors, images.step_attachments(step.index)) for step in test_case.steps],
        }

    def render_summary(self, source_name: str) -> str:
        return escape_jira_emoticons(source_name)

    def render_description(self, test_case: TestCase, anchors: Mapping[int, str]) -> str:
        sections: list[str] = []
        if test_case.title:
            sections.append(f"h1. {escape_jira_emoticons(test_case.title)}")
        if test_case.info_table is not None:
            sections.append(render_wiki_table(test_case.info_table, anchors))
        return "\n\n".join(sections)

    def render_step(self, step: TestStep, anchors: Mapping[int, str], attachments: list[str]) -> dict[str, Any]:
        action = render_rich_text(step.action, anchors)
        data = clean_multiline(step.data)
        if data:
            data = escape_jira_emoticons(data)
            action = f"{action}\n{data}" if action else data
        return {
            "system": self.render_system(step),
            "action": action,
            "data": "",
            "expected_result": self.render_expected_result(step, anchors),
            "tester": "",
            "attachments": list(attachments),
            "screenshots": [],
        }

    def render_system(self, step: TestStep) -> str:
        # Fehlende Systemspalte oder leere Systemzelle.
        return escape_jira_emoticons(clean_multiline(step.system)) or UNDEFINED_SYSTEM

    def render_expected_result(self, step: TestStep, anchors: Mapping[int, str]) -> str:
        # Übergangslösung: tatsächliches Ergebnis als eigener Block am erwarteten Ergebnis.
        expected = render_rich_text(step.expected_result, anchors) or EMPTY_EXPECTED_RESULT
        actual = render_rich_text(step.actual_result, anchors)
        if actual:
            return f"{expected}\n\n{ACTUAL_RESULT_HEADING}\n{actual}"
        return expected

    def render_labels(self, test_case: TestCase) -> list[str]:
        return list(test_case.labels)

    def render_components(self, test_case: TestCase) -> list[str]:
        return []

    def render_custom_fields(self, test_case: TestCase) -> dict[str, Any]:
        if not test_case.process_path:
            return {}
        value = escape_jira_emoticons(PROCESS_PATH_SEPARATOR.join(test_case.process_path))
        return {PROCESS_PATH_FIELD: value}


def write_testcase_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
