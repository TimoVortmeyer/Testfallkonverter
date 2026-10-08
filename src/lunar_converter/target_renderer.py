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
from .semantic_model import CheckboxMarker, InfoTable, RichText, StyledText, TestCase, TestStep
from .text_normalizer import clean_multiline

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
            if isinstance(item, StyledText):
                raw_text = item.text
                leading = raw_text[: len(raw_text) - len(raw_text.lstrip())]
                trailing = raw_text[len(raw_text.rstrip()):]
                core = raw_text.strip()
                text = escape_jira_emoticons(core)
                if item.underline:
                    text = f"+{text}+"
                if item.italic:
                    text = f"_{text}_"
                if item.bold:
                    text = f"*{text}*"
                buffer += leading + text + trailing
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
    lines = [line.replace("|", "\\|") for line in text.split("\n") if line]
    if any(line.lstrip().startswith(("*", "#")) for line in lines):
        return "\n".join(lines)
    if len(lines) > 1 and all(re.match(r"^\s*[-–•▪]\s+\S", line) for line in lines):
        return "\n".join(lines)
    return " \\\\ ".join(lines)


class XrayImportRenderer:
    """Erzeugt das Import-JSON. Methoden können für abweichende Zielformate überschrieben werden."""

    def render(self, test_case: TestCase, images: ImageAssignmentResult, source_name: str) -> dict[str, Any]:
        anchors = images.anchors
        payload = {
            "summary": self.render_summary(test_case.name),
            "description": self.render_description(test_case, anchors),
            "labels": self.render_labels(test_case),
            "components": self.render_components(test_case),
            "custom_fields": self.render_custom_fields(test_case),
            "screenshots": images.global_screenshots,
            "source_word_filename": test_case.source_word_filename or test_case.source_file.name,
        }
        if test_case.steps:
            payload["steps"] = [self.render_step(step, anchors, images.step_attachments(step.index)) for step in test_case.steps]
        if test_case.reporter_email:
            payload["reporter_email"] = test_case.reporter_email
        return payload

    def render_summary(self, testcase_name: str) -> str:
        return escape_jira_emoticons(clean_multiline(testcase_name))

    def render_description(self, test_case: TestCase, anchors: Mapping[int, str]) -> str:
        sections: list[str] = []
        if test_case.title:
            sections.append(f"h1. {escape_jira_emoticons(test_case.title)}")
        if test_case.info_table is not None:
            sections.append(render_wiki_table(test_case.info_table, anchors))
        return "\n\n".join(sections)

    def render_step(self, step: TestStep, anchors: Mapping[int, str], attachments: list[str]) -> dict[str, Any]:
        action = render_rich_text(step.action, anchors)
        action = action or EMPTY_EXPECTED_RESULT
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
        return render_rich_text(step.expected_result, anchors) or EMPTY_EXPECTED_RESULT

    def render_labels(self, test_case: TestCase) -> list[str]:
        labels: list[str] = []
        seen: set[str] = set()
        for label in test_case.labels:
            sanitized = "".join("_" if character.isspace() else character for character in label)[:255]
            key = sanitized.casefold()
            if sanitized and key not in seen:
                labels.append(sanitized)
                seen.add(key)
        return labels

    def render_components(self, test_case: TestCase) -> list[str]:
        return []

    def render_custom_fields(self, test_case: TestCase) -> dict[str, Any]:
        if not test_case.process_path:
            return {}
        value = escape_jira_emoticons(PROCESS_PATH_SEPARATOR.join(test_case.process_path))
        return {PROCESS_PATH_FIELD: value}


def write_testcase_json(path: Path, payload: Mapping[str, Any]) -> None:
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
