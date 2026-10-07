"""Ziel-Renderer und Bildzuordnung auf Ebene des semantischen Modells."""

from __future__ import annotations

from pathlib import Path

from lunar_converter.image_assignment import assign_images
from lunar_converter.semantic_model import CheckboxMarker, ImageMarker, InfoTable, RichText, TestCase, TestStep, UnassignedImage
from lunar_converter.target_renderer import XrayImportRenderer, render_rich_text, render_wiki_table


def _test_case() -> TestCase:
    return TestCase(
        source_file=Path("a.docx"),
        profile_id="lunar_standard_v1",
        name="TF_1",
        process_path=["03.02 Einkauf", "03.02.001 Konditionen", "03.02.001.01 GH"],
        title="Prüfung Konditionen",
        info_table=InfoTable(
            rows=[
                [RichText.from_text("Kurzbeschreibung"), RichText(lines=[["Text"], [ImageMarker(1)]])],
                [RichText.from_text("Voraussetzungen"), RichText()],
            ]
        ),
        steps=[
            TestStep(
                index=0,
                source_location="Tabelle 1, Zeile 2",
                system="SAP",
                action=RichText(lines=[["Klick auf", ImageMarker(2), "Knopf"]]),
                expected_result=RichText.from_text("OK"),
                actual_result=RichText(lines=[["Wie erwartet"], [ImageMarker(3)]]),
            )
        ],
        unassigned_images=[UnassignedImage(image_id=4, reason="Test", location="Absatz 9")],
    )


def test_render_rich_text_setzt_anker_mit_leerzeichen() -> None:
    rich = RichText(lines=[["über", ImageMarker(1), "."], ["", ImageMarker(2)]])
    assert render_rich_text(rich, {1: "0001.png", 2: "0002.png"}) == "über !0001.png!.\n!0002.png!"
    assert render_rich_text(rich, {}) == "über."


def test_normale_ausrufezeichen_bleiben_im_fachtext_erhalten() -> None:
    text = "Erfolgreich!\n\nAchtung!!\n\nBitte notieren!!"

    assert render_rich_text(RichText.from_text(text), {}) == text


def test_renderer_komplett() -> None:
    exported = {1: "0001.png", 2: "0002.png", 3: "0003.png", 4: "0004.png"}
    assignment = assign_images(_test_case(), exported)
    payload = XrayImportRenderer().render(_test_case(), assignment, "a")

    assert payload["summary"] == "TF_1"
    assert payload["description"] == "h1. Prüfung Konditionen\n\n|Kurzbeschreibung|Text \\\\ !0001.png!|\n|Voraussetzungen| |"
    assert payload["custom_fields"] == {"customfield_15909": "03.02 Einkauf/03.02.001 Konditionen/03.02.001.01 GH"}
    assert payload["screenshots"] == ["0001.png", "0004.png"]
    step = payload["steps"][0]
    assert step["action"] == "Klick auf !0002.png! Knopf"
    assert step["expected_result"] == "OK"
    assert step["attachments"] == ["0002.png"]
    assert step["data"] == "" and step["tester"] == "" and step["screenshots"] == []
    assert [w.code for w in assignment.warnings] == ["image_assignment_unclear"]


def test_labels_ersetzen_whitespace_und_begrenzen_laenge() -> None:
    case = _test_case()
    case.labels = ["FICO-EH", "RWWS Team", "Årea_2/东", "x" * 300, "x" * 255, "RWWS_Team"]

    rendered = XrayImportRenderer().render_labels(case)
    assert rendered[:3] == ["FICO-EH", "RWWS_Team", "Årea_2/东"]
    assert rendered[3] == "x" * 255
    assert len(rendered) == 4


def test_nicht_exportierte_bilder_erhalten_keinen_anker() -> None:
    assignment = assign_images(_test_case(), {1: "0001.png"})
    payload = XrayImportRenderer().render(_test_case(), assignment, "a")
    assert payload["screenshots"] == ["0001.png"]
    assert payload["steps"][0]["action"] == "Klick auf Knopf"
    assert payload["steps"][0]["attachments"] == []


def test_description_und_custom_fields_ohne_deckblattdaten() -> None:
    case = _test_case()
    case.title = ""
    case.info_table = None
    case.process_path = []
    assert XrayImportRenderer().render_description(case, {}) == ""
    assert XrayImportRenderer().render_custom_fields(case) == {}


def test_wiki_tabelle_maskiert_pipes_und_zeilenumbrueche() -> None:
    table = InfoTable(rows=[[RichText.from_text("A|B"), RichText(lines=[["Zeile 1"], [""], ["Zeile 2"]])]])
    assert render_wiki_table(table, {}) == "|A\\|B|Zeile 1 \\\\ Zeile 2|"


def test_checkbox_emoticons_sind_auf_info_tabelle_begrenzt() -> None:
    checked = RichText(lines=[[CheckboxMarker(checked=True)]])
    unchecked = RichText(lines=[[CheckboxMarker(checked=False)]])
    table = InfoTable(rows=[[checked, unchecked]])

    assert render_wiki_table(table, {}) == "|(/)|(x)|"
    assert render_rich_text(checked, {}) == "☒"
    assert render_rich_text(unchecked, {}) == "☐"
