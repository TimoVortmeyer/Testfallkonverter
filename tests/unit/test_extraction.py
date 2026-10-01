"""Technische Extraktion und fachliches Parsing ohne CLI."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.oxml import parse_xml

from lunar_converter.docx_reader import read_docx
from lunar_converter.lunar_parser import parse_test_case
from lunar_converter.models import ProfileDefinition, SourceTable
from lunar_converter.profile_loader import load_profiles
from lunar_converter.semantic_model import ImageMarker
from lunar_converter.target_renderer import render_wiki_table
from tests.fixtures.docx_factory import PNG, DocSpec, StepSpec, build_lunar_docx
from tests.helpers import CONFIG_DIR


def _standard_profile() -> ProfileDefinition:
    return next(profile for profile in load_profiles(CONFIG_DIR) if profile.id == "lunar_standard_v1")


def test_bildreihenfolge_und_positionen(tmp_path: Path) -> None:
    path = build_lunar_docx(
        tmp_path / "a.docx",
        DocSpec(kurzbeschreibung_images=[PNG], steps=[StepSpec(action_images=[PNG], expected_images=[PNG])], body_images_after_table=[PNG]),
    )
    document = read_docx(path)

    assert [image.image_id for image in document.images] == [1, 2, 3, 4]
    assert "Tabelle 2" in document.images[0].location
    assert document.images[1].location.endswith("Zelle 3")
    assert document.images[2].location.endswith("Zelle 4")
    assert document.images[3].location.startswith("Absatz")
    assert all(image.blob and image.part_name for image in document.images)


def test_parser_felder_und_schritte(tmp_path: Path) -> None:
    path = build_lunar_docx(
        tmp_path / "a.docx",
        DocSpec(steps=[StepSpec(action="Zeile 1\nZeile 2", actual="Tatsächlich"), StepSpec(number="20", system="BW")]),
    )
    document = read_docx(path)
    test_case = parse_test_case(document, _standard_profile())

    assert test_case.name == "TF_Beispiel_001"
    assert test_case.process_path == [
        "03.02 Einkaufsverwaltung",
        "03.02.001 Pflege Einkaufskonditionen",
        "Prüfung der",
        "Beispielkonditionen",
    ]
    assert test_case.title == "Prüfung der Beispielkonditionen"
    assert test_case.info_table is not None
    assert [[cell.plain_text() for cell in row] for row in test_case.info_table.rows] == [
        ["Fachbereich", "Finanzen"],
        ["Kurzbeschreibung", "Es wird ein Beispiel geprüft."],
        ["Voraussetzungen", "Stammdaten sind vorhanden."],
    ]
    assert [step.step_number for step in test_case.steps] == ["10", "20"]
    assert test_case.steps[0].action.plain_text() == "Zeile 1\nZeile 2"
    assert test_case.steps[0].actual_result.plain_text() == "Tatsächlich"
    assert test_case.steps[1].system == "BW"
    assert test_case.unassigned_images == []


def test_parser_merkt_bild_in_action(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "a.docx", DocSpec(steps=[StepSpec(action_images=[PNG])]))
    document = read_docx(path)
    test_case = parse_test_case(document, _standard_profile())

    assert test_case.steps[0].action.lines[-1] == [ImageMarker(1)]
    assert isinstance(document.blocks[-1], SourceTable)


def test_info_tabelle_schuetzt_emoticons_und_rendert_checkboxen(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "checkboxen.docx")
    document = Document(str(path))
    info_table = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    paragraph = info_table.cell(1, 1).paragraphs[0]
    paragraph.add_run(" Betriebsgruppe")
    paragraph.add_run("(")
    paragraph.add_run("n)")
    for checked in (True, False):
        paragraph.add_run(" ")
        checked_xml = '<w14:checked w14:val="1"/>' if checked else ""
        glyph = "☒" if checked else "☐"
        paragraph._p.append(
            parse_xml(
                '<w:sdt xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
                'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
                f"<w:sdtPr><w14:checkbox>{checked_xml}</w14:checkbox></w:sdtPr>"
                f"<w:sdtContent><w:r><w:t>{glyph}</w:t></w:r></w:sdtContent></w:sdt>"
            )
        )
    w_ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    # Wie in den LUNAR-Vorlagen: FORMCHECKBOX ohne Symbol, Zustand nur in w:default bzw. w:checked.
    for state_xml in ('<w:default w:val="1"/>', '<w:default w:val="0"/>', '<w:default w:val="1"/><w:checked w:val="0"/>'):
        paragraph._p.append(
            parse_xml(
                f"<w:r {w_ns}><w:fldChar w:fldCharType=\"begin\"><w:ffData><w:checkBox><w:sizeAuto/>"
                f"{state_xml}</w:checkBox></w:ffData></w:fldChar></w:r>"
            )
        )
        paragraph._p.append(parse_xml(f'<w:r {w_ns}><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r>'))
        paragraph._p.append(parse_xml(f'<w:r {w_ns}><w:fldChar w:fldCharType="separate"/></w:r>'))
        paragraph._p.append(parse_xml(f'<w:r {w_ns}><w:fldChar w:fldCharType="end"/></w:r>'))
        paragraph.add_run(" System")
    paragraph.add_run(" ☐ ☒")
    document.save(str(path))

    source = read_docx(path)
    test_case = parse_test_case(source, _standard_profile())

    assert test_case.info_table is not None
    rendered = render_wiki_table(test_case.info_table, {})
    assert r"Betriebsgruppe\(n)" in rendered
    assert "(/) (x)(/) System(x) System(x) System (x) (/)" in rendered
