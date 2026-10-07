"""Technische Extraktion und fachliches Parsing ohne CLI."""

from __future__ import annotations

import csv
from pathlib import Path

from docx import Document
from docx.oxml import parse_xml

from lunar_converter.docx_reader import read_docx
from lunar_converter.cli import main
from lunar_converter.lunar_parser import parse_test_case
from lunar_converter.models import ProfileDefinition, SourceTable
from lunar_converter.profile_loader import load_profiles
from lunar_converter.responsibles import cover_responsibles
from lunar_converter.image_assignment import assign_images
from lunar_converter.semantic_model import ImageMarker
from lunar_converter.target_renderer import XrayImportRenderer, render_wiki_table
from tests.fixtures.docx_factory import PNG, DocSpec, StepSpec, build_lunar_docx
from tests.helpers import CONFIG_DIR


def _standard_profile() -> ProfileDefinition:
    return next(profile for profile in load_profiles(CONFIG_DIR) if profile.id == "lunar_standard_v1")


def test_verantwortliche_aus_absatz_und_tabelle_nur_vom_deckblatt(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    text_doc = Document()
    text_doc.add_paragraph("Testfall: TF_001")
    text_doc.add_paragraph("Verantwortlicher: Markus Gerich")
    text_doc.add_table(rows=1, cols=1).cell(0, 0).text = "Fachbereich: Einkauf"
    text_doc.add_table(rows=1, cols=1).cell(0, 0).text = "Verantwortlicher: Nicht vom Deckblatt"
    text_path = input_dir / "text.docx"
    text_doc.save(text_path)

    table_doc = Document()
    cover = table_doc.add_table(rows=1, cols=1)
    cover.cell(0, 0).text = "Testfall: TF_002\nVerantwortlicher: Team RWWS EH1 (SP)"
    table_doc.add_table(rows=1, cols=1).cell(0, 0).text = "Fachbereich: Einkauf"
    table_doc.add_table(rows=1, cols=1).cell(0, 0).text = "Verantwortlicher: Nicht vom Deckblatt"
    table_path = input_dir / "tabelle.docx"
    table_doc.save(table_path)

    assert cover_responsibles(read_docx(text_path)) == ["Markus Gerich"]
    assert cover_responsibles(read_docx(table_path)) == ["Team RWWS EH1 (SP)"]

    csv_path = tmp_path / "verantwortliche.csv"
    assert main(["verantwortliche", "--input-dir", str(input_dir), "--csv-path", str(csv_path)]) == 0
    with csv_path.open(encoding="utf-8-sig", newline="") as output:
        assert list(csv.reader(output, delimiter=";")) == [
            ["verantwortlicher", "datei"],
            ["Team RWWS EH1 (SP)", "tabelle.docx"],
            ["Markus Gerich", "text.docx"],
        ]
    log_text = (tmp_path / "verantwortliche.verantwortliche.log").read_text(encoding="utf-8")
    assert "Start der Verantwortlichen-Erfassung: 2 Datei(en)" in log_text
    assert "[Datei: tabelle.docx]" in log_text
    assert "Verantwortliche gefunden: Team RWWS EH1 (SP)." in log_text
    assert "Ende der Verantwortlichen-Erfassung: 2 gesamt, 0 ohne Ergebnis." in log_text
    assert main(["verantwortliche", "--input-dir", str(input_dir), "--csv-path", str(csv_path)]) == 2


def test_verantwortliche_protokolliert_fehlenden_fund(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    document = Document()
    document.add_paragraph("Testfall: TF_ohne_Verantwortlichen")
    document.save(input_dir / "ohne.docx")
    csv_path = tmp_path / "liste.csv"

    assert main(["verantwortliche", "--input-dir", str(input_dir), "--csv-path", str(csv_path)]) == 1
    log_text = (tmp_path / "liste.verantwortliche.log").read_text(encoding="utf-8")
    assert "[Datei: ohne.docx]" in log_text
    assert "Kein Verantwortlicher auf dem Deckblatt gefunden" in log_text
    assert "Ende der Verantwortlichen-Erfassung: 1 gesamt, 1 ohne Ergebnis." in log_text


def test_verantwortliche_aus_kontaktblock_und_verantwortlichem_team(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    legacy = Document()
    title = legacy.add_table(rows=1, cols=1)
    title.cell(0, 0).text = "EG930_RWWS2.0_EH\nTestfall: EG930_0704001"
    contact = legacy.add_table(rows=2, cols=2)
    contact.cell(0, 0).text = "Angela Hoppe\nT: +49/40/6377-8986\nF: +49/40/6377-xxx\nangela.hoppe@edeka.de"
    contact.cell(1, 0).text = "STATUS\nin Bearbeitung"
    contact.cell(1, 1).text = "VERSION\n1.0"
    legacy.add_table(rows=1, cols=1).cell(0, 0).text = "Verantwortlicher: Späterer Inhalt"
    legacy_path = input_dir / "legacy.docx"
    legacy.save(legacy_path)

    team = Document()
    team.add_paragraph("Testfall: TFB_03.02.001.01_GH")
    team.add_table(rows=1, cols=1).cell(0, 0).text = "Verantwortliches Team: SKA"
    team.add_table(rows=1, cols=1).cell(0, 0).text = "Verantwortlicher: Späterer Inhalt"
    team_path = input_dir / "team.docx"
    team.save(team_path)

    assert cover_responsibles(read_docx(legacy_path)) == ["Angela Hoppe"]
    assert cover_responsibles(read_docx(team_path)) == ["SKA"]
    csv_path = tmp_path / "liste.csv"
    assert main(["verantwortliche", "--input-dir", str(input_dir), "--csv-path", str(csv_path)]) == 0
    with csv_path.open(encoding="utf-8-sig", newline="") as output:
        assert list(csv.reader(output, delimiter=";")) == [
            ["verantwortlicher", "datei"],
            ["Angela Hoppe", "legacy.docx"],
            ["SKA", "team.docx"],
        ]


def test_verantwortliche_nicht_aus_unbeschriftetem_statusfeld(tmp_path: Path) -> None:
    document = Document()
    document.add_paragraph("Testfall: TF_001")
    document.add_table(rows=1, cols=1).cell(0, 0).text = "STATUS\nin Bearbeitung\nVERSION\nservice@example.de"
    path = tmp_path / "status.docx"
    document.save(path)

    assert cover_responsibles(read_docx(path)) == []


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
    assert test_case.responsible_names == ["Max Mustermann"]
    assert test_case.process_path == [
        "03.02 Einkaufsverwaltung",
        "03.02.001 Pflege Einkaufskonditionen",
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
    assert test_case.steps[0].actual_result.is_empty()
    assert test_case.steps[1].system == "BW"
    assert test_case.unassigned_images == []


def test_deckblatt_beschreibung_zwischen_prozessebenen_bleibt_ausserhalb_des_pfad(tmp_path: Path) -> None:
    path = build_lunar_docx(
        tmp_path / "beschreibung_zwischen_prozessen.docx",
        DocSpec(
            process_lines=("03.02 Einkauf", "03.02.001 Konditionen"),
            title_lines=("Fachliche Beschreibung",),
        ),
    )
    document = Document(str(path))
    title_cell = document.tables[0].cell(1, 0)
    insertion = title_cell.add_paragraph("Zwischenbeschreibung")
    title_cell.paragraphs[1]._p.addprevious(insertion._p)
    document.save(str(path))

    test_case = parse_test_case(read_docx(path), _standard_profile())

    assert test_case.process_path == ["03.02 Einkauf", "03.02.001 Konditionen"]
    assert "Zwischenbeschreibung" in test_case.title
    assert "Fachliche Beschreibung" in test_case.title


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


def test_wingdings_word_listen_einrueckung_und_warnung(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "formatierung.docx", DocSpec(steps=[StepSpec(action="Normaler Text bleibt.")]))
    document = Document(str(path))
    info_table = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    info_paragraph = info_table.cell(1, 1).paragraphs[0]
    info_paragraph.add_run(" ")
    info_paragraph.add_run("\uf0e0").font.name = "Wingdings"

    step_table = next(table for table in document.tables if any(cell.text == "Beschreibung des Testschritts" for cell in table.rows[0].cells))
    action_cell = step_table.cell(1, 2)
    action_cell.add_paragraph("Erfassen Sie:")
    action_cell.add_paragraph("Wertartikel", style="List Bullet")
    action_cell.add_paragraph("Buchungskreis", style="List Bullet 2")
    action_cell.add_paragraph("\t\tBetrieb")
    arrow = action_cell.add_paragraph().add_run("\uf0d8")
    arrow.font.name = "Wingdings"
    unknown = action_cell.add_paragraph().add_run("\uf0b0")
    unknown.font.name = "Wingdings"
    document.save(str(path))

    source = read_docx(path)
    test_case = parse_test_case(source, _standard_profile())
    payload = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)

    assert "→" in payload["description"]
    action = payload["steps"][0]["action"]
    assert "Normaler Text bleibt." in action
    assert "• Wertartikel" in action
    assert "↳ • Buchungskreis" in action
    assert "↳ ↳ Betrieb" in action
    assert "→" in action
    assert "[Wingdings U+F0B0]" in action
    warning = next(issue for issue in source.warnings if issue.code == "unsupported_wingdings_glyph")
    assert warning.message.startswith("Tabelle ") and "U+F0B0" in warning.message
