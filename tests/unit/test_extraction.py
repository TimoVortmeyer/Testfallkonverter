"""Technische Extraktion und fachliches Parsing ohne CLI."""

from __future__ import annotations

import csv
from pathlib import Path

import pytest
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import parse_xml
from docx.shared import Inches, RGBColor

from lunar_converter.docx_reader import read_docx
from lunar_converter.exceptions import InputFileError
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

    assert test_case.name == "TFB_Beispiel_001"
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
    indented = action_cell.add_paragraph("Absatz mit Einzug")
    indented.paragraph_format.left_indent = Inches(0.5)
    action_cell.add_paragraph("Nummeriert", style="List Number")
    action_cell.add_paragraph("Unternummeriert", style="List Number 2")
    arrow = action_cell.add_paragraph().add_run("\uf0d8")
    arrow.font.name = "Wingdings"
    expected_arrow = action_cell.add_paragraph().add_run("\uf0f0")
    expected_arrow.font.name = "Wingdings"
    unknown = action_cell.add_paragraph().add_run("\uf0b0")
    unknown.font.name = "Wingdings"
    document.save(str(path))

    source = read_docx(path)
    test_case = parse_test_case(source, _standard_profile())
    payload = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)

    assert "→" in payload["description"]
    action = payload["steps"][0]["action"]
    assert "Normaler Text bleibt." in action
    assert "* Wertartikel" in action
    assert "** Buchungskreis" in action
    assert "Betrieb" in action and "↳" not in action
    assert "Absatz mit Einzug" in action and "↳" not in action
    assert "↳" not in action
    assert "# Nummeriert" in action
    assert "## Unternummeriert" in action
    assert "→" in action
    assert "[Wingdings" not in action
    assert action.count("→") >= 2
    warning = next(issue for issue in source.warnings if issue.code == "unsupported_wingdings_glyph")
    assert warning.message.startswith("Tabelle ") and "U+F0B0" in warning.message


def test_inline_formatierung_bleibt_in_fachtext_erhalten_und_labels_bleiben_plain(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "inline.docx")
    document = Document(str(path))
    info = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    info.cell(1, 0).paragraphs[0].runs[0].bold = True
    value = info.cell(1, 1).paragraphs[0]
    value.clear()
    value.add_run("Fett").bold = True
    value.add_run(" kursiv").italic = True
    value.add_run(" unterstrichen").underline = True
    value.add_run("\tTabulator")
    steps = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    action = steps.cell(1, 2).paragraphs[0]
    action.clear()
    action.add_run("Fett").bold = True
    action.add_run(" kursiv").italic = True
    action.add_run(" unterstrichen").underline = True
    action.add_run(" Platzhalter <Marktnummer>")
    expected = steps.cell(1, 3).paragraphs[0]
    expected.clear()
    expected.add_run("Erwartet").bold = True
    expected.add_run(" kursiv").italic = True
    expected.add_run(" unterstrichen").underline = True
    expected.add_run("\tweiter")
    document.save(str(path))

    source = read_docx(path)
    test_case = parse_test_case(source, _standard_profile())
    payload = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)

    assert "|Kurzbeschreibung|*Fett* _kursiv_ +unterstrichen+ Tabulator|" in payload["description"]
    assert "*Kurzbeschreibung*" not in payload["description"]
    assert payload["steps"][0]["action"] == "*Fett* _kursiv_ +unterstrichen+ Platzhalter <Marktnummer>"
    assert payload["steps"][0]["expected_result"] == "*Erwartet* _kursiv_ +unterstrichen+ weiter"
    assert payload["source_word_filename"] == "inline.docx"


def test_nested_info_table_is_rejected_with_outer_cell_location(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "nested.docx")
    document = Document(str(path))
    info = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    nested = info.cell(1, 1).add_table(rows=3, cols=2)
    nested.cell(0, 0).text, nested.cell(0, 1).text = "Bereich:", "EDEKA"
    nested.cell(0, 1).paragraphs[0].runs[0].bold = True
    nested.cell(1, 0).text, nested.cell(1, 1).text = "Filiale", "1234"
    nested.cell(2, 0).text, nested.cell(2, 1).text = "", ""
    info.cell(2, 1).text = "<Sicherung>"
    info.cell(2, 1).add_table(rows=1, cols=2)
    steps = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    steps.cell(1, 2).text = "Platzhalter <Marktnummer> im Fachtext"
    document.save(str(path))

    with pytest.raises(InputFileError) as raised:
        read_docx(path)

    assert raised.value.code == "nested_table_not_supported"
    assert "Tabelle 2" in (raised.value.field or "")
    assert "Zeile 2, Zelle 2" in (raised.value.field or "")
    assert "Zeile 3, Zelle 2" in (raised.value.field or "")
    assert "2 innere Tabelle(n)" in (raised.value.details or "")


@pytest.mark.parametrize(
    ("progid", "expected_error"),
    [("Word.Document.12", True), ("Excel.Sheet.12", False)],
)
def test_embedded_word_document_is_rejected_but_other_ole_object_is_not(
    tmp_path: Path, progid: str, expected_error: bool
) -> None:
    path = build_lunar_docx(tmp_path / "ole.docx")
    document = Document(str(path))
    steps = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    paragraph = steps.cell(1, 2).paragraphs[0]
    paragraph._p.append(
        parse_xml(
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            'xmlns:o="urn:schemas-microsoft-com:office:office" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<w:object><o:OLEObject Type="Embed" ProgID="'
            f'{progid}'
            '" r:id="rIdEmbedded"/></w:object></w:r>'
        )
    )
    document.save(str(path))

    if expected_error:
        with pytest.raises(InputFileError) as raised:
            read_docx(path)
        assert raised.value.code == "embedded_word_document_not_processed"
        assert "Tabelle 3, Zeile 2, Zelle 3" in (raised.value.field or "")
        assert "Word.Document.12" in (raised.value.details or "")
        assert "Eingebettetes Word-Dokument" in raised.value.message
    else:
        assert read_docx(path).blocks


def test_bullets_in_info_cell_are_separate_jira_list_items(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "info_liste.docx")
    document = Document(str(path))
    info = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    cell = info.cell(1, 1)
    cell.paragraphs[0].text = "Eintrag 1"
    cell.paragraphs[0].style = "List Bullet"
    for text in ("Eintrag 2", "Eintrag 3", "Eintrag 4"):
        cell.add_paragraph(text, style="List Bullet")
    document.save(str(path))

    test_case = parse_test_case(read_docx(path), _standard_profile())
    payload = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)

    assert "|Kurzbeschreibung|* Eintrag 1\n* Eintrag 2\n* Eintrag 3\n* Eintrag 4|" in payload["description"]
    assert " \\\\ * Eintrag" not in payload["description"]


def test_consecutive_manual_bullet_paragraphs_stay_separate(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "manual_liste.docx")
    document = Document(str(path))
    info = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    cell = info.cell(1, 1)
    cell.paragraphs[0].text = "- Stammdaten 1"
    for text in ("- Stammdaten 2", "- Stammdaten 3", "- Stammdaten 4"):
        cell.add_paragraph(text)
    document.save(str(path))

    case = parse_test_case(read_docx(path), _standard_profile())
    rendered = XrayImportRenderer().render(case, assign_images(case, {}), path.stem)["description"]
    assert all(f"- Stammdaten {index}" in rendered for index in range(1, 5))
    assert rendered.count("- Stammdaten") == 4
    assert " \\\\ - Stammdaten" not in rendered


def test_bracketed_cover_process_segments_are_normalized_before_path_detection(tmp_path: Path) -> None:
    path = build_lunar_docx(
        tmp_path / "prozess.docx",
        DocSpec(
            process_lines=(
                "<06.04. Verkaufsabwicklung>",
                "<06.04.003 GH Kundenauftragsabwicklung >",
                "06.04.003 GH Tabak Rückverfolgbarkeit",
                "<18-00144-004 Tabak Rückverfolgbarkeit>",
            )
        ),
    )

    case = parse_test_case(read_docx(path), _standard_profile())
    payload = XrayImportRenderer().render(case, assign_images(case, {}), path.stem)

    assert case.process_path == [
        "06.04. Verkaufsabwicklung",
        "06.04.003 GH Kundenauftragsabwicklung",
        "06.04.003 GH Tabak Rückverfolgbarkeit",
        "18-00144-004 Tabak Rückverfolgbarkeit",
    ]
    assert payload["custom_fields"]["customfield_15909"] == "/".join(case.process_path)


def test_bracketed_process_line_with_trailing_period_and_inline_testcase_label(tmp_path: Path) -> None:
    path = build_lunar_docx(
        tmp_path / "inline_cover.docx",
        DocSpec(process_lines=(), title_lines=()),
    )
    document = Document(str(path))
    cover = document.tables[0].cell(1, 0)
    cover.text = "<06.04. Verkaufsabwicklung>\n<06.04.003 GH Kundenauftragsabwicklung>\n<18-00144-004 Tabak Rückverfolgbarkeit> Testfall: TFB_1"
    document.save(str(path))

    case = parse_test_case(read_docx(path), _standard_profile())

    assert case.name == "TFB_1"
    assert case.process_path == [
        "06.04. Verkaufsabwicklung",
        "06.04.003 GH Kundenauftragsabwicklung",
        "18-00144-004 Tabak Rückverfolgbarkeit",
    ]


def test_word_color_is_not_guessed_as_jira_markup_and_is_warned(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "farbe.docx")
    document = Document(str(path))
    info = next(table for table in document.tables if any(cell.text == "Kurzbeschreibung" for row in table.rows for cell in row.cells))
    run = info.cell(1, 1).paragraphs[0].runs[0]
    run.font.color.rgb = RGBColor(0xFF, 0, 0)
    steps = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    action_paragraph = steps.cell(1, 2).paragraphs[0]
    action_paragraph.runs[0].font.color.rgb = RGBColor(0xFF, 0, 0)
    action_paragraph.add_run(" blau").font.color.rgb = RGBColor(0, 0, 0xFF)
    expected_style = document.styles.add_style("ExpectedRed", WD_STYLE_TYPE.CHARACTER)
    expected_style.font.color.rgb = RGBColor(0xFF, 0, 0)
    steps.cell(1, 3).paragraphs[0].runs[0].style = expected_style
    document.save(str(path))

    source = read_docx(path)
    case = parse_test_case(source, _standard_profile())
    payload = XrayImportRenderer().render(case, assign_images(case, {}), path.stem)

    assert "Es wird ein Beispiel geprüft." in payload["description"]
    assert "Transaktion aufrufen." in payload["steps"][0]["action"]
    assert "blau" in payload["steps"][0]["action"]
    assert "Maske wird angezeigt." in payload["steps"][0]["expected_result"]
    assert "{color" not in payload["description"] + payload["steps"][0]["action"]
    color_warnings = [issue for issue in source.warnings if issue.code == "unsupported_text_color"]
    assert len(color_warnings) >= 4
    assert any("ff0000" in issue.message for issue in color_warnings)
    assert any("0000ff" in issue.message for issue in color_warnings)


def test_tracked_insertions_are_kept_and_deletions_ignored(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "revision.docx")
    document = Document(str(path))
    step_table = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    paragraph = step_table.cell(1, 3).paragraphs[0]
    paragraph.clear()
    paragraph._p.append(parse_xml('<w:del xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:delText>alter Text</w:delText></w:r></w:del>'))
    paragraph._p.append(parse_xml('<w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>sichtbarer aktueller Text</w:t></w:r></w:ins>'))
    document.save(str(path))

    test_case = parse_test_case(read_docx(path), _standard_profile())
    rendered = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)["steps"][0]["expected_result"]

    assert "sichtbarer aktueller Text" in rendered
    assert "alter Text" not in rendered


def test_word_comment_markers_are_not_imported_as_fachtext(tmp_path: Path) -> None:
    path = build_lunar_docx(tmp_path / "comment.docx")
    document = Document(str(path))
    step_table = next(table for table in document.tables if table.rows[0].cells[0].text == "Schritt-Nr.")
    paragraph = step_table.cell(1, 2).paragraphs[0]
    paragraph.clear()
    paragraph.add_run("Sichtbarer Text ")
    paragraph._p.append(parse_xml('<w:commentRangeStart xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="4"/>'))
    paragraph.add_run("mit Kommentar")
    paragraph._p.append(parse_xml('<w:commentRangeEnd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:id="4"/>'))
    paragraph._p.append(parse_xml('<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:commentReference w:id="4"/></w:r>'))
    document.save(str(path))

    test_case = parse_test_case(read_docx(path), _standard_profile())
    action = XrayImportRenderer().render(test_case, assign_images(test_case, {}), path.stem)["steps"][0]["action"]

    assert action == "Sichtbarer Text mit Kommentar"
    assert "commentReference" not in action
