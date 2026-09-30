"""Bildextraktion, PNG-Export, Nummerierung und Bildzuordnung."""

from __future__ import annotations

from pathlib import Path

from PIL import Image

from tests.fixtures.docx_factory import GIF, JPEG, PNG, DocSpec, StepSpec, build_lunar_docx
from tests.helpers import RunConvert, file_entry, load_report, load_testcase, warning_codes


def _all_step_attachments(testcase: dict) -> list[str]:
    return [name for step in testcase["steps"] for name in step["attachments"]]


def test_bild_in_info_tabelle(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(kurzbeschreibung_images=[PNG]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["screenshots"] == ["0001.png"]
    assert "|Kurzbeschreibung|Es wird ein Beispiel geprüft. \\\\ !0001.png!|" in testcase["description"]
    assert _all_step_attachments(testcase) == []


def test_bild_in_action_zelle(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(), StepSpec(number="20", action_images=[PNG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["steps"][1]["attachments"] == ["0001.png"]
    assert testcase["steps"][1]["action"] == "Transaktion aufrufen.\n!0001.png!"
    assert testcase["steps"][0]["attachments"] == []
    assert testcase["screenshots"] == []


def test_bild_in_expected_result_zelle(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(expected_images=[JPEG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    step = testcase["steps"][0]
    assert step["attachments"] == ["0001.png"]
    assert step["expected_result"] == "Maske wird angezeigt.\n!0001.png!"
    assert "!0001.png!" not in step["action"]
    assert testcase["screenshots"] == []
    with Image.open(output_dir / "a" / "screenshots" / "0001.png") as image:
        assert image.format == "PNG"


def test_inline_bild_im_fliesstext(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(inline_vml_image_in_action=PNG))

    assert run_convert() == 0

    step = load_testcase(output_dir, "a")["steps"][0]
    assert step["action"] == "Transaktion aufrufen.\nKlicken Sie auf !0001.png! ein."
    assert step["attachments"] == ["0001.png"]


def test_unklares_bild_wird_globaler_screenshot_ohne_anker(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action_images=[PNG])], body_images_after_table=[JPEG]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["screenshots"] == ["0002.png"]
    assert "!0002.png!" not in testcase["description"]
    assert all("!0002.png!" not in step["action"] + step["expected_result"] for step in testcase["steps"])
    assert _all_step_attachments(testcase) == ["0001.png"]
    assert "image_assignment_unclear" in warning_codes(file_entry(load_report(output_dir), "a.docx"))
    assert "WARNING" in (output_dir / "conversion.log").read_text(encoding="utf-8")


def test_bild_in_systemspalte_ist_unklar(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(system_images=[PNG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["screenshots"] == ["0001.png"]
    assert testcase["steps"][0]["attachments"] == []
    assert testcase["steps"][0]["system"] == "SAP FI"


def test_schwebendes_bild_ist_unklar(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action_floating_vml_image=PNG)]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["screenshots"] == ["0001.png"]
    assert testcase["steps"][0]["attachments"] == []
    assert "!0001.png!" not in testcase["steps"][0]["action"]


def test_nummerierung_startet_je_dokument_neu(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(kurzbeschreibung_images=[PNG], steps=[StepSpec(action_images=[JPEG, PNG])]))
    build_lunar_docx(input_dir / "b.docx", DocSpec(steps=[StepSpec(expected_images=[PNG])]))

    assert run_convert() == 0

    assert sorted(p.name for p in (output_dir / "a" / "screenshots").iterdir()) == ["0001.png", "0002.png", "0003.png"]
    assert sorted(p.name for p in (output_dir / "b" / "screenshots").iterdir()) == ["0001.png"]
    a = load_testcase(output_dir, "a")
    assert a["screenshots"] == ["0001.png"]
    assert a["steps"][0]["attachments"] == ["0002.png", "0003.png"]
    assert load_testcase(output_dir, "b")["steps"][0]["attachments"] == ["0001.png"]


def test_wiederholte_mediendatei_ergibt_eigene_dateien(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action_images=[PNG]), StepSpec(number="20", action_images=[PNG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["steps"][0]["attachments"] == ["0001.png"]
    assert testcase["steps"][1]["attachments"] == ["0002.png"]
    screenshots = output_dir / "a" / "screenshots"
    assert (screenshots / "0001.png").read_bytes() == (screenshots / "0002.png").read_bytes()


def test_nicht_unterstuetztes_bildformat(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action_images=[GIF, PNG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["steps"][0]["attachments"] == ["0001.png"]
    assert testcase["steps"][0]["action"] == "Transaktion aufrufen.\n!0001.png!"
    assert testcase["screenshots"] == []
    assert sorted(p.name for p in (output_dir / "a" / "screenshots").iterdir()) == ["0001.png"]
    entry = file_entry(load_report(output_dir), "a.docx")
    assert "unsupported_image_format" in warning_codes(entry)
    assert entry["exported_image_count"] == 1


def test_schrittanhang_nicht_in_globalen_screenshots(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(kurzbeschreibung_images=[PNG], steps=[StepSpec(action_images=[PNG])]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert set(testcase["screenshots"]).isdisjoint(_all_step_attachments(testcase))


def test_bilder_in_kopf_und_fusszeile_werden_ignoriert(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(
        input_dir / "a.docx",
        DocSpec(header_images=[JPEG], footer_images=[PNG, PNG], steps=[StepSpec(action_images=[PNG])]),
    )

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert sorted(p.name for p in (output_dir / "a" / "screenshots").iterdir()) == ["0001.png"]
    assert testcase["screenshots"] == []
    assert testcase["steps"][0]["attachments"] == ["0001.png"]
    assert "!0002.png!" not in str(testcase)
    assert "Logo" not in testcase["description"]
    entry = file_entry(load_report(output_dir), "a.docx")
    assert entry["exported_image_count"] == 1
    assert entry["ignored_header_footer_image_count"] == 3
    assert entry["warnings"] == []
    assert "3 Bild(er) in Kopf-/Fußzeilen" in (output_dir / "conversion.log").read_text(encoding="utf-8")


def test_nur_kopfzeilenbilder_ergeben_keine_screenshots(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(header_images=[PNG], footer_images=[JPEG]))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["screenshots"] == []
    assert all(step["attachments"] == [] for step in testcase["steps"])
    assert list((output_dir / "a" / "screenshots").iterdir()) == []
