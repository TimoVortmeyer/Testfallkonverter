"""Batch-Verarbeitung, Output-Regeln, Report und Dry Run."""

from __future__ import annotations

import csv
from pathlib import Path

from lunar_converter.cli import main
from tests.fixtures.docx_factory import PNG, DocSpec, StepSpec, build_lunar_docx, build_poc_sample_docx
from tests.helpers import (
    CONFIG_DIR,
    SCHEMA_PATH,
    RunConvert,
    error_codes,
    file_entry,
    final_entries,
    load_report,
    load_testcase,
)


def test_erfolgreiche_verarbeitung_einer_docx(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(
        input_dir / "TFB_03.03.004_EH_012_MDE_Direktbestellung.docx",
        DocSpec(steps=[StepSpec(actual="Wie erwartet.")]),
    )

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "TFB_03.03.004_EH_012_MDE_Direktbestellung")
    assert testcase["summary"] == "TFB_03.03.004_EH_012_MDE_Direktbestellung"
    assert testcase["description"] == (
        "h1. Prüfung der Beispielkonditionen\n\n"
        "|Fachbereich|Finanzen|\n"
        "|Kurzbeschreibung|Es wird ein Beispiel geprüft.|\n"
        "|Voraussetzungen|Stammdaten sind vorhanden.|"
    )
    assert testcase["labels"] == ["RWWS", "Max Mustermann"] and testcase["components"] == []
    assert "reporter_email" not in testcase
    assert testcase["custom_fields"] == {
        "customfield_15909": "03.02 Einkaufsverwaltung/03.02.001 Pflege Einkaufskonditionen/Prüfung der/Beispielkonditionen"
    }
    assert testcase["steps"] == [
        {
            "system": "SAP FI",
            "action": "Transaktion aufrufen.",
            "data": "",
            "expected_result": "Maske wird angezeigt.\n\nh3. Tatsächliches Ergebnis\nWie erwartet.",
            "tester": "",
            "attachments": [],
            "screenshots": [],
        }
    ]
    assert (output_dir / "TFB_03.03.004_EH_012_MDE_Direktbestellung" / "screenshots").is_dir()
    assert (output_dir / "conversion.log").is_file()


def test_responsibles_csv_fuellt_reporter_email_oder_label(
    input_dir: Path, output_dir: Path, run_convert: RunConvert, tmp_path: Path
) -> None:
    build_lunar_docx(
        input_dir / "FiCo EH" / "Debitoren.docx",
        DocSpec(responsible="Markus Gerich"),
    )
    build_lunar_docx(
        input_dir / "SKA" / "Konditionen.docx",
        DocSpec(responsible="SKA"),
    )
    mapping_csv = tmp_path / "verantwortliche-email.csv"
    with mapping_csv.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.writer(output, delimiter=";")
        writer.writerow(["verantwortlicher", "datei", "Gefunden", "Email", "Status"])
        writer.writerow(["Markus Gerich", "FiCo EH/Debitoren.docx", "True", "markus.gerich@edeka.de", ""])
        writer.writerow(["SKA", "SKA/Konditionen.docx", "False", "", "Name/Gruppe nicht gefunden"])

    assert run_convert("--responsibles-csv", str(mapping_csv)) == 0

    mapped = load_testcase(output_dir, "Debitoren")
    assert mapped["reporter_email"] == "markus.gerich@edeka.de"
    assert mapped["labels"] == ["RWWS"]
    unmapped = load_testcase(output_dir, "Konditionen")
    assert "reporter_email" not in unmapped
    assert unmapped["labels"] == ["RWWS", "SKA"]


def test_poc_beispiel_wird_weiterhin_verarbeitet(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_poc_sample_docx(input_dir / "sample.docx")

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "sample")
    assert testcase["summary"] == "sample"
    assert testcase["labels"] == []
    # Ohne Deckblatt-Prozesszeilen und ohne Tabelle vor dem Testablauf bleibt die Description leer.
    assert testcase["description"] == ""
    assert testcase["custom_fields"] == {}
    assert testcase["steps"][0]["system"] == "System A"
    assert testcase["steps"][0]["action"] == "Artikel erfassen"
    assert testcase["steps"][0]["expected_result"] == "Maske erscheint"


def test_nicht_leerer_output_ordner_bricht_global_ab(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx")
    output_dir.mkdir()
    existing = output_dir / "alt.txt"
    existing.write_text("bestehend", encoding="utf-8")

    assert run_convert() == 2

    assert sorted(p.name for p in output_dir.iterdir()) == ["alt.txt"]
    assert existing.read_text(encoding="utf-8") == "bestehend"


def test_nicht_existierender_output_ordner_wird_angelegt(input_dir: Path, tmp_path: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx")
    assert run_convert() == 0
    assert (tmp_path / "output" / "a" / "testcase.json").is_file()


def test_fehler_einer_datei_bricht_batch_nicht_ab(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    (input_dir / "a_defekt.docx").write_bytes(b"kein zip")
    build_lunar_docx(input_dir / "b_ok.docx")

    assert run_convert() == 1

    assert final_entries(output_dir) == ["b_ok"]
    report = load_report(output_dir)
    assert [Path(entry["input_file"]).name for entry in report["files"]] == ["a_defekt.docx", "b_ok.docx"]
    assert file_entry(report, "a_defekt.docx")["status"] == "failed"
    assert error_codes(file_entry(report, "a_defekt.docx")) == ["input_file_error"]
    assert file_entry(report, "b_ok.docx")["status"] == "success"


def test_fail_fast_ueberspringt_restliche_dateien(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    (input_dir / "a_defekt.docx").write_bytes(b"kein zip")
    build_lunar_docx(input_dir / "b_ok.docx")

    assert run_convert("--fail-fast") == 1

    assert final_entries(output_dir) == []
    report = load_report(output_dir)
    assert file_entry(report, "b_ok.docx")["status"] == "skipped"
    assert report["summary"] == {"total": 2, "success": 0, "failed": 1, "skipped": 1}


def test_bei_fehler_kein_finaler_testfallordner(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(
        input_dir / "fehlerhaft.docx",
        DocSpec(steps=[StepSpec(action="", action_images=[], expected_images=[PNG])]),
    )

    assert run_convert() == 1

    assert final_entries(output_dir) == []
    assert not any(p.name.startswith(".") for p in output_dir.iterdir())


def test_unterordner_werden_verarbeitet_output_bleibt_flach(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "Bereich A" / "Fall_1.docx", DocSpec(name="TF_A1"))
    build_lunar_docx(input_dir / "Bereich B" / "Tief" / "Fall_2.docx", DocSpec(name="TF_B2"))
    build_lunar_docx(input_dir / "Bereich B" / "Fall_1.docx", DocSpec(name="TF_B1"))

    assert run_convert() == 1

    report = load_report(output_dir)
    assert [Path(entry["input_file"]).relative_to(input_dir).as_posix() for entry in report["files"]] == [
        "Bereich A/Fall_1.docx",
        "Bereich B/Fall_1.docx",
        "Bereich B/Tief/Fall_2.docx",
    ]
    assert final_entries(output_dir) == ["Fall_1", "Fall_2"]
    assert load_testcase(output_dir, "Fall_1")["summary"] == "Fall_1"
    conflict = report["files"][1]
    assert error_codes(conflict) == ["output_name_conflict"]
    assert "Bereich A/Fall_1.docx" in conflict["errors"][0]["message"]
    assert "[Datei: Bereich B/Tief/Fall_2.docx]" in (output_dir / "conversion.log").read_text(encoding="utf-8")


def test_output_ordner_im_input_ordner_wird_nicht_durchsucht(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    build_lunar_docx(input_dir / "a.docx")
    output_dir = input_dir / "export"

    args = ["convert", "--input-dir", str(input_dir), "--output-dir", str(output_dir)]
    assert main([*args, "--schema-path", str(SCHEMA_PATH), "--config-dir", str(CONFIG_DIR)]) == 0
    assert load_report(output_dir)["summary"]["total"] == 1


def test_dateiauswahl_sortierung_und_temporaere_dateien(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "B.docx")
    build_lunar_docx(input_dir / "a.DOCX")
    build_lunar_docx(input_dir / "~$a.docx")
    (input_dir / "notiz.txt").write_text("x", encoding="utf-8")

    assert run_convert() == 0

    report = load_report(output_dir)
    assert [Path(entry["input_file"]).name for entry in report["files"]] == ["a.DOCX", "B.docx"]
    assert final_entries(output_dir) == ["B", "a"]


def test_conversion_report_inhalt_und_status(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a_ok.docx", DocSpec(steps=[StepSpec(action_images=[PNG]), StepSpec(number="20")]))
    build_lunar_docx(input_dir / "b_ohne_profil.docx", DocSpec(include_testablauf=False))

    assert run_convert() == 1

    report = load_report(output_dir)
    assert set(report) >= {"started_at", "finished_at", "input_dir", "output_dir", "summary", "files"}
    assert report["summary"] == {"total": 2, "success": 1, "failed": 1, "skipped": 0}
    assert report["input_dir"] == str(input_dir)
    assert report["output_dir"] == str(output_dir)

    ok = file_entry(report, "a_ok.docx")
    assert ok["status"] == "success"
    assert ok["profile_detection_status"] == "matched"
    assert ok["detected_profile"] == "lunar_standard_v1"
    matches = {check["profile"]: check["matched"] for check in ok["checked_profiles"]}
    assert matches["lunar_standard_v1"] is True
    assert [profile for profile, matched in matches.items() if matched] == ["lunar_standard_v1"]
    assert ok["output_directory"] == str(output_dir / "a_ok")
    assert ok["step_count"] == 2
    assert ok["exported_image_count"] == 1
    assert ok["errors"] == []

    failed = file_entry(report, "b_ohne_profil.docx")
    assert failed["status"] == "failed"
    assert failed["profile_detection_status"] == "no_matching_profile"
    assert failed["detected_profile"] is None
    assert failed["output_directory"] is None
    assert all(check["matched"] is False for check in failed["checked_profiles"])
    assert all(any("Testablauf" in reason for reason in check["reasons"]) for check in failed["checked_profiles"])
    assert error_codes(failed) == ["no_matching_profile"]


def test_log_enthaelt_datei_und_profil(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action="")]))

    assert run_convert() == 1

    log = (output_dir / "conversion.log").read_text(encoding="utf-8")
    assert "[Datei: a.docx] [Profil: lunar_standard_v1]" in log
    assert "Fehler: required_field_missing" in log
    assert "Feld: steps[0].action" in log


def test_dry_run_schreibt_keine_exportdaten(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=[StepSpec(action_images=[PNG])]))

    assert run_convert("--dry-run") == 0

    assert final_entries(output_dir) == []
    report = load_report(output_dir)
    entry = file_entry(report, "a.docx")
    assert report["dry_run"] is True
    assert entry["status"] == "success"
    assert entry["output_directory"] is None
    assert entry["planned_output_directory"] == str(output_dir / "a")
    assert entry["exported_image_count"] == 1


def test_summary_wird_auch_ohne_dokumentnamen_aus_dateiname_gebildet(
    input_dir: Path, output_dir: Path, run_convert: RunConvert
) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(name=None))

    assert run_convert() == 0

    assert load_testcase(output_dir, "a")["summary"] == "a"
    assert error_codes(file_entry(load_report(output_dir), "a.docx")) == []


def test_ordnername_konflikt_wird_erkannt(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "Test Fall.docx")
    build_lunar_docx(input_dir / "Test_Fall.docx")

    assert run_convert() == 1

    report = load_report(output_dir)
    assert file_entry(report, "Test Fall.docx")["status"] == "success"
    assert error_codes(file_entry(report, "Test_Fall.docx")) == ["output_name_conflict"]
    assert final_entries(output_dir) == ["Test_Fall"]


def test_fehlender_eingabeordner_ist_globaler_fehler(tmp_path: Path, output_dir: Path) -> None:
    from lunar_converter.cli import main
    from tests.helpers import CONFIG_DIR, SCHEMA_PATH

    code = main(
        [
            "convert",
            "--input-dir",
            str(tmp_path / "gibt_es_nicht"),
            "--output-dir",
            str(output_dir),
            "--config-dir",
            str(CONFIG_DIR),
            "--schema-path",
            str(SCHEMA_PATH),
        ]
    )
    assert code == 2
    assert not output_dir.exists()
