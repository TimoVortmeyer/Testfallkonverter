"""Profil ``lunar_legacy_v1`` (ältere Vorlage ohne Systemspalte, mit Eingabedaten)."""

from __future__ import annotations

from pathlib import Path

from tests.fixtures.docx_factory import DocSpec, build_legacy_docx, build_lunar_docx
from tests.helpers import RunConvert, file_entry, load_report, load_testcase


def test_aeltere_vorlage_wird_erkannt_und_exportiert(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_legacy_docx(
        input_dir / "alt.docx",
        [
            ("1", "Transaktion SE16, Tabelle SKA1", "", "Buchungskreise W001 und W002", "Konten sind angelegt.\nDaten korrekt"),
            ("2", "Transaktion SE16, Tabelle SKB1", "Feldstatusgruppe", "Buchungskreise W001\nOP-Verwaltung = X", ""),
        ],
        cover_label="FICO-EH",
    )

    assert run_convert() == 0

    entry = file_entry(load_report(output_dir), "alt.docx")
    assert entry["detected_profile"] == "lunar_legacy_v1"
    assert [check["profile"] for check in entry["checked_profiles"] if check["matched"]] == ["lunar_legacy_v1"]

    testcase = load_testcase(output_dir, "alt")
    assert testcase["summary"] == "alt"
    assert testcase["labels"] == ["FICO-EH"]
    assert testcase["description"] == (
        "h1. Initiale Übernahme der Sachkonten\n\n"
        "|Beschreibung|Prüfung der Übernahme|Datum der letzten Änderung|31.01.2011|\n"
        "|Erwartete Ergebnisse|Alle Konten wurden übernommen.| | |"
    )
    assert testcase["custom_fields"] == {
        "customfield_15909": "07.04 Sachkontenpflege/07.04.001 Sachkontenstammdatenpflege/Initiale Übernahme der Sachkonten"
    }
    first, second = testcase["steps"]
    assert first == {
        "system": "nicht definiert",
        "action": "Transaktion SE16, Tabelle SKA1\nBuchungskreise W001 und W002",
        "data": "",
        "expected_result": "Konten sind angelegt.\nDaten korrekt",
        "tester": "",
        "attachments": [],
        "screenshots": [],
    }
    assert second["action"] == "Transaktion SE16, Tabelle SKB1\nBuchungskreise W001\nOP-Verwaltung = X"
    assert second["data"] == ""
    assert second["expected_result"] == "-"
    assert "Feldstatusgruppe" not in str(testcase)


def test_standardvorlage_passt_nicht_zum_legacy_profil(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "neu.docx", DocSpec())

    assert run_convert() == 0

    entry = file_entry(load_report(output_dir), "neu.docx")
    assert entry["detected_profile"] == "lunar_standard_v1"
    assert [check["profile"] for check in entry["checked_profiles"] if check["matched"]] == ["lunar_standard_v1"]
    assert load_testcase(output_dir, "neu")["steps"][0]["data"] == ""
