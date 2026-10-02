"""Profil ``gh_standard_v1`` (GH-Vorlage mit Geschäftsvorfall, Varianten und ohne Systemspalte)."""

from __future__ import annotations

from pathlib import Path

from tests.fixtures.docx_factory import build_gh_docx
from tests.helpers import RunConvert, file_entry, load_report, load_testcase


def test_gh_vorlage_wird_erkannt_und_exportiert(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_gh_docx(
        input_dir / "gh.docx",
        [
            ("1", "", "MEK1 - EK-Konditionen anlegen", "", "", ""),
            ("1.1", "0001", "Einkaufskondition anlegen", "Lieferant", "Einkaufsorganisation 0010", "Eintrag wird angelegt."),
        ],
    )

    assert run_convert() == 0

    entry = file_entry(load_report(output_dir), "gh.docx")
    assert entry["detected_profile"] == "gh_standard_v1"
    assert [check["profile"] for check in entry["checked_profiles"] if check["matched"]] == ["gh_standard_v1"]
    assert entry["warnings"] == []

    testcase = load_testcase(output_dir, "gh")
    assert testcase["summary"] == "gh"
    assert testcase["labels"] == ["RWWS-GH", "Erika Muster"]
    assert testcase["custom_fields"] == {
        "customfield_15909": "03.02 Einkaufsverwaltung/03.02.001 GH Pflege Einkaufskonditionen (EGKE)/MEK1 - EK-Konditionen anlegen"
    }
    # Info-Tabelle ist die Tabelle mit "Geschäftsvorfall", nicht die letzte Tabelle vor dem Testablauf.
    assert testcase["description"] == (
        "h1. MEK1 - EK-Konditionen anlegen\n\n"
        "|Geschäftsvorfall|MEK1 - EK-Konditionen anlegen|Verantwortlicher Tester| |\n"
        "|Beschreibung| | | |"
    )
    first, second = testcase["steps"]
    assert first["system"] == "nicht definiert"
    assert first["action"] == "MEK1 - EK-Konditionen anlegen"
    assert first["expected_result"] == "-"
    assert second["action"] == "Einkaufskondition anlegen\nEinkaufsorganisation 0010"
    assert second["data"] == ""
    assert second["expected_result"] == "Eintrag wird angelegt."
    assert "Lieferant" not in str(testcase) and "0001" not in str(testcase["steps"])
