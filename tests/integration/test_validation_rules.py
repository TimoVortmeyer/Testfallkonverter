"""Fachliche Pflichtfelder je Testfall und Testschritt."""

from __future__ import annotations

from pathlib import Path

import pytest

from tests.fixtures.docx_factory import DocSpec, StepSpec, build_lunar_docx
from tests.helpers import RunConvert, error_codes, file_entry, final_entries, load_report, load_testcase


def test_fehlender_testfallname_wird_abgelehnt(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "ohne_name.docx", DocSpec(name=None))

    assert run_convert() == 1

    entry = file_entry(load_report(output_dir), "ohne_name.docx")
    assert entry["detected_profile"] == "lunar_standard_v1"
    assert entry["errors"][0]["code"] == "required_field_missing"
    assert entry["errors"][0]["field"] == "summary"
    assert final_entries(output_dir) == []


def test_leerer_testfallname_wird_abgelehnt(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "leer.docx", DocSpec(name="   "))

    assert run_convert() == 1
    assert file_entry(load_report(output_dir), "leer.docx")["errors"][0]["field"] == "summary"


@pytest.mark.parametrize(
    ("step_overrides", "field"),
    [
        ({"action": ""}, "action"),
    ],
)
def test_fehlendes_schrittfeld_wird_abgelehnt(
    input_dir: Path, output_dir: Path, run_convert: RunConvert, step_overrides: dict[str, str], field: str
) -> None:
    steps = [StepSpec(), StepSpec(number="20"), StepSpec(number="30", **step_overrides)]
    build_lunar_docx(input_dir / "schritt.docx", DocSpec(steps=steps))

    assert run_convert() == 1

    entry = file_entry(load_report(output_dir), "schritt.docx")
    assert error_codes(entry) == ["required_field_missing"]
    assert entry["errors"][0]["field"] == f"steps[2].{field}"
    assert "3. Testschritt" in entry["errors"][0]["message"]
    assert final_entries(output_dir) == []


def test_leeres_erwartetes_ergebnis_wird_zu_strich(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    steps = [StepSpec(expected=""), StepSpec(number="20", expected="", actual="Nur tatsächlich")]
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=steps))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert testcase["steps"][0]["expected_result"] == "-"
    assert testcase["steps"][1]["expected_result"] == "-\n\nh3. Tatsächliches Ergebnis\nNur tatsächlich"
    assert file_entry(load_report(output_dir), "a.docx")["warnings"] == []


def test_leeres_system_wird_nicht_definiert(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    steps = [StepSpec(), StepSpec(number="", system="", action="Folgeschritt", expected="Ok")]
    build_lunar_docx(input_dir / "a.docx", DocSpec(steps=steps))

    assert run_convert() == 0

    testcase = load_testcase(output_dir, "a")
    assert [step["system"] for step in testcase["steps"]] == ["SAP FI", "nicht definiert"]


def test_alle_fehlenden_felder_werden_gemeldet(input_dir: Path, output_dir: Path, run_convert: RunConvert) -> None:
    build_lunar_docx(input_dir / "a.docx", DocSpec(name=None, steps=[StepSpec(system="", action="", expected="Ergebnis")]))

    assert run_convert() == 1

    fields = [error["field"] for error in file_entry(load_report(output_dir), "a.docx")["errors"]]
    assert fields == ["summary", "steps[0].action"]
