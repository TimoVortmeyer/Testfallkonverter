from __future__ import annotations

import json
from pathlib import Path

import pytest

from lunar_converter.exceptions import ConfigurationError
from lunar_converter.profile_loader import load_profiles, parse_profile
from tests.helpers import CONFIG_DIR


def _minimal() -> dict:
    return {
        "id": "test_v1",
        "name": "Test",
        "required_markers": ["Testablauf"],
        "step_table": {"required_columns": ["Schritt-Nr."]},
        "field_aliases": {"system": ["System"]},
    }


def test_standardprofile_werden_geladen() -> None:
    profiles = load_profiles(CONFIG_DIR)
    ids = [profile.id for profile in profiles]
    assert {"lunar_legacy_v1", "lunar_standard_v1"} <= set(ids)
    assert ids == sorted(ids)
    profile = next(p for p in profiles if p.id == "lunar_standard_v1")
    assert profile.required_markers == ("Testablauf", "Kurzbeschreibung")
    assert "Erwartete Ergebnisse" in profile.required_columns
    assert "Testfall" in profile.aliases_for("testfallname")


@pytest.mark.parametrize("extra_key", ["mapping", "scoring", "python"])
def test_unerlaubte_schluessel_werden_abgelehnt(extra_key: str) -> None:
    data = _minimal()
    data[extra_key] = {}
    with pytest.raises(ConfigurationError):
        parse_profile(data)


def test_unerlaubte_schluessel_in_step_table() -> None:
    data = _minimal()
    data["step_table"]["min_score"] = 3
    with pytest.raises(ConfigurationError):
        parse_profile(data)


def test_unbekannte_alias_schluessel_werden_abgelehnt() -> None:
    data = _minimal()
    data["field_aliases"]["expected_testcase_result"] = ["Erwartetes Ergebnis"]
    with pytest.raises(ConfigurationError, match="expected_testcase_result"):
        parse_profile(data)


def test_info_table_labels() -> None:
    data = _minimal()
    data["info_table"] = {"required_labels": ["Geschäftsvorfall"]}
    assert parse_profile(data).info_table_labels == ("Geschäftsvorfall",)
    data["info_table"] = {"required_labels": [], "mapping": {}}
    with pytest.raises(ConfigurationError):
        parse_profile(data)


def test_leere_pflichtspalten_werden_abgelehnt() -> None:
    data = _minimal()
    data["step_table"]["required_columns"] = []
    with pytest.raises(ConfigurationError):
        parse_profile(data)


def test_doppelte_profil_ids(tmp_path: Path) -> None:
    profiles = tmp_path / "profiles"
    profiles.mkdir()
    for name in ("a.json", "b.json"):
        (profiles / name).write_text(json.dumps(_minimal()), encoding="utf-8")
    with pytest.raises(ConfigurationError, match="doppelt"):
        load_profiles(tmp_path)


def test_fehlendes_profilverzeichnis(tmp_path: Path) -> None:
    with pytest.raises(ConfigurationError):
        load_profiles(tmp_path)
