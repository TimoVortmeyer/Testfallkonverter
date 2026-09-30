"""Laden und formale Prüfung der deklarativen Dokumentprofile.

Profile dürfen ausschließlich ``id``, ``name``, ``required_markers``,
``step_table.required_columns``, optional ``info_table.required_labels`` und
``field_aliases`` enthalten. Alles andere
wird als Konfigurationsfehler abgewiesen, damit Profile schlank bleiben.
"""

from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .exceptions import ConfigurationError
from .models import ProfileDefinition
from .semantic_model import PROFILE_ALIAS_KEYS

PROFILES_SUBDIR = "profiles"
_ALLOWED_KEYS = frozenset({"id", "name", "required_markers", "step_table", "info_table", "field_aliases"})
_ALLOWED_STEP_TABLE_KEYS = frozenset({"required_columns"})
_ALLOWED_INFO_TABLE_KEYS = frozenset({"required_labels"})
_PROFILE_ID_RE = re.compile(r"^[A-Za-z0-9_.-]+$")


def load_profiles(config_dir: Path) -> list[ProfileDefinition]:
    """Lädt alle Profile aus ``<config_dir>/profiles`` in alphabetischer Dateireihenfolge."""
    profiles_dir = config_dir / PROFILES_SUBDIR
    if not profiles_dir.is_dir():
        raise ConfigurationError(f"Profilverzeichnis '{profiles_dir}' wurde nicht gefunden.")
    files = sorted((path for path in profiles_dir.glob("*.json") if path.is_file()), key=lambda p: p.name.casefold())
    if not files:
        raise ConfigurationError(f"Im Profilverzeichnis '{profiles_dir}' wurden keine Profile (*.json) gefunden.")

    profiles: list[ProfileDefinition] = []
    seen: dict[str, Path] = {}
    for path in files:
        profile = load_profile(path)
        if profile.id in seen:
            raise ConfigurationError(
                f"Profil-ID '{profile.id}' ist doppelt vergeben ('{seen[profile.id].name}' und '{path.name}')."
            )
        seen[profile.id] = path
        profiles.append(profile)
    return profiles


def load_profile(path: Path) -> ProfileDefinition:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ConfigurationError(f"Profildatei '{path}' ist nicht lesbar oder kein gültiges JSON.", details=str(exc)) from exc
    return parse_profile(data, source=path)


def parse_profile(data: Any, source: Path | None = None) -> ProfileDefinition:
    label = f"Profil '{source.name}'" if source else "Profil"
    if not isinstance(data, dict):
        raise ConfigurationError(f"{label}: Wurzelelement muss ein JSON-Objekt sein.")
    unknown = sorted(set(data) - _ALLOWED_KEYS)
    if unknown:
        raise ConfigurationError(f"{label}: Nicht erlaubte Schlüssel {unknown}. Erlaubt sind nur {sorted(_ALLOWED_KEYS)}.")

    profile_id = data.get("id")
    if not isinstance(profile_id, str) or not _PROFILE_ID_RE.match(profile_id):
        raise ConfigurationError(f"{label}: 'id' fehlt oder enthält ungültige Zeichen (erlaubt: A-Z, a-z, 0-9, _ . -).")
    name = data.get("name")
    if not isinstance(name, str) or not name.strip():
        raise ConfigurationError(f"{label}: 'name' fehlt oder ist leer.")

    markers = _string_list(data.get("required_markers", []), f"{label}: 'required_markers'")

    step_table = data.get("step_table")
    if not isinstance(step_table, dict):
        raise ConfigurationError(f"{label}: 'step_table' fehlt oder ist kein Objekt.")
    unknown_step = sorted(set(step_table) - _ALLOWED_STEP_TABLE_KEYS)
    if unknown_step:
        raise ConfigurationError(f"{label}: Nicht erlaubte Schlüssel in 'step_table': {unknown_step}.")
    columns = _string_list(step_table.get("required_columns"), f"{label}: 'step_table.required_columns'")
    if not columns:
        raise ConfigurationError(f"{label}: 'step_table.required_columns' darf nicht leer sein.")

    info_labels: list[str] = []
    info_table = data.get("info_table")
    if info_table is not None:
        if not isinstance(info_table, dict) or set(info_table) - _ALLOWED_INFO_TABLE_KEYS:
            raise ConfigurationError(f"{label}: 'info_table' darf nur 'required_labels' enthalten.")
        info_labels = _string_list(info_table.get("required_labels"), f"{label}: 'info_table.required_labels'")
        if not info_labels:
            raise ConfigurationError(f"{label}: 'info_table.required_labels' darf nicht leer sein.")

    raw_aliases = data.get("field_aliases", {})
    if not isinstance(raw_aliases, dict):
        raise ConfigurationError(f"{label}: 'field_aliases' muss ein Objekt sein.")
    unknown_aliases = sorted(set(raw_aliases) - set(PROFILE_ALIAS_KEYS))
    if unknown_aliases:
        raise ConfigurationError(
            f"{label}: Unbekannte Schlüssel in 'field_aliases': {unknown_aliases}. "
            f"Erlaubt sind nur {list(PROFILE_ALIAS_KEYS)}."
        )
    aliases: dict[str, tuple[str, ...]] = {}
    for key, values in raw_aliases.items():
        aliases[str(key)] = tuple(_string_list(values, f"{label}: 'field_aliases.{key}'"))

    return ProfileDefinition(
        id=profile_id,
        name=name.strip(),
        required_markers=tuple(markers),
        required_columns=tuple(columns),
        field_aliases=aliases,
        info_table_labels=tuple(info_labels),
        source_path=source,
    )


def _string_list(value: Any, label: str) -> list[str]:
    if not isinstance(value, list) or not all(isinstance(item, str) and item.strip() for item in value):
        raise ConfigurationError(f"{label} muss eine Liste nicht leerer Zeichenketten sein.")
    return [item.strip() for item in value]
