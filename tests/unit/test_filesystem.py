from __future__ import annotations

from pathlib import Path

import pytest

from lunar_converter import filesystem
from lunar_converter.exceptions import ExportError, OutputDirectoryError
from lunar_converter.filesystem import prepare_output_directory, publish_directory, sanitize_folder_name
from lunar_converter.source_discovery import discover_source_files, relative_display_path


@pytest.mark.parametrize(
    ("stem", "expected"),
    [
        ("TFB_03.03.004_EH_012_MDE_Direktbestellung", "TFB_03.03.004_EH_012_MDE_Direktbestellung"),
        ('Test:Fall*1?"<>|', "Test_Fall_1"),
        ("Test  Fall", "Test_Fall"),
        ("CON", "CON_"),
        ("...", "dokument"),
    ],
)
def test_sanitize_folder_name(stem: str, expected: str) -> None:
    assert sanitize_folder_name(stem) == expected


def _staged(tmp_path: Path) -> Path:
    staged = tmp_path / "staged"
    staged.mkdir()
    (staged / "testcase.json").write_text("{}", encoding="utf-8")
    return staged


def test_publish_directory_wiederholt_umbenennen_bei_kurzer_sperre(tmp_path: Path, monkeypatch) -> None:
    real_rename = filesystem.os.rename
    calls = {"count": 0}

    def flaky_rename(source, target):
        calls["count"] += 1
        if calls["count"] < 3:
            raise PermissionError(5, "Zugriff verweigert")
        real_rename(source, target)

    monkeypatch.setattr(filesystem.os, "rename", flaky_rename)
    monkeypatch.setattr(filesystem.time, "sleep", lambda _: None)
    final = tmp_path / "out" / "fall"
    final.parent.mkdir()

    publish_directory(_staged(tmp_path), final)

    assert (final / "testcase.json").is_file()
    assert calls["count"] == 3
    assert [p.name for p in final.parent.iterdir()] == ["fall"]


def test_publish_directory_meldet_dauerhafte_sperre_und_raeumt_auf(tmp_path: Path, monkeypatch) -> None:
    def locked(source, target):
        raise PermissionError(5, "Zugriff verweigert")

    monkeypatch.setattr(filesystem.os, "rename", locked)
    monkeypatch.setattr(filesystem.time, "sleep", lambda _: None)
    final = tmp_path / "out" / "fall"
    final.parent.mkdir()

    with pytest.raises(ExportError) as info:
        publish_directory(_staged(tmp_path), final)

    assert info.value.code == "export_failed"
    assert list(final.parent.iterdir()) == []


def test_prepare_output_directory_legt_an_und_bricht_bei_abgelehnter_loeschung_ab(tmp_path: Path, monkeypatch) -> None:
    target = tmp_path / "neu" / "out"
    prepare_output_directory(target)
    assert target.is_dir()
    prepare_output_directory(target)
    (target / "x").mkdir()
    monkeypatch.setattr("builtins.input", lambda _: "nein")
    with pytest.raises(OutputDirectoryError):
        prepare_output_directory(target)
    assert (target / "x").is_dir()


def test_prepare_output_directory_loescht_nach_bestaetigung(tmp_path: Path, monkeypatch) -> None:
    target = tmp_path / "out"
    (target / "nested").mkdir(parents=True)
    (target / "nested" / "old.txt").write_text("old", encoding="utf-8")
    monkeypatch.setattr("builtins.input", lambda _: "ja")

    prepare_output_directory(target)

    assert target.is_dir()
    assert list(target.iterdir()) == []


def test_prepare_output_directory_schuetzt_eingabeordner(tmp_path: Path, monkeypatch) -> None:
    target = tmp_path / "input"
    source = target / "source.docx"
    target.mkdir()
    source.write_text("source", encoding="utf-8")
    monkeypatch.setattr("builtins.input", lambda _: pytest.fail("Es darf keine Löschbestätigung abgefragt werden."))

    with pytest.raises(OutputDirectoryError, match="Eingabeordner"):
        prepare_output_directory(target, protected_dir=target)
    assert source.is_file()


def test_prepare_output_directory_datei_statt_ordner(tmp_path: Path) -> None:
    file_path = tmp_path / "datei"
    file_path.write_text("x", encoding="utf-8")
    with pytest.raises(OutputDirectoryError):
        prepare_output_directory(file_path)


def test_publish_directory_ueberschreibt_nicht(tmp_path: Path) -> None:
    staged = tmp_path / "staged"
    (staged / "screenshots").mkdir(parents=True)
    (staged / "testcase.json").write_text("{}", encoding="utf-8")
    final = tmp_path / "out" / "fall"
    final.parent.mkdir()

    publish_directory(staged, final)
    assert (final / "testcase.json").is_file()
    assert (final / "screenshots").is_dir()

    with pytest.raises(Exception, match="existiert bereits"):
        publish_directory(staged, final)
    assert sorted(p.name for p in final.parent.iterdir()) == ["fall"]


def test_discover_source_files(tmp_path: Path) -> None:
    for name in ("b.doc", "A.docx", "~$A.docx", "c.pdf", "a2.DOC"):
        (tmp_path / name).write_bytes(b"")
    (tmp_path / "ordner.docx").mkdir()
    assert [p.name for p in discover_source_files(tmp_path)] == ["A.docx", "a2.DOC", "b.doc"]


def test_discover_source_files_in_unterordnern(tmp_path: Path) -> None:
    for relative in ("z.docx", "Sub/b.docx", "Sub/Tiefer/a.doc", "sub2/~$x.docx", "a.docx"):
        (tmp_path / relative).parent.mkdir(parents=True, exist_ok=True)
        (tmp_path / relative).write_bytes(b"")
    output = tmp_path / "out"
    (output / "alt").mkdir(parents=True)
    (output / "alt" / "export.docx").write_bytes(b"")

    found = discover_source_files(tmp_path, exclude_dir=output)

    assert [relative_display_path(p, tmp_path) for p in found] == ["a.docx", "Sub/b.docx", "Sub/Tiefer/a.doc", "z.docx"]
