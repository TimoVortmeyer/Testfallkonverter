from __future__ import annotations

from pathlib import Path

import pytest

from lunar_converter.exceptions import OutputDirectoryError
from lunar_converter.filesystem import prepare_output_directory, publish_directory, sanitize_folder_name
from lunar_converter.source_discovery import discover_source_files


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


def test_prepare_output_directory_legt_an_und_prueft_leer(tmp_path: Path) -> None:
    target = tmp_path / "neu" / "out"
    prepare_output_directory(target)
    assert target.is_dir()
    prepare_output_directory(target)
    (target / "x").mkdir()
    with pytest.raises(OutputDirectoryError):
        prepare_output_directory(target)


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
