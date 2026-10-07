"""DOC-Konvertierung über Microsoft Word (PowerShell-Aufruf gemockt, kein Word nötig).

Optional echter Word-Test: Umgebungsvariable ``LUNAR_TEST_WORD_DOC`` auf eine
vorhandene ``.doc``-Datei setzen.
"""

from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path
from typing import Any

import pytest

from lunar_converter import doc_converter
from lunar_converter.cli import main
from lunar_converter.exceptions import ConfigurationError, OutputDirectoryError
from tests.fixtures.docx_factory import DocSpec, build_lunar_docx
from tests.helpers import RunConvert, error_codes, file_entry, final_entries, load_report, load_testcase


@pytest.fixture
def powershell_found(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(doc_converter.shutil, "which", lambda name: r"C:\PS\pwsh.exe" if name == "pwsh" else None)


def _fake_run_factory(source_docx: Path, calls: list[list[str]], *, returncode: int = 0, create_output: bool = True) -> Any:
    def fake_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        assert kwargs["check"] is False and kwargs["capture_output"] is True and kwargs["text"] is True
        assert kwargs["env"]["LUNAR_UTF8_OUTPUT"] == "1"
        in_dir = Path(command[command.index("-InputDir") + 1])
        out_dir = Path(command[command.index("-OutputDir") + 1])
        assert [p.name for p in in_dir.iterdir()] == ["quelle.doc"]
        if create_output:
            out_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy(source_docx, out_dir / "quelle.docx")
        return subprocess.CompletedProcess(command, returncode, stdout="Word ok", stderr="Word-Fehler" if returncode else "")

    return fake_run


@pytest.mark.usefixtures("powershell_found")
def test_doc_wird_mit_word_umgewandelt(
    input_dir: Path,
    output_dir: Path,
    tmp_path: Path,
    run_convert: RunConvert,
    monkeypatch: pytest.MonkeyPatch,
    capsys,
) -> None:
    template = build_lunar_docx(tmp_path / "vorlage.docx", DocSpec(name="TF_aus_DOC"))
    (input_dir / "Alt Format ä.doc").write_bytes(b"\xd0\xcf\x11\xe0 binaeres Word")
    calls: list[list[str]] = []
    monkeypatch.setattr(doc_converter.subprocess, "run", _fake_run_factory(template, calls))

    assert run_convert() == 0
    assert "Konvertiere DOC nach DOCX" in capsys.readouterr().out

    command = calls[0]
    assert command[0] == r"C:\PS\pwsh.exe"
    assert command[1:4] == ["-NoProfile", "-NonInteractive", "-File"]
    assert "ExecutionPolicy" not in " ".join(command)
    assert Path(command[4]).name == "convert_doc_to_docx.ps1" and Path(command[4]).is_file()
    assert load_testcase(output_dir, "Alt_Format_ä")["summary"] == "TF_aus_DOC"
    assert sorted(p.name for p in input_dir.iterdir()) == ["Alt Format ä.doc"]


@pytest.mark.usefixtures("powershell_found")
def test_word_konvertierung_schlaegt_fehl(
    input_dir: Path, output_dir: Path, tmp_path: Path, run_convert: RunConvert, monkeypatch: pytest.MonkeyPatch
) -> None:
    template = build_lunar_docx(tmp_path / "vorlage.docx")
    (input_dir / "a.doc").write_bytes(b"doc")
    build_lunar_docx(input_dir / "b.docx")
    monkeypatch.setattr(doc_converter.subprocess, "run", _fake_run_factory(template, [], returncode=1, create_output=False))

    assert run_convert() == 1

    entry = file_entry(load_report(output_dir), "a.doc")
    assert error_codes(entry) == ["doc_conversion_failed"]
    assert "Word-Fehler" in entry["errors"][0]["details"]
    assert entry["profile_detection_status"] is None
    assert final_entries(output_dir) == ["b"]


@pytest.mark.usefixtures("powershell_found")
def test_word_erzeugt_keine_datei(
    input_dir: Path, output_dir: Path, tmp_path: Path, run_convert: RunConvert, monkeypatch: pytest.MonkeyPatch
) -> None:
    template = build_lunar_docx(tmp_path / "vorlage.docx")
    (input_dir / "a.doc").write_bytes(b"doc")
    monkeypatch.setattr(doc_converter.subprocess, "run", _fake_run_factory(template, [], create_output=False))

    assert run_convert() == 1
    assert error_codes(file_entry(load_report(output_dir), "a.doc")) == ["doc_conversion_failed"]


@pytest.mark.usefixtures("powershell_found")
def test_word_timeout(input_dir: Path, output_dir: Path, run_convert: RunConvert, monkeypatch: pytest.MonkeyPatch) -> None:
    (input_dir / "a.doc").write_bytes(b"doc")

    def raise_timeout(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        raise subprocess.TimeoutExpired(command, 1)

    monkeypatch.setattr(doc_converter.subprocess, "run", raise_timeout)

    assert run_convert() == 1
    assert error_codes(file_entry(load_report(output_dir), "a.doc")) == ["doc_conversion_failed"]


def test_powershell_nicht_verfuegbar(input_dir: Path, output_dir: Path, run_convert: RunConvert, monkeypatch: pytest.MonkeyPatch) -> None:
    (input_dir / "a.doc").write_bytes(b"doc")
    build_lunar_docx(input_dir / "b.docx")
    monkeypatch.setattr(doc_converter.shutil, "which", lambda name: None)

    assert run_convert() == 1

    report = load_report(output_dir)
    assert error_codes(file_entry(report, "a.doc")) == ["doc_conversion_failed"]
    assert file_entry(report, "b.docx")["status"] == "success"


def test_stapelvorbereitung_nutzt_rekursive_progress_optionen(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_dir = tmp_path / "input"
    (input_dir / "Bereich").mkdir(parents=True)
    (input_dir / "a.docx").write_bytes(b"docx")
    (input_dir / "Bereich" / "b.doc").write_bytes(b"doc")
    output_dir = tmp_path / "prepared"
    calls: list[tuple[list[str], dict[str, Any]]] = []
    monkeypatch.setattr(doc_converter.shutil, "which", lambda name: r"C:\PS\pwsh.exe" if name == "pwsh" else None)

    def fake_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append((command, kwargs))
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr(doc_converter.subprocess, "run", fake_run)

    total, result = doc_converter.prepare_docx_directory(input_dir, output_dir)

    assert (total, result) == (2, 0)
    command, kwargs = calls[0]
    assert command[command.index("-InputDir") + 1] == str(input_dir)
    assert command[command.index("-OutputDir") + 1] == str(output_dir)
    assert command[-2:] == ["-Recurse", "-ShowProgress"]
    assert kwargs["capture_output"] is False
    assert output_dir.is_dir()


@pytest.mark.parametrize(
    ("answer", "skip_existing", "target_survives"),
    [
        ("Löschen", False, False),
        ("Nicht Löschen", True, True),
    ],
)
def test_stapelvorbereitung_behandelt_bestehenden_output_ordner(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    answer: str,
    skip_existing: bool,
    target_survives: bool,
) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    (input_dir / "a.doc").write_bytes(b"doc")
    output_dir = tmp_path / "prepared"
    output_dir.mkdir()
    existing_target = output_dir / "a.docx"
    existing_target.write_bytes(b"existing DOCX")
    monkeypatch.setattr(doc_converter.shutil, "which", lambda name: r"C:\PS\pwsh.exe" if name == "pwsh" else None)
    monkeypatch.setattr("builtins.input", lambda _: answer)
    calls: list[list[str]] = []

    def fake_run(command: list[str], **kwargs: Any) -> subprocess.CompletedProcess[str]:
        calls.append(command)
        assert ("-SkipExisting" in command) is skip_existing
        return subprocess.CompletedProcess(command, 0, stdout="", stderr="")

    monkeypatch.setattr(doc_converter.subprocess, "run", fake_run)

    total, result = doc_converter.prepare_docx_directory(input_dir, output_dir)

    assert (total, result) == (1, 0)
    assert len(calls) == 1
    assert existing_target.exists() is target_survives


def test_stapelvorbereitung_bestehender_output_ordner_kann_abgebrochen_werden(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    (input_dir / "a.doc").write_bytes(b"doc")
    output_dir = tmp_path / "prepared"
    output_dir.mkdir()
    existing_target = output_dir / "a.docx"
    existing_target.write_bytes(b"existing DOCX")
    monkeypatch.setattr("builtins.input", lambda _: "Abbrechen")
    monkeypatch.setattr(doc_converter, "find_powershell", lambda: Path(r"C:\PS\pwsh.exe"))
    monkeypatch.setattr(doc_converter.subprocess, "run", lambda *args, **kwargs: pytest.fail("PowerShell darf nicht gestartet werden"))

    with pytest.raises(OutputDirectoryError, match="abgebrochen"):
        doc_converter.prepare_docx_directory(input_dir, output_dir)

    assert existing_target.read_bytes() == b"existing DOCX"


def test_powershell_skip_existing_laesst_zieldocx_unveraendert(tmp_path: Path) -> None:
    if doc_converter.shutil.which("pwsh") is None and doc_converter.shutil.which("powershell") is None:
        pytest.skip("PowerShell nicht verfügbar")
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    (input_dir / "a.doc").write_bytes(b"source DOC; Word must not open this")
    output_dir = tmp_path / "prepared"
    output_dir.mkdir()
    existing_target = output_dir / "a.docx"
    existing_target.write_bytes(b"keep this existing DOCX")
    script = Path(__file__).resolve().parents[2] / "src" / "lunar_converter" / "resources" / doc_converter.WORD_SCRIPT_NAME
    command = doc_converter.build_conversion_command(
        doc_converter.find_powershell(),
        script,
        input_dir,
        output_dir,
        recursive=True,
        skip_existing=True,
    )

    completed = subprocess.run(command, check=False, capture_output=True, text=True, timeout=30)

    assert completed.returncode == 0, completed.stderr
    assert existing_target.read_bytes() == b"keep this existing DOCX"
    assert "uebersprungen" in completed.stdout
    assert "fehlgeschlagen" in completed.stdout and "0 fehlgeschlagen" in completed.stdout


def test_stapelvorbereitung_schuetzt_ueberlappende_ordner(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    input_dir.mkdir()
    (input_dir / "a.docx").write_bytes(b"docx")

    with pytest.raises(ConfigurationError, match="ineinander liegen"):
        doc_converter.prepare_docx_directory(input_dir, input_dir / "prepared")


def test_prepare_docx_mirrors_subdirectories(tmp_path: Path) -> None:
    if doc_converter.shutil.which("pwsh") is None and doc_converter.shutil.which("powershell") is None:
        pytest.skip("PowerShell nicht verfügbar")
    input_dir = tmp_path / "input"
    first = build_lunar_docx(input_dir / "a.docx")
    nested = build_lunar_docx(input_dir / "Bereich" / "b.docx")
    output_dir = tmp_path / "prepared"

    result = main(["prepare-docx", "--input-dir", str(input_dir), "--output-dir", str(output_dir)])

    assert result == 0
    assert (output_dir / "a.docx").read_bytes() == first.read_bytes()
    assert (output_dir / "Bereich" / "b.docx").read_bytes() == nested.read_bytes()


def test_libreoffice_parameter_existiert_nicht(run_convert: RunConvert) -> None:
    with pytest.raises(SystemExit) as info:
        run_convert("--libreoffice-path", "soffice.exe")
    assert info.value.code == 2


@pytest.mark.skipif(not os.environ.get("LUNAR_TEST_WORD_DOC"), reason="LUNAR_TEST_WORD_DOC nicht gesetzt (echter Word-Test)")
def test_echte_word_konvertierung(tmp_path: Path) -> None:
    source = Path(os.environ["LUNAR_TEST_WORD_DOC"])
    result = doc_converter.convert_doc_to_docx(source, tmp_path)
    assert result.is_file() and result.suffix == ".docx"
