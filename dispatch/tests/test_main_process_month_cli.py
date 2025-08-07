from pathlib import Path
import subprocess

import pytest

from dispatch.main import main as cli_main


def test_cli_process_month(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    month_dir = tmp_path / "2025-07"
    liste = tmp_path / "Liste.xlsx"

    called: dict[str, Path] = {}

    def fake_process_month(m_dir, liste_path):
        called["month_dir"] = m_dir
        called["liste"] = liste_path

    monkeypatch.setattr("dispatch.process_reports.process_month", fake_process_month)

    cli_main(["process-month", str(month_dir), str(liste)])

    assert called["month_dir"] == month_dir
    assert called["liste"] == liste


def test_process_month_logs_subprocess_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from run_all_gui import process_month

    logs: list[str] = []

    def fake_log(msg: str) -> None:
        logs.append(msg)

    monkeypatch.setattr("run_all_gui._log", fake_log)

    def fake_run(*args, **kwargs):
        raise subprocess.CalledProcessError(
            returncode=1,
            cmd="cmd",
            output="out",
            stderr="err",
        )

    monkeypatch.setattr("run_all_gui.subprocess.run", fake_run)

    month_dir = tmp_path / "2025-07"
    month_dir.mkdir()
    liste = tmp_path / "Liste.xlsx"
    liste.touch()
    output = tmp_path / "report.csv"

    result = process_month(month_dir, liste, output)

    assert result is False
    assert any("out" in msg and "err" in msg for msg in logs)
