from pathlib import Path

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


def test_process_month_logs_failure(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from run_dispatch import process_month

    logs: list[tuple[str, object | None]] = []

    def fake_log(msg: str, data: object | None = None) -> None:
        logs.append((msg, data))

    monkeypatch.setattr("run_dispatch._log", fake_log)

    def fake_process_month(m_dir, liste_path):
        raise RuntimeError("kaputt")

    monkeypatch.setattr("run_dispatch.process_reports.process_month", fake_process_month)
    monkeypatch.setattr("run_dispatch.analyze_month.main", lambda args: None)

    month_dir = tmp_path / "2025-07"
    month_dir.mkdir()
    liste = tmp_path / "Liste.xlsx"
    liste.touch()
    output = tmp_path / "report.csv"

    result = process_month(month_dir, liste, output)

    assert result is False
    assert any("kaputt" in msg for msg, _ in logs)


def test_process_month_logs_calls(
    monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    from run_dispatch import process_month

    logs: list[tuple[str, object | None]] = []

    def fake_log(msg: str, data: object | None = None) -> None:
        logs.append((msg, data))

    monkeypatch.setattr("run_dispatch._log", fake_log)

    monkeypatch.setattr("run_dispatch.process_reports.process_month", lambda m, l: None)
    monkeypatch.setattr("run_dispatch.analyze_month.main", lambda args: None)

    def fake_summarize_day(day_dir, liste_path, calls):
        if calls is not None:
            calls["42"] = ["A1", "B2"]
        return True

    monkeypatch.setattr("run_dispatch.summarize_day", fake_summarize_day)

    month_dir = tmp_path / "2025-07"
    day_dir = month_dir / "01"
    day_dir.mkdir(parents=True)
    liste = tmp_path / "Liste.xlsx"
    liste.touch()
    output = tmp_path / "report.csv"

    result = process_month(month_dir, liste, output)

    assert result is True
    assert any(
        isinstance(data, dict) and data.get("01", {}).get("42") == ["A1", "B2"]
        for _, data in logs
    )
