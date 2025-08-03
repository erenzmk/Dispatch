import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from main import main as cli_main


def test_cli_process_month(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    month_dir = tmp_path / "Juli_25"
    liste = tmp_path / "Liste.xlsx"

    called: dict[str, Path] = {}

    def fake_process_month(m_dir, liste_path):
        called["month_dir"] = m_dir
        called["liste"] = liste_path

    monkeypatch.setattr("dispatch.process_reports.process_month", fake_process_month)

    cli_main(["process-month", str(month_dir), str(liste)])

    assert called["month_dir"] == month_dir
    assert called["liste"] == liste
