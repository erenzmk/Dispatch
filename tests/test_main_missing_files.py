import sys
from pathlib import Path

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dispatch.process_reports import main


def create_liste(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    wb.save(path)


def test_main_missing_morning_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    liste = tmp_path / "Liste.xlsx"
    create_liste(liste)

    day_dir = tmp_path / "01.07"
    day_dir.mkdir()

    # create an evening file to ensure only morning is missing
    wb = Workbook()
    wb.save(day_dir / "evening19.xlsx")

    monkeypatch.setattr(sys, "argv", ["process_reports.py", str(day_dir), str(liste)])
    with pytest.raises(FileNotFoundError) as excinfo:
        main()
    assert "Morning report" in str(excinfo.value)
    assert str(day_dir) in str(excinfo.value)


def test_main_missing_evening_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    liste = tmp_path / "Liste.xlsx"
    create_liste(liste)

    day_dir = tmp_path / "01.07"
    day_dir.mkdir()

    # create a morning file but no evening file
    wb = Workbook()
    wb.save(day_dir / "morning7.xlsx")

    monkeypatch.setattr(sys, "argv", ["process_reports.py", str(day_dir), str(liste)])
    with pytest.raises(FileNotFoundError) as excinfo:
        main()
    assert "Evening report" in str(excinfo.value)
    assert str(day_dir) in str(excinfo.value)
