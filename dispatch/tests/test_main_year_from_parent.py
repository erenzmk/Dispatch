import sys
from pathlib import Path
import datetime as dt

import pytest
from openpyxl import Workbook

from dispatch.process_reports import main


def create_liste(path: Path, sheet: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    wb.save(path)


def test_main_uses_year_from_parent(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    base_dir = tmp_path / "2026-07"
    day_dir = base_dir / "01"
    day_dir.mkdir(parents=True)

    liste = tmp_path / "Liste.xlsx"
    create_liste(liste, "Juli_26")

    # Create dummy morning file only
    wb = Workbook()
    wb.save(day_dir / "morning7.xlsx")

    def fake_load_calls(path, valid_names=None):
        return dt.date(2026, 7, 1), {}, []

    called = {}

    def fake_update_liste(
        liste_path,
        month_sheet,
        target_date,
        morning_summary,
        fix_mismatched_dates=False,
    ):
        called["month_sheet"] = month_sheet
        called["target_date"] = target_date

    monkeypatch.setattr(sys, "argv", ["process_reports.py", str(day_dir), str(liste)])
    monkeypatch.setattr("dispatch.process_reports.load_calls", fake_load_calls)
    monkeypatch.setattr("dispatch.process_reports.update_liste", fake_update_liste)

    main()

    assert called["month_sheet"] == "Juli_26"
    assert called["target_date"] == dt.date(2026, 7, 1)
