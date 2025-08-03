import datetime as dt
from pathlib import Path
import sys

from openpyxl import Workbook, load_workbook
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dispatch.process_reports import update_liste


def test_update_liste(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=3, column=1, value="Bob")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 3, "new": 1, "old": 2}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]

    assert ws2.cell(row=2, column=8).value == 3
    assert ws2.cell(row=2, column=9).value == 3
    assert ws2.cell(row=2, column=10).value == 2
    assert ws2.cell(row=2, column=11).value == 1
    assert ws2.cell(row=3, column=8).value is None
    assert ws2.cell(row=3, column=9).value is None
    assert ws2.cell(row=3, column=10).value is None
    assert ws2.cell(row=3, column=11).value is None


def test_update_liste_empty_morning(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    with pytest.raises(ValueError, match="no data"):
        update_liste(file, "Juli_25", dt.date(2025, 7, 1), {})


def test_update_liste_multiple_runs(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 1, "new": 0, "old": 1}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)
    update_liste(file, "Juli_25", dt.date(2025, 7, 2), morning)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]
    assert ws2.cell(row=2, column=8).value == 1
    wb2.close()
