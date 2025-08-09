import datetime as dt
from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from dispatch.process_reports import update_liste, excel_to_date


def test_update_liste(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=2, column=2, value=dt.date(2025, 7, 1))
    ws.cell(row=3, column=1, value="Bob")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 3, "new": 1, "old": 2}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]

    assert ws2.cell(row=2, column=8).value is None
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
    ws.cell(row=2, column=2, value=dt.date(2025, 7, 1))
    ws.cell(row=2, column=15, value=dt.date(2025, 7, 2))
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 1, "new": 0, "old": 1}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)
    update_liste(file, "Juli_25", dt.date(2025, 7, 2), morning)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]
    assert ws2.cell(row=2, column=9).value == 1
    assert ws2.cell(row=2, column=22).value == 1
    wb2.close()


def test_update_liste_uses_matching_date(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=3, column=1, value="Alice")
    ws.cell(row=2, column=2, value=dt.date(2025, 7, 2))
    ws.cell(row=3, column=2, value=dt.date(2025, 7, 1))
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 5, "new": 2, "old": 3}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]
    # Erste Zeile mit falschem Datum bleibt unverändert
    assert excel_to_date(ws2.cell(row=2, column=2).value) == dt.date(2025, 7, 2)
    assert ws2.cell(row=2, column=9).value is None
    # Zweite Zeile wird für den passenden Tag beschrieben
    assert excel_to_date(ws2.cell(row=3, column=2).value) == dt.date(2025, 7, 1)
    assert ws2.cell(row=3, column=9).value == 5
    assert ws2.cell(row=3, column=10).value == 3
    assert ws2.cell(row=3, column=11).value == 2
    wb2.close()


def test_update_liste_creates_missing_sheet(tmp_path: Path):
    file = tmp_path / "liste.xlsx"
    wb = Workbook()
    wb.save(file)

    morning = {"Alice": {"total": 2, "new": 1, "old": 1}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)

    wb2 = load_workbook(file)
    assert "Juli_25" in wb2.sheetnames
    ws = wb2["Juli_25"]
    assert ws.cell(row=1, column=1).value == "Techniker"
    assert ws.max_row == 1
    wb2.close()


def test_excel_to_date_none():
    with pytest.raises(ValueError, match="Leere Zelle"):
        excel_to_date(None)


def test_excel_to_date_invalid():
    with pytest.raises(ValueError, match="Ungültiger Datumswert"):
        excel_to_date("abc")


def test_update_liste_missing_date_cell(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 1, "new": 0, "old": 1}}

    with pytest.raises(ValueError, match="Leere Zelle"):
        update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)


def test_update_liste_invalid_date_cell(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=2, column=2, value="abc")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 1, "new": 0, "old": 1}}

    with pytest.raises(ValueError, match="Ungültiger Datumswert"):
        update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning)
