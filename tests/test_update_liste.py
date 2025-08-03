import datetime as dt
from pathlib import Path

from openpyxl import Workbook, load_workbook

import name_aliases as na
from process_reports import update_liste


def test_update_liste(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=3, column=1, value="Bob")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    morning = {"Alice": {"total": 3, "new": 1, "old": 2}}
    evening = {"Alice": {"total": 1, "new": 0, "old": 1}}

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning, evening)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]

    assert ws2.cell(row=2, column=8).value == 2
    assert ws2.cell(row=2, column=9).value == 3
    assert ws2.cell(row=2, column=10).value == 2
    assert ws2.cell(row=2, column=11).value == 1


def test_update_liste_handles_variants(tmp_path: Path, monkeypatch):
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli_25"
    ws.cell(row=2, column=1, value="Alice")
    ws.cell(row=3, column=1, value="Bob")
    file = tmp_path / "liste.xlsx"
    wb.save(file)

    monkeypatch.setitem(na.ALIASES, "bobby", "Bob")
    na.refresh_alias_map()

    morning = {
        "ALICE": {"total": 5, "new": 2, "old": 3},
        "bobby": {"total": 2, "new": 1, "old": 1},
    }
    evening = {
        "alice": {"total": 1, "new": 0, "old": 1},
        "BOBBY": {"total": 1, "new": 0, "old": 1},
    }

    update_liste(file, "Juli_25", dt.date(2025, 7, 1), morning, evening)

    wb2 = load_workbook(file)
    ws2 = wb2["Juli_25"]

    assert ws2.cell(row=2, column=1).value == "Alice"
    assert ws2.cell(row=3, column=1).value == "Bob"

    assert ws2.cell(row=2, column=8).value == 4
    assert ws2.cell(row=2, column=9).value == 5
    assert ws2.cell(row=2, column=10).value == 3
    assert ws2.cell(row=2, column=11).value == 2

    assert ws2.cell(row=3, column=8).value == 1
    assert ws2.cell(row=3, column=9).value == 2
    assert ws2.cell(row=3, column=10).value == 1
    assert ws2.cell(row=3, column=11).value == 1

    na.refresh_alias_map()
