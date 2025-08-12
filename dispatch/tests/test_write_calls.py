import datetime as dt
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from dispatch.write_calls import write_calls


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Juli"
    names = [f"Name{i}" for i in range(1, 27)]
    for idx, name in enumerate(names, start=2):
        ws.cell(row=idx, column=1, value=name)
    # Namen für weitere Tage kopieren
    for day in range(1, 7):
        for idx, name in enumerate(names, start=2 + 26 * day):
            ws.cell(row=idx, column=1, value=name)
    wb.save(path)


def test_write_calls(tmp_path: Path) -> None:
    file = tmp_path / "Juli.xlsx"
    _create_workbook(file)
    records = pd.DataFrame(
        [
            {"name": "Name1", "date": dt.date(2025, 7, 1), "value": 5},
            {"name": "Name2", "date": dt.date(2025, 7, 15), "value": 3},
        ]
    )
    write_calls(file, records)
    wb = load_workbook(file)
    ws = wb["Juli"]
    assert ws["B28"].value == 5
    assert ws["AD29"].value == 3
