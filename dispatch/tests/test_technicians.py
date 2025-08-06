from pathlib import Path

from openpyxl import Workbook

from dispatch.technicians import load_id_map


def test_load_id_map(tmp_path: Path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ID", "Techniker"])
    ws.append([" 1 ", " Alice "])
    ws.append(["2", "Bob"])
    ws.append([None, "Charlie"])
    ws.append(["3", None])
    ws.append([None, None])
    file = tmp_path / "Liste.xlsx"
    wb.save(file)
    wb.close()

    assert load_id_map(file) == {"1": "Alice", "2": "Bob"}
