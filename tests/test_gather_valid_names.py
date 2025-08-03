import sys
from pathlib import Path
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dispatch.aggregate_warnings import gather_valid_names


def test_gather_valid_names_reads_sheet_and_deduplicates(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Technikernamen"
    ws.append(["Technikername", "PUOOS"])
    ws.append(["Alice", "Ali"])
    ws.append(["Bob", None])
    ws.append(["Alice", ""])
    wb.save(tmp_path / "Liste.xlsx")
    wb.close()

    names = gather_valid_names(tmp_path / "Liste.xlsx")
    assert names == ["Ali", "Alice", "Bob"]
moke7a-codex/fix-duplicate-technician-names-display


def test_gather_valid_names_autodetects_sheet(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Technikernamen + PUDO"
    ws.append(["Technikername"])
    ws.append(["Eva"])
    wb.save(tmp_path / "Liste.xlsx")
    wb.close()

    assert gather_valid_names(tmp_path / "Liste.xlsx") == ["Eva"]
=======
main
