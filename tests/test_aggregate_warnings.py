from collections import Counter
from pathlib import Path
import sys
import datetime as dt

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dispatch.aggregate_warnings import aggregate_warnings


def test_aggregate_warnings_skips_liste(tmp_path, capsys):
    wb = Workbook()
    wb.save(tmp_path / "Liste.xlsx")
    wb.save(tmp_path / "Liste_copy.xlsx")
    wb.close()

    result = aggregate_warnings(tmp_path, [])
    captured = capsys.readouterr()
    assert captured.out == ""
    assert result == Counter()


def test_aggregate_warnings_counts_unknown_names(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Report"
    ws["A2"] = dt.datetime(2025, 7, 1)
    ws["A5"] = "Employee ID"
    ws["B5"] = "Employee Name"
    ws["C5"] = "Open Date Time"
    ws["A6"] = 1
    ws["B6"] = "Bob"
    ws["C6"] = dt.datetime(2025, 6, 30)
    wb.save(tmp_path / "report.xlsx")
    wb.close()

    result = aggregate_warnings(tmp_path, ["Alice"])
    assert result == Counter({"Bob": 1})
