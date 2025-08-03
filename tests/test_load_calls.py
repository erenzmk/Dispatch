import datetime as dt
from pathlib import Path
import sys

import pytest
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from process_reports import load_calls


def test_load_calls_selects_correct_sheet(tmp_path):
    wb = Workbook()
    ws1 = wb.active
    ws1["A1"] = "not the sheet"

    ws2 = wb.create_sheet("Report")
    # Place the date somewhere within the first few rows but not in A2 to
    # ensure ``load_calls`` searches for it dynamically.
    ws2["C3"] = dt.datetime(2025, 7, 1)
    ws2["A5"] = "Employee ID"
    ws2["B5"] = "Employee Name"
    ws2["C5"] = "Open Date Time"
    ws2["A6"] = 1
    ws2["B6"] = "Alice"
    ws2["C6"] = dt.datetime(2025, 6, 30)

    path = tmp_path / "report.xlsx"
    wb.save(path)

    target_date, summary = load_calls(path)

    assert target_date == dt.date(2025, 7, 1)
    assert summary == {"Alice": {"total": 1, "new": 1, "old": 0}}


def test_load_calls_requires_date(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws["A5"] = "Employee ID"
    ws["B5"] = "Employee Name"
    ws["C5"] = "Open Date Time"
    ws["A6"] = 1
    ws["B6"] = "Alice"
    ws["C6"] = dt.datetime(2025, 6, 30)

    path = tmp_path / "report.xlsx"
    wb.save(path)

    with pytest.raises(ValueError, match="Date not found"):
        load_calls(path)
