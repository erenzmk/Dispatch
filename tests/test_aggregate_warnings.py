from collections import Counter
from pathlib import Path
import sys

from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from dispatch.aggregate_warnings import aggregate_warnings


def test_aggregate_warnings_skips_liste(tmp_path, capsys):
    wb = Workbook()
    wb.save(tmp_path / "Liste.xlsx")
    wb.save(tmp_path / "Liste_copy.xlsx")
    wb.save(tmp_path / "Listen.xlsx")
    wb.close()

    result = aggregate_warnings(tmp_path, [])
    captured = capsys.readouterr()
    assert "Liste.xlsx" not in captured.out
    assert "Liste_copy.xlsx" not in captured.out
    assert "Listen.xlsx" in captured.out
    assert "Header row not found in report" in captured.out
    assert result == Counter()
