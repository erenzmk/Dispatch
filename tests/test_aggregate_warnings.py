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
    wb.close()

    result = aggregate_warnings(tmp_path, [])
    captured = capsys.readouterr()
    assert captured.out == ""
    assert result == Counter()
