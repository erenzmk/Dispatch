import warnings
from pathlib import Path
import pytest
from dispatch import process_reports

@pytest.mark.parametrize("message", process_reports.OPENPYXL_WARNINGS)
def test_safe_load_workbook_suppresses_openpyxl_warnings(tmp_path, monkeypatch, message):
    dummy = tmp_path / "dummy.xlsx"
    dummy.write_text("dummy")

    def fake_load_workbook(filename, *args, **kwargs):
        warnings.warn(message, UserWarning)
        return object()

    monkeypatch.setattr(process_reports, "load_workbook", fake_load_workbook)

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        result = process_reports.safe_load_workbook(dummy)
    assert not caught
    assert result is not None


def test_safe_load_workbook_requires_openpyxl(tmp_path, monkeypatch):
    dummy = tmp_path / "dummy.xlsx"
    dummy.write_text("dummy")
    monkeypatch.setattr(process_reports, "load_workbook", None)
    with pytest.raises(RuntimeError, match="openpyxl is required"):
        process_reports.safe_load_workbook(dummy)
