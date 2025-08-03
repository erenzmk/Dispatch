import types
from pathlib import Path
import sys
import pytest
import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import assign_gui


def test_export_reraises_unexpected(tmp_path, monkeypatch, capsys):
    dummy = types.SimpleNamespace(
        mappings={"foo": "bar"},
        liste_path=tmp_path / "Liste.xlsx",
        destroy=lambda: None,
    )
    monkeypatch.setattr(assign_gui, "__file__", str(tmp_path / "assign_gui.py"))

    def boom(*args, **kwargs):
        raise ValueError("boom")

    monkeypatch.setattr(openpyxl, "load_workbook", boom)

    with pytest.raises(ValueError, match="boom"):
        assign_gui.AssignmentApp._export(dummy)

    out = capsys.readouterr().out
    assert "Konnte Liste.xlsx nicht aktualisieren: boom" in out
