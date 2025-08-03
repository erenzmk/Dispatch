from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import dispatch.name_aliases as na


def test_alias_lookup_is_case_insensitive(monkeypatch):
    monkeypatch.setitem(na.ALIASES, "bOb", "Robert")
    na.refresh_alias_map()
    assert na.canonical_name("BOB", ["Robert"]) == "Robert"
    na.refresh_alias_map()

def test_removes_parentheses_and_surname():
    valid = ["Daniyal", "Efe"]
    assert na.canonical_name("Ahmad, Daniyal (Keskin)", valid) == "Daniyal"


def test_falls_back_to_norm_when_no_match():
    valid = ["Daniyal"]
    assert na.canonical_name("Unknown", valid) == "Unknown"
