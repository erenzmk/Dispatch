"""Tests for the ``canonical_name`` utility."""

from pathlib import Path
import sys

sys.path.append(str(Path(__file__).resolve().parents[1]))

import name_aliases as na


def test_canonical_name_matches_case_insensitively():
    """Exact names should match regardless of case."""

    assert na.canonical_name("ALICE", ["Alice"]) == "Alice"


def test_alias_lookup_is_case_insensitive(monkeypatch):
    """Alias keys should be resolved without regard to case."""

    monkeypatch.setitem(na.ALIASES, "bObBy", "Bob")
    monkeypatch.setitem(na._ALIAS_MAP, "bobby", "Bob")

    assert na.canonical_name("BOBBY", ["Alice", "Bob"]) == "Bob"

