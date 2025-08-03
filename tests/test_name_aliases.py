import name_aliases as na


def test_alias_lookup_is_case_insensitive(monkeypatch):
    monkeypatch.setitem(na.ALIASES, "bOb", "Robert")
    na.refresh_alias_map()
    assert na.canonical_name("BOB", ["Robert"]) == "Robert"
    na.refresh_alias_map()
