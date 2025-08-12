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


def test_matches_lastname_firstname_combination():
    valid = ["Ammar Alali"]
    assert na.canonical_name("Alali, Ammar (KZEWSKI)", valid) == "Ammar Alali"


def test_canonicalize_loaded_names_warns_on_duplicates(caplog):
    names = ["Oussama", "Osama"]
    with caplog.at_level("WARNING"):
        canon, occ = na.canonicalize_loaded_names(names)
    assert canon == ["Osama", "Osama"]
    assert occ["Osama"] == [0, 1]
    assert "Osama" in caplog.text

