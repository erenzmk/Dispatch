"""Utilities for resolving technician name aliases."""

from difflib import get_close_matches
from typing import Iterable

# Map alternate spellings or abbreviations to their canonical form.
# Extend this mapping as needed. Keys are treated case-insensitively.
ALIASES = {
    # "jon doe": "John Doe",
}

# Cached lower-case mapping of aliases to canonical names.  Tests and callers
# may mutate :data:`ALIASES`, so provide a helper to rebuild this cache when
# needed.
_ALIAS_MAP = {k.lower(): v for k, v in ALIASES.items()}


def refresh_alias_map() -> None:
    """Rebuild the cached alias mapping from :data:`ALIASES`."""
    global _ALIAS_MAP
    _ALIAS_MAP = {k.lower(): v for k, v in ALIASES.items()}


def canonical_name(name: str, valid_names: Iterable[str], cutoff: float = 0.8) -> str:
    """Return the canonical representation for *name*.

    ``valid_names`` should contain the canonical names available in
    ``Liste.xlsx``.  First the static :data:`ALIASES` mapping is checked,
    otherwise :func:`difflib.get_close_matches` is used to find the closest
    match above ``cutoff``. Matching is performed case-insensitively so that
    names like ``"ALICE"`` still resolve to ``"Alice"``. Leading and trailing
    whitespace is removed before matching. If no match is found, the stripped
    name is returned unchanged.
    """

    norm = name.strip()
    if not norm:
        return norm

    key = norm.lower()
    if key in _ALIAS_MAP:
        return _ALIAS_MAP[key]

    valid_map = {v.lower(): v for v in valid_names}
    matches = get_close_matches(key, list(valid_map.keys()), n=1, cutoff=cutoff)
    return valid_map[matches[0]] if matches else norm
