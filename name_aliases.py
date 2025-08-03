"""Utilities for resolving technician name aliases."""

from difflib import get_close_matches
from typing import Iterable

# Map alternate spellings or abbreviations to their canonical form.
# Extend this mapping as needed. Keys are treated case-insensitively using
# :py:meth:`str.casefold` for robust comparisons.
ALIASES = {
    # "jon doe": "John Doe",
}


def canonical_name(name: str, valid_names: Iterable[str], cutoff: float = 0.8) -> str:
    """Return the canonical representation for *name*.

    ``valid_names`` should contain the canonical names available in
    ``Liste.xlsx``.  First the static :data:`ALIASES` mapping is checked,
    otherwise :func:`difflib.get_close_matches` is used to find the closest
    match above ``cutoff``. Matching is performed case-insensitively so that
    names like ``"ALICE"`` still resolve to ``"Alice"``. If no match is found,
    ``name`` is returned unchanged.
    """

    norm = name.strip()
    if not norm:
        return norm

    key = norm.casefold()
    alias_map = {k.casefold(): v for k, v in ALIASES.items()}
    if key in alias_map:
        return alias_map[key]

    valid_map = {v.casefold(): v for v in valid_names}
    matches = get_close_matches(key, list(valid_map.keys()), n=1, cutoff=cutoff)
    return valid_map[matches[0]] if matches else norm
