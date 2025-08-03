"""Utilities for resolving technician name aliases."""

from difflib import get_close_matches
from typing import Iterable

# Map alternate spellings or abbreviations to their canonical form.
# Extend this mapping as needed.
ALIASES = {
    # "Jon Doe": "John Doe",
}


def canonical_name(name: str, valid_names: Iterable[str], cutoff: float = 0.8) -> str:
    """Return the canonical representation for *name*.

    ``valid_names`` should contain the canonical names available in
    ``Liste.xlsx``.  First the static :data:`ALIASES` mapping is checked,
    otherwise :func:`difflib.get_close_matches` is used to find the closest
    match above ``cutoff``. If no match is found, ``name`` is returned
    unchanged.
    """

    norm = name.strip()
    if not norm:
        return norm
    if norm in ALIASES:
        return ALIASES[norm]
    matches = get_close_matches(norm, list(valid_names), n=1, cutoff=cutoff)
    return matches[0] if matches else norm
