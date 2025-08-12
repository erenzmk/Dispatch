"""Utilities for resolving technician name aliases.

This module contains a small helper :func:`canonical_name` that is used
throughout the project to normalise technician names.  In the real world the
source data is far from consistent – sometimes we only receive the first name,
other times the full name in ``"Lastname, Firstname (Extra)"`` form.  The
previous implementation performed a simple case-insensitive fuzzy match which
meant that entries like ``"Ahmad, Daniyal (Keskin)"`` were not matched to the
existing ``"Daniyal"`` entry in ``Liste.xlsx``.  As a consequence new rows were
added for these verbose names and the daily numbers ended up in the wrong
place.

To make the matching more robust we now strip auxiliary information like
parentheses and leading surnames before performing the fuzzy comparison.  This
mirrors how dispatchers refer to colleagues in the spreadsheet and keeps the
output tidy.
"""

from difflib import get_close_matches
import logging
import re
from typing import Iterable, List, Dict

# Map alternate spellings or abbreviations to their canonical form.
# Extend this mapping as needed. Keys are treated case-insensitively.
ALIASES = {
    "oussama": "Osama",
    "danyal": "Daniyal",
}

# Cached lower-case mapping of aliases to canonical names.  Tests and callers
# may mutate :data:`ALIASES`, so provide a helper to rebuild this cache when
# needed.
_ALIAS_MAP = {k.lower(): v for k, v in ALIASES.items()}

logger = logging.getLogger(__name__)


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

    # ``name`` may come in various formats, e.g. ``"Doe, John (Team)"`` or
    # just ``"john"``.  Normalise by removing parenthetical information and by
    # converting "Lastname, Firstname" into the more common
    # "Firstname Lastname" format used in the spreadsheet.
    norm = name.strip()
    if not norm:
        return norm

    # Drop anything inside parentheses to ignore organisational hints.
    norm = re.sub(r"\([^)]*\)", "", norm)
    candidates: list[str]
    if "," in norm:
        parts = [p.strip() for p in norm.split(",") if p.strip()]
        flipped = " ".join(parts[1:] + parts[:1]) if len(parts) > 1 else parts[0]
        first_only = parts[1] if len(parts) > 1 else parts[0]
        candidates = [flipped, first_only]
    else:
        candidates = [norm]

    valid_map = {v.lower(): v for v in valid_names}
    for cand in candidates:
        key = cand.strip().lower()
        if key in _ALIAS_MAP:
            return _ALIAS_MAP[key]
        matches = get_close_matches(key, list(valid_map.keys()), n=1, cutoff=cutoff)
        if matches:
            return valid_map[matches[0]]

    return candidates[0].strip()


def canonicalize_loaded_names(names: List[str]) -> tuple[List[str], Dict[str, List[int]]]:
    """Kanonisiere *names* und liefere Auftretenspositionen zurück.

    Für jede Position in der Eingabeliste wird der kanonische Name ermittelt.
    Zusätzlich wird eine Zuordnung von kanonischem Namen zu allen Positionen
    aufgebaut, an denen dieser Name vorkommt.  Falls mehrere unterschiedliche
    Zeilen auf denselben Namen fallen, wird eine Warnung protokolliert.
    """

    canonical: List[str] = []
    occurrences: Dict[str, List[int]] = {}
    for idx, name in enumerate(names):
        canon = canonical_name(name, canonical)
        canonical.append(canon)
        occ = occurrences.setdefault(canon, [])
        occ.append(idx)

    for canon, idxs in occurrences.items():
        if len(idxs) > 1:
            logger.warning(
                "Mehrere Zeilen fallen nach Kanonisierung auf %s zusammen: %s",
                canon,
                ", ".join(str(i) for i in idxs),
            )

    return canonical, occurrences
