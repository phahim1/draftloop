"""Clause references: normalisation, sorting and scope matching.

Comments refer to clauses in many ways ("Clause 5.2", "5.2 a)", "§5.2",
"Annex B.3", "General"). Everything is reduced to a canonical key such as
"5.2", "B.3" or "general" so comments and drafts can be joined reliably.
"""

from __future__ import annotations

import re

GENERAL = "general"

_ANNEX_RE = re.compile(r"\bannex\s+([A-Z])((?:\.\d{1,3})*)", re.IGNORECASE)
_GENERAL_WORDS = {"", "general", "all", "whole document", "document", "n/a", "na", "-", "—", "ge"}


def normalise_clause(raw: object) -> str:
    """Return a canonical clause key, or ``"general"`` when none is found.

    >>> normalise_clause("Clause 5.2 a)")
    '5.2'
    >>> normalise_clause("Annex b.3")
    'B.3'
    >>> normalise_clause("  ")
    'general'
    """
    if raw is None:
        return GENERAL
    text = str(raw).strip()
    if text.lower() in _GENERAL_WORDS:
        return GENERAL

    annex = _ANNEX_RE.search(text)
    if annex:
        return annex.group(1).upper() + annex.group(2)

    # Drop common prefixes so they cannot be read as annex letters.
    cleaned = re.sub(r"(?i)\b(sub-?clause|clause|cl|section|sec|para(graph)?)\b\.?", " ", text)
    cleaned = cleaned.replace("§", " ")

    # Prefer a purely numeric reference ("5.2"); accept "A.1" style only when
    # it has a dotted part, to avoid reading a stray capital letter as an annex.
    numeric = re.search(r"(?<![0-9.])(\d{1,3}(?:\.\d{1,3})*)(?![0-9])", cleaned)
    if numeric:
        return numeric.group(1).rstrip(".")
    annex_style = re.search(r"\b([A-Z])(\.\d{1,3})+\b", cleaned)
    if annex_style:
        return annex_style.group(0)
    return GENERAL


def sort_key(clause: str) -> tuple:
    """Natural sort: general first, numbered clauses, then annexes A, B, ..."""
    if clause == GENERAL:
        return (0,)
    core = clause.split(" ")[0]          # drop suffixes such as " (new)"
    parts = core.split(".")
    head = parts[0]
    rest = [int(p) for p in parts[1:] if p.isdigit()]
    extra = (1,) if core != clause else (0,)
    if head.isdigit():
        return (1, int(head), *rest, -1, *extra)
    return (2, ord(head[0]), *rest, -1, *extra)


def in_scope(candidate: str, clause: str) -> bool:
    """True when ``candidate`` is ``clause`` or one of its subclauses."""
    return candidate == clause or candidate.startswith(clause + ".")
