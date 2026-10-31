"""Closed-loop traceability between two drafts.

Checks, clause by clause, that:
  * every accepted comment is reflected by a change in the next draft;
  * every change in the next draft can be traced to at least one comment;
  * rejected or noted comments did not lead to an unexplained change.

The checks are deliberately conservative. A flag means "a person should
look", not "something is wrong".
"""

from __future__ import annotations

import difflib
import re
from dataclasses import dataclass

import pandas as pd
from rapidfuzz import fuzz

from .clauses import GENERAL, in_scope, sort_key
from .ingest import ACCEPTING, Clause

# Status codes and their meaning. Severity drives colour in the reports.
STATUS = {
    "NOT_IMPLEMENTED": ("high", "Accepted, but the clause is unchanged in the new draft."),
    "CLAUSE_REMOVED": ("medium", "Clause no longer exists in the new draft; confirm this was intended."),
    "CHANGED_DESPITE_REJECTION": ("medium", "Rejected or noted, but the clause changed and no accepted comment explains it."),
    "NO_DECISION": ("medium", "No decision recorded yet."),
    "UNKNOWN_CLAUSE": ("medium", "Clause not found in the earlier draft; check the reference."),
    "IMPLEMENTED": ("ok", "Accepted and the clause changed; confirm the change matches the decision."),
    "NOT_TRACEABLE": ("info", "General comment; cannot be checked automatically."),
    "NO_CHANGE_EXPECTED": ("ok", "Rejected or noted, and the clause is unchanged."),
}

_WS = re.compile(r"\s+")


def _norm(text: str) -> str:
    return _WS.sub(" ", text).strip().lower()


@dataclass
class ClauseChange:
    clause: str          # clause number in the earlier draft ("5.3 (new)" for added clauses)
    title: str
    status: str          # unchanged | modified | added | removed
    similarity: float    # 0-100, old vs new text
    diff: str            # short word-level summary
    new_number: str = "" # clause number in the revised draft, when it differs

    @property
    def text_changed(self) -> bool:
        return self.status in ("modified", "added", "removed")


KEEP_SAME_NUMBER = 60  # same number, at least this similar: treated as the same clause
RENUMBER_MATCH = 85    # different number, at least this similar: treated as renumbered

NEW_SUFFIX = " (new)"


def align(old: dict[str, Clause], new: dict[str, Clause]) -> dict[str, tuple[str, float]]:
    """Match clauses of the earlier draft to clauses of the revised draft.

    Returns {old number: (new number, similarity)}. A clause keeps its number
    when the text is still broadly the same; otherwise it is matched to the
    most similar unmatched clause, so a renumbered clause is followed rather
    than reported as deleted and re-added.
    """
    def sim(a: Clause, b: Clause) -> float:
        return 100.0 if _norm(a.text) == _norm(b.text) else fuzz.ratio(_norm(a.text), _norm(b.text))

    pairs: dict[str, tuple[str, float]] = {}
    for key in old:
        if key in new:
            score = sim(old[key], new[key])
            if key == GENERAL or score >= KEEP_SAME_NUMBER:
                pairs[key] = (key, score)

    taken = {n for n, _ in pairs.values()}
    candidates = sorted(
        ((sim(old[o], new[n]), o, n) for o in old if o not in pairs
         for n in new if n not in taken and n != GENERAL),
        reverse=True)
    for score, o, n in candidates:
        if score < RENUMBER_MATCH:
            break
        if o in pairs or n in taken:
            continue
        pairs[o] = (n, score)
        taken.add(n)

    # Same number but very different text, and no better match elsewhere:
    # treat as a rewrite of the same clause.
    for key in old:
        if key not in pairs and key in new and key not in taken:
            pairs[key] = (key, sim(old[key], new[key]))
            taken.add(key)
    return pairs


def _word_diff(old: str, new: str, limit: int = 300, context: int = 2) -> str:
    """Compact 'old → new' snippets, with a little context, one per change."""
    a, b = old.split(), new.split()
    snippets = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(a=a, b=b, autojunk=False).get_opcodes():
        if op == "equal":
            continue
        before = " ".join(a[max(0, i1 - context):i1])
        after = " ".join(a[i2:i2 + context])
        was = " ".join(a[i1:i2]) or "∅"
        now = " ".join(b[j1:j2]) or "∅"
        snippets.append(f"…{before} [{was} → {now}] {after}…".replace("… ", "…").replace(" …", "…"))
    summary = "  ".join(snippets)
    return summary if len(summary) <= limit else summary[: limit - 1] + "…"


def compare_drafts(old: dict[str, Clause], new: dict[str, Clause]
                   ) -> tuple[list[ClauseChange], dict[str, Clause]]:
    """Compare two drafts clause by clause.

    Returns the list of changes and the revised draft re-keyed by the earlier
    draft's clause numbers (added clauses get a "(new)" suffix).
    """
    pairs = align(old, new)
    matched_new = {n for n, _ in pairs.values()}
    aligned: dict[str, Clause] = {}
    changes = []

    for key in old:
        o = old[key]
        if key not in pairs:
            changes.append(ClauseChange(key, o.title, "removed", 0.0, "(clause deleted)"))
            continue
        new_key, score = pairs[key]
        n = new[new_key]
        aligned[key] = n
        same = _norm(o.text) == _norm(n.text)
        changes.append(ClauseChange(
            key, n.title, "unchanged" if same else "modified", round(score, 1),
            "" if same else _word_diff(o.text, n.text),
            new_number=new_key if new_key != key else ""))

    for key in new:
        if key not in matched_new:
            label = key + NEW_SUFFIX
            aligned[label] = new[key]
            changes.append(ClauseChange(label, new[key].title, "added", 0.0, _word_diff("", new[key].text)))

    changes.sort(key=lambda c: sort_key(c.clause))
    return changes, aligned


def trace(comments: pd.DataFrame, old: dict[str, Clause], new: dict[str, Clause]
          ) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return (comment traceability, clause change register)."""
    changes, new = compare_drafts(old, new)
    changed = {c.clause for c in changes if c.text_changed}
    removed = {c.clause for c in changes if c.status == "removed"}

    def scope_changed(clause: str) -> bool:
        return any(in_scope(k, clause) for k in changed)

    accepted_clauses = set(comments.loc[comments["decision"].isin(ACCEPTING), "clause"])

    rows = []
    for _, c in comments.iterrows():
        clause, decision = c["clause"], c["decision"]
        evidence = ""
        if clause == GENERAL:
            status = "NOT_TRACEABLE"
        elif not any(in_scope(k, clause) for k in old):
            status = "UNKNOWN_CLAUSE"
        elif decision in ("", "unrecognised"):
            status = "NO_DECISION"
        elif clause in removed:
            status = "IMPLEMENTED" if decision in ACCEPTING else "CLAUSE_REMOVED"
            evidence = "clause deleted in new draft"
        elif decision in ACCEPTING:
            status = "IMPLEMENTED" if scope_changed(clause) else "NOT_IMPLEMENTED"
            proposal = str(c.get("proposed_change", "")).strip()
            new_text = " ".join(v.text for k, v in new.items() if in_scope(k, clause))
            if proposal and new_text:
                score = fuzz.partial_ratio(_norm(proposal), _norm(new_text))
                evidence = f"proposed text {'found' if score >= 85 else 'not found'} in new draft ({score:.0f}%)"
        else:  # rejected or noted
            explained = any(in_scope(clause, a) or in_scope(a, clause) for a in accepted_clauses)
            status = "CHANGED_DESPITE_REJECTION" if scope_changed(clause) and not explained else "NO_CHANGE_EXPECTED"
        severity, meaning = STATUS[status]
        rows.append({**c.to_dict(), "trace_status": status, "severity": severity,
                     "meaning": meaning, "evidence": evidence})
    traced = pd.DataFrame(rows)

    # Clause register: flag changes that no comment points to.
    commented = set(comments["clause"]) - {GENERAL}

    def linked(clause: str, pool: set[str]) -> bool:
        return any(in_scope(clause, c) or in_scope(c, clause) for c in pool)

    reg = []
    for ch in changes:
        if ch.status == "unchanged":
            note, severity = "", "ok"
            if ch.new_number:
                note, severity = f"Renumbered to {ch.new_number}; text unchanged", "info"
        elif ch.clause == GENERAL:
            note, severity = "Front matter changed (title, foreword or introduction); not traced", "info"
        elif linked(ch.clause, accepted_clauses - {GENERAL}):
            note, severity = "Linked to an accepted comment", "ok"
        elif linked(ch.clause, commented):
            note, severity = "Only rejected, noted or undecided comments refer to this clause", "medium"
        else:
            note, severity = "UNTRACED_CHANGE: no comment refers to this clause", "medium"
        if ch.new_number and ch.status != "unchanged":
            note = f"Renumbered to {ch.new_number}. {note}"
        reg.append({"clause": ch.clause, "new_number": ch.new_number or ch.clause.replace(NEW_SUFFIX, ""),
                    "title": ch.title, "change": ch.status,
                    "similarity_%": ch.similarity, "trace_note": note, "severity": severity,
                    "what_changed": ch.diff})
    return traced, pd.DataFrame(reg)
