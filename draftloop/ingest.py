"""Read comments, decisions and drafts.

Column names vary between committees and templates, so columns are matched
by common aliases (case-insensitive). Anything unusual can be mapped
explicitly with ``column_map``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from .clauses import GENERAL, normalise_clause

# Canonical field -> aliases seen in common commenting templates.
COMMENT_ALIASES: dict[str, list[str]] = {
    "comment_id": ["comment id", "id", "ref", "no.", "no", "number", "#"],
    "member": ["member", "mb", "nc", "member body", "national committee", "country", "organisation", "organization", "commenter"],
    "clause": ["clause", "clause/subclause", "subclause", "section", "location"],
    "type": ["type", "type of comment", "comment type", "category"],
    "comment": ["comment", "comments", "comment text", "issue"],
    "proposed_change": ["proposed change", "proposal", "suggested change", "proposed text"],
    "decision": ["decision", "resolution", "disposition", "secretariat decision", "convener decision"],
    "rationale": ["rationale", "observations", "observations of the secretariat", "reason", "response"],
}

DECISION_ALIASES: dict[str, list[str]] = {
    "comment_id": COMMENT_ALIASES["comment_id"],
    "decision": COMMENT_ALIASES["decision"],
    "rationale": COMMENT_ALIASES["rationale"],
}

TYPE_MAP = {"ge": "general", "general": "general", "te": "technical", "technical": "technical",
            "ed": "editorial", "editorial": "editorial"}

# Decision vocabulary. Configure to match your committee's directives.
DECISION_MAP = {
    "accepted": "accepted", "accept": "accepted", "agreed": "accepted", "yes": "accepted",
    "accepted in principle": "accepted in principle", "partly accepted": "accepted in principle",
    "accepted with modification": "accepted in principle", "principle": "accepted in principle",
    "rejected": "rejected", "reject": "rejected", "not accepted": "rejected", "no": "rejected",
    "noted": "noted", "note": "noted", "no action": "noted",
}
ACCEPTING = {"accepted", "accepted in principle"}


def _match_columns(columns: list[str], aliases: dict[str, list[str]],
                   column_map: dict[str, str] | None) -> dict[str, str]:
    """Map canonical field names to the actual column names in a sheet."""
    lookup = {str(c).strip().lower(): c for c in columns}
    found: dict[str, str] = {}
    for canonical, names in aliases.items():
        if column_map and canonical in column_map:
            found[canonical] = column_map[canonical]
            continue
        for name in names:
            if name in lookup:
                found[canonical] = lookup[name]
                break
    return found


def _read_table(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, dtype=str).fillna("")
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path, dtype=str).fillna("")
    if path.suffix.lower() == ".docx":
        return _read_docx_table(path)
    raise ValueError(f"Unsupported file type: {path.suffix}")


def _read_docx_table(path: Path) -> pd.DataFrame:
    """Read the first table of a Word file (many comment templates are tables)."""
    doc = Document(path)
    if not doc.tables:
        raise ValueError(f"No table found in {path}")
    rows = [[cell.text.strip() for cell in row.cells] for row in doc.tables[0].rows]
    return pd.DataFrame(rows[1:], columns=rows[0]).fillna("")


def normalise_decision(raw: str) -> str:
    text = str(raw).strip().lower()
    if not text:
        return ""
    if text in DECISION_MAP:
        return DECISION_MAP[text]
    for key, value in sorted(DECISION_MAP.items(), key=lambda kv: -len(kv[0])):
        if text.startswith(key):
            return value
    return "unrecognised"


def load_comments(path: str | Path, column_map: dict[str, str] | None = None) -> pd.DataFrame:
    """Load member comments into a standard frame.

    Output columns: comment_id, member, clause_raw, clause, type, comment,
    proposed_change, decision, rationale.
    """
    raw = _read_table(path)
    cols = _match_columns(list(raw.columns), COMMENT_ALIASES, column_map)
    missing = [f for f in ("clause", "comment") if f not in cols]
    if missing:
        raise ValueError(f"Could not find column(s) {missing} in {path}. "
                         f"Columns present: {list(raw.columns)}. Use a column map.")

    out = pd.DataFrame()
    out["member"] = raw[cols["member"]].str.strip() if "member" in cols else ""
    out["clause_raw"] = raw[cols["clause"]].astype(str).str.strip()
    out["clause"] = out["clause_raw"].map(normalise_clause)
    out["type"] = (raw[cols["type"]].str.strip().str.lower().map(lambda t: TYPE_MAP.get(t, t))
                   if "type" in cols else "")
    out["comment"] = raw[cols["comment"]].astype(str).str.strip()
    out["proposed_change"] = raw[cols["proposed_change"]].astype(str).str.strip() if "proposed_change" in cols else ""
    out["decision"] = raw[cols["decision"]].map(normalise_decision) if "decision" in cols else ""
    out["rationale"] = raw[cols["rationale"]].astype(str).str.strip() if "rationale" in cols else ""

    if "comment_id" in cols and raw[cols["comment_id"]].str.strip().ne("").all():
        out["comment_id"] = raw[cols["comment_id"]].str.strip()
    else:
        out["comment_id"] = [f"C{i:03d}" for i in range(1, len(raw) + 1)]

    out = out[out["comment"].str.len() > 0].reset_index(drop=True)
    return out[["comment_id", "member", "clause_raw", "clause", "type", "comment",
                "proposed_change", "decision", "rationale"]]


def apply_decisions(comments: pd.DataFrame, path: str | Path,
                    column_map: dict[str, str] | None = None) -> pd.DataFrame:
    """Merge a separate decisions file (comment_id, decision, rationale)."""
    raw = _read_table(path)
    cols = _match_columns(list(raw.columns), DECISION_ALIASES, column_map)
    if "comment_id" not in cols or "decision" not in cols:
        raise ValueError(f"Decisions file needs an ID and a decision column. Found: {list(raw.columns)}")
    dec = pd.DataFrame({
        "comment_id": raw[cols["comment_id"]].str.strip(),
        "decision_new": raw[cols["decision"]].map(normalise_decision),
        "rationale_new": raw[cols["rationale"]].str.strip() if "rationale" in cols else "",
    })
    merged = comments.merge(dec, on="comment_id", how="left")
    merged["decision"] = merged["decision_new"].where(merged["decision_new"].notna() & (merged["decision_new"] != ""),
                                                      merged["decision"])
    merged["rationale"] = merged["rationale_new"].where(merged["rationale_new"].notna() & (merged["rationale_new"] != ""),
                                                        merged["rationale"])
    return merged.drop(columns=["decision_new", "rationale_new"]).fillna("")


# --- Drafts ---------------------------------------------------------------

_HEADING_RE = re.compile(r"^\s*((?:Annex\s+)?(?:[A-Z]|\d{1,3})(?:\.\d{1,3})*)\.?\s+(\S.*)$")


def _iter_blocks(doc):
    """Yield paragraphs and tables in the order they appear in the body."""
    for child in doc.element.body.iterchildren():
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "p":
            yield Paragraph(child, doc)
        elif tag == "tbl":
            yield Table(child, doc)


@dataclass
class Clause:
    key: str
    title: str
    body: list[str] = field(default_factory=list)

    @property
    def text(self) -> str:
        return " ".join([self.title, *self.body]).strip()


def load_draft(path: str | Path) -> dict[str, Clause]:
    """Split a Word draft into clauses keyed by canonical clause number.

    A paragraph starts a new clause when it uses a Heading style, or when it
    begins with a clause number and is short enough to be a title. Text before
    the first numbered heading is kept under "general".
    """
    doc = Document(path)
    clauses: dict[str, Clause] = {}
    current = Clause(GENERAL, "")
    clauses[GENERAL] = current

    # Walk the body in document order so tables stay with their clause.
    for block in _iter_blocks(doc):
        if isinstance(block, Table):
            cells = " | ".join(cell.text.strip() for row in block.rows for cell in row.cells)
            if cells:
                current.body.append(cells)
            continue
        text = block.text.strip()
        if not text:
            continue
        style = (block.style.name or "").lower() if block.style is not None else ""
        match = _HEADING_RE.match(text)
        styled = style.startswith("heading")
        is_heading = match is not None and (styled or len(text) <= 80)
        # A bare capital letter ("A report shall...") is only an annex heading
        # when it says "Annex" or uses a heading style.
        if is_heading and re.fullmatch(r"[A-Z]", match.group(1)) and not styled:
            is_heading = False
        if is_heading:
            key = normalise_clause(match.group(1))
            if key != GENERAL:
                current = Clause(key, match.group(2).strip())
                clauses[key] = current
                continue
        current.body.append(text)

    if not clauses[GENERAL].text:
        clauses.pop(GENERAL)
    return clauses
