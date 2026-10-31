"""Compile comments into a clause-ordered disposition table.

Duplicate and near-duplicate comments (the same point raised by several
members) are grouped so the convener can resolve them together. Grouping is
a suggestion only; nothing is merged or removed.
"""

from __future__ import annotations

import pandas as pd
from rapidfuzz import fuzz, process

from .clauses import sort_key

DEFAULT_THRESHOLD = 85  # token-set similarity, 0-100


def _groups(texts: list[str], threshold: int) -> list[int]:
    """Union-find over pairs whose similarity meets the threshold."""
    n = len(texts)
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    if n > 1:
        scores = process.cdist(texts, texts, scorer=fuzz.token_set_ratio, score_cutoff=threshold)
        for i in range(n):
            for j in range(i + 1, n):
                if scores[i][j] >= threshold:
                    parent[find(j)] = find(i)
    return [find(i) for i in range(n)]


def compile_comments(comments: pd.DataFrame, threshold: int = DEFAULT_THRESHOLD) -> pd.DataFrame:
    """Sort comments by clause and flag similar comments.

    Adds ``similar_group`` (e.g. "S01", empty when unique) and
    ``similar_to`` (IDs of the other comments in the group).
    """
    df = comments.copy()
    df["_sort"] = df["clause"].map(sort_key)
    df = df.sort_values(["_sort", "member", "comment_id"], kind="stable").drop(columns="_sort")
    df = df.reset_index(drop=True)

    texts = (df["comment"].str.lower() + " " + df["proposed_change"].str.lower()).tolist()
    roots = _groups(texts, threshold)
    members: dict[int, list[int]] = {}
    for idx, root in enumerate(roots):
        members.setdefault(root, []).append(idx)

    label: dict[int, str] = {}
    counter = 0
    for root, idxs in sorted(members.items(), key=lambda kv: min(kv[1])):
        if len(idxs) > 1:
            counter += 1
            for i in idxs:
                label[i] = f"S{counter:02d}"

    df["similar_group"] = [label.get(i, "") for i in range(len(df))]
    ids = df["comment_id"].tolist()
    df["similar_to"] = [
        ", ".join(ids[j] for j in members[roots[i]] if j != i) if label.get(i) else ""
        for i in range(len(df))
    ]
    return df
