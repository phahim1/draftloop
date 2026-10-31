"""Optional assistance hook, OFF by default.

DraftLoop never calls an AI service on its own. If you choose to plug in a
model (for example a locally hosted one), it can suggest themes for a
clause's comments. Every suggestion is written to an append-only audit log
with a hash of the input, so reviewers can see what was suggested and by
what. Suggestions never change decisions.

A provider is any callable ``provider(clause, comments) -> str``.
The bundled ``keyword_provider`` uses no AI at all; it is a deterministic
example showing the interface.
"""

from __future__ import annotations

import hashlib
import json
import re
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

import pandas as pd

Provider = Callable[[str, list[str]], str]

_STOP = set("""a an the and or of to in on for with by is are be this that it as at from shall should
may not no clause text draft requirement requirements propose proposed change replace add delete
please consider suggest suggests use used using than more less into which""".split())


def keyword_provider(clause: str, comments: list[str]) -> str:
    """Deterministic example: the most frequent content words per clause."""
    words = Counter(w for c in comments for w in re.findall(r"[a-z]{4,}", c.lower()) if w not in _STOP)
    top = ", ".join(w for w, _ in words.most_common(4))
    return f"Recurring terms: {top}" if top else "No recurring terms"


def suggest_themes(comments: pd.DataFrame, provider: Provider, log_path: str | Path,
                   provider_name: str) -> pd.DataFrame:
    """Return one advisory suggestion per clause and log each one."""
    log_path = Path(log_path)
    rows = []
    with log_path.open("a", encoding="utf-8") as log:
        for clause, group in comments.groupby("clause", sort=False):
            texts = group["comment"].tolist()
            suggestion = provider(clause, texts)
            entry = {
                "timestamp": datetime.now(timezone.utc).isoformat(timespec="seconds"),
                "provider": provider_name,
                "clause": clause,
                "comment_ids": group["comment_id"].tolist(),
                "input_sha256": hashlib.sha256("\n".join(texts).encode()).hexdigest(),
                "suggestion": suggestion,
                "status": "advisory - awaiting human review",
            }
            log.write(json.dumps(entry, ensure_ascii=False) + "\n")
            rows.append({"clause": clause, "comments": len(texts), "suggestion": suggestion})
    return pd.DataFrame(rows)
