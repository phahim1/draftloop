# Adapting DraftLoop to your committee

Standards bodies use similar but not identical procedures. Before using DraftLoop on real work, check the following against your own directives or procedures, and adjust where needed.

| What to check | Where it lives in DraftLoop | Default |
|---|---|---|
| Fields on your comment form | `COMMENT_ALIASES` in `ingest.py`, or a `--column-map` file | Member, clause, type, comment, proposed change, observations |
| Comment types | `TYPE_MAP` in `ingest.py` | general (ge), technical (te), editorial (ed) |
| Decision categories, and which count as "accepted" | `DECISION_MAP` and `ACCEPTING` in `ingest.py` | accepted, accepted in principle, rejected, noted |
| How clauses and annexes are numbered | `clauses.py` | `5.2.1`, `Annex A`, `A.1` |
| Which stage the check is used at | Your process | Between two successive drafts, after comment resolution |
| Whether a disposition of comments must be circulated, and in what form | `cli.py` (`write_disposition_docx`) | Landscape Word table |
| Similarity threshold for grouping comments | `--threshold` | 85 (0 to 100) |

## Confidentiality

Run DraftLoop on your own machine or your organisation's systems. Do not upload committee drafts or comments to external services, and only enable the `--assist` hook with a provider your organisation has approved.

## A sensible pilot

1. Run DraftLoop on a past, closed project and compare its flags with what actually happened.
2. Agree with one convener how the flags will be used, for example as a checklist before a draft is circulated.
3. Measure the time saved and the issues caught, and keep only what helps.
