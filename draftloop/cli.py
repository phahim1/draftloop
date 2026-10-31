"""Command-line entry point.

Example:
    python -m draftloop --comments data/synthetic/comments.xlsx \\
        --decisions data/synthetic/decisions.xlsx \\
        --old data/synthetic/draft_v1.docx --new data/synthetic/draft_v2.docx \\
        --out output
"""

from __future__ import annotations

import argparse
import importlib
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Cm, Pt

from . import __version__
from .compile import DEFAULT_THRESHOLD, compile_comments
from .ingest import apply_decisions, load_comments, load_draft
from .report import summary_counts, write_excel, write_html
from .trace import trace


def write_disposition_docx(path: Path, df, title: str) -> Path:
    """Disposition table as a landscape Word document, ready for a meeting."""
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Cm(1.5))
    doc.add_heading(title, level=1)
    doc.add_paragraph("Comments sorted by clause. Similar comments share a group label. "
                      "The Decision and Observations columns are for the convener.")
    cols = [("comment_id", "ID"), ("member", "Member"), ("clause", "Clause"), ("type", "Type"),
            ("comment", "Comment"), ("proposed_change", "Proposed change"),
            ("similar_group", "Similar"), ("decision", "Decision"), ("rationale", "Observations")]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = "Table Grid"
    for cell, (_, label) in zip(table.rows[0].cells, cols):
        cell.text = label
        cell.paragraphs[0].runs[0].bold = True
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for cell, (key, _) in zip(cells, cols):
            cell.text = str(row.get(key, ""))
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(8)
    doc.save(path)
    return path


def _load_provider(spec: str):
    """Load 'module:function' as an assistance provider."""
    module, _, func = spec.partition(":")
    return getattr(importlib.import_module(module), func)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="draftloop", description="Compile ballot comments and trace them into the next draft.")
    p.add_argument("--comments", required=True, help="Comments file (.xlsx, .csv or .docx table)")
    p.add_argument("--decisions", help="Optional decisions file (ID, decision, rationale)")
    p.add_argument("--old", help="Earlier draft (.docx)")
    p.add_argument("--new", help="Revised draft (.docx)")
    p.add_argument("--out", default="output", help="Output folder")
    p.add_argument("--title", default="Comment resolution and traceability report")
    p.add_argument("--threshold", type=int, default=DEFAULT_THRESHOLD, help="Similarity threshold 0-100")
    p.add_argument("--column-map", help="JSON file mapping fields to your column names")
    p.add_argument("--assist", metavar="MODULE:FUNC",
                   help="Optional advisory provider, e.g. draftloop.assist:keyword_provider (off by default)")
    p.add_argument("--version", action="version", version=f"draftloop {__version__}")
    a = p.parse_args(argv)

    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    cmap = json.loads(Path(a.column_map).read_text()) if a.column_map else None

    comments = load_comments(a.comments, cmap)
    if a.decisions:
        comments = apply_decisions(comments, a.decisions, cmap)
    disposition = compile_comments(comments, a.threshold)
    write_disposition_docx(out / "disposition.docx", disposition, a.title)
    print(f"Compiled {len(disposition)} comments -> {out / 'disposition.docx'}")

    if a.old and a.new:
        traced, register = trace(disposition, load_draft(a.old), load_draft(a.new))
        write_excel(out / "draftloop_report.xlsx", disposition, traced, register, a.title)
        write_html(out / "draftloop_report.html", traced, register, a.title)
        c = summary_counts(traced, register)
        print(f"Traceability: {c.get('NOT_IMPLEMENTED', 0)} accepted-not-implemented, "
              f"{c['UNTRACED_CHANGE']} untraced changes, "
              f"{c.get('CHANGED_DESPITE_REJECTION', 0)} changed-despite-rejection, "
              f"{c.get('NO_DECISION', 0)} without decision")
        print(f"Reports -> {out / 'draftloop_report.xlsx'}, {out / 'draftloop_report.html'}")
    elif a.old or a.new:
        print("Both --old and --new are needed for traceability; skipped.", file=sys.stderr)

    if a.assist:
        from .assist import suggest_themes
        themes = suggest_themes(disposition, _load_provider(a.assist), out / "assist_audit.jsonl", a.assist)
        themes.to_csv(out / "assist_suggestions.csv", index=False)
        print(f"Advisory suggestions -> {out / 'assist_suggestions.csv'} (logged in assist_audit.jsonl)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
