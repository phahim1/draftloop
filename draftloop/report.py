"""Write the Excel workbook and the HTML summary."""

from __future__ import annotations

import html
from datetime import date
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

FILL = {
    "high": PatternFill("solid", fgColor="F8D7DA"),
    "medium": PatternFill("solid", fgColor="FFF3CD"),
    "ok": PatternFill("solid", fgColor="D1E7DD"),
    "info": PatternFill("solid", fgColor="E2E3E5"),
}
HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F3A5F")

DISPOSITION_COLS = {
    "comment_id": "ID", "member": "Member", "clause": "Clause", "type": "Type",
    "comment": "Comment", "proposed_change": "Proposed change",
    "similar_group": "Similar group", "similar_to": "Similar to",
    "decision": "Decision", "rationale": "Rationale / observations",
}
TRACE_COLS = {
    "comment_id": "ID", "member": "Member", "clause": "Clause", "decision": "Decision",
    "trace_status": "Trace status", "meaning": "Meaning", "evidence": "Evidence",
    "comment": "Comment",
}
REGISTER_COLS = {
    "clause": "Clause", "new_number": "New number", "title": "Title", "change": "Change", "similarity_%": "Similarity %",
    "trace_note": "Trace note", "what_changed": "What changed (word level)",
}
WIDTHS = {"Comment": 60, "Proposed change": 45, "Rationale / observations": 40, "Meaning": 45,
          "What changed (word level)": 70, "Evidence": 35, "Trace note": 40, "Title": 30}


def summary_counts(traced: pd.DataFrame, register: pd.DataFrame) -> dict[str, int]:
    counts = traced["trace_status"].value_counts().to_dict()
    counts["UNTRACED_CHANGE"] = int(register["trace_note"].str.startswith("UNTRACED").sum())
    counts["comments"] = len(traced)
    counts["clauses_changed"] = int(((register["change"] != "unchanged") & (register["clause"] != "general")).sum())
    counts["renumbered"] = int((register["new_number"] != register["clause"].str.replace(" (new)", "", regex=False)).sum())
    counts["similar_groups"] = int(traced.get("similar_group", pd.Series(dtype=str)).replace("", pd.NA).nunique())
    return counts


def _sheet(wb: Workbook, title: str, df: pd.DataFrame, cols: dict[str, str],
           severity: pd.Series | None = None) -> None:
    ws = wb.create_sheet(title)
    data = df[[c for c in cols if c in df.columns]].rename(columns=cols)
    ws.append(list(data.columns))
    for row in data.itertuples(index=False):
        ws.append(list(row))
    for cell in ws[1]:
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
    for idx, name in enumerate(data.columns, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = WIDTHS.get(name, 14)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    if severity is not None:
        for r, sev in enumerate(severity.tolist(), start=2):
            for cell in ws[r]:
                cell.fill = FILL.get(sev, FILL["info"])
    if len(data):
        ref = f"A1:{get_column_letter(len(data.columns))}{len(data) + 1}"
        safe = "".join(ch for ch in title if ch.isalnum()) or "Data"
        ws.add_table(Table(displayName=safe, ref=ref,
                           tableStyleInfo=TableStyleInfo(name="TableStyleLight1", showRowStripes=False)))
    ws.freeze_panes = "A2"


def write_excel(path: str | Path, disposition: pd.DataFrame, traced: pd.DataFrame,
                register: pd.DataFrame, title: str) -> Path:
    path = Path(path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    counts = summary_counts(traced, register)
    rows = [
        (title, ""), (f"Generated {date.today().isoformat()} by DraftLoop (prototype)", ""), ("", ""),
        ("Comments received", counts["comments"]),
        ("Groups of similar comments", counts["similar_groups"]),
        ("Clauses changed between drafts", counts["clauses_changed"]), ("", ""),
        ("Accepted but not implemented", counts.get("NOT_IMPLEMENTED", 0)),
        ("Changes with no linked comment", counts["UNTRACED_CHANGE"]),
        ("Changed despite rejection", counts.get("CHANGED_DESPITE_REJECTION", 0)),
        ("No decision recorded", counts.get("NO_DECISION", 0)),
        ("Unknown clause references", counts.get("UNKNOWN_CLAUSE", 0)),
        ("Accepted and implemented (to confirm)", counts.get("IMPLEMENTED", 0)), ("", ""),
        ("Flags are prompts for review, not findings. Decisions remain with the convener.", ""),
    ]
    for r in rows:
        ws.append(list(r))
    ws["A1"].font = Font(bold=True, size=14)
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 10
    for r in range(8, 13):
        if ws.cell(r, 2).value:
            ws.cell(r, 1).fill = ws.cell(r, 2).fill = FILL["high" if r == 8 else "medium"]

    _sheet(wb, "Disposition", disposition, DISPOSITION_COLS)
    _sheet(wb, "Traceability", traced, TRACE_COLS, traced["severity"])
    _sheet(wb, "Clause changes", register, REGISTER_COLS, register["severity"])
    wb.save(path)
    return path


def write_html(path: str | Path, traced: pd.DataFrame, register: pd.DataFrame, title: str) -> Path:
    path = Path(path)
    counts = summary_counts(traced, register)
    esc = html.escape

    tiles = [
        ("Accepted, not implemented", counts.get("NOT_IMPLEMENTED", 0), "high"),
        ("Untraced changes", counts["UNTRACED_CHANGE"], "medium"),
        ("Changed despite rejection", counts.get("CHANGED_DESPITE_REJECTION", 0), "medium"),
        ("No decision", counts.get("NO_DECISION", 0), "medium"),
        ("Implemented (to confirm)", counts.get("IMPLEMENTED", 0), "ok"),
    ]
    tile_html = "".join(
        f'<div class="tile {sev if n else "zero"}"><span class="n">{n}</span><span>{esc(label)}</span></div>'
        for label, n, sev in tiles)

    flagged = traced[traced["severity"].isin(["high", "medium"])]
    trace_rows = "".join(
        f'<tr class="{esc(r.severity)}"><td>{esc(r.comment_id)}</td><td>{esc(r.member)}</td>'
        f'<td>{esc(r.clause)}</td><td>{esc(r.decision or "—")}</td><td><b>{esc(r.trace_status)}</b><br>'
        f'<small>{esc(r.meaning)}</small></td><td>{esc(r.comment)}</td></tr>'
        for r in flagged.itertuples())
    untraced = register[register["severity"] == "medium"]
    reg_rows = "".join(
        f'<tr class="medium"><td>{esc(r.clause)}</td><td>{esc(r.title)}</td><td>{esc(r.trace_note)}</td>'
        f'<td><code>{esc(r.what_changed)}</code></td></tr>' for r in untraced.itertuples())

    page = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>DraftLoop report</title>
<style>
:root{{--ink:#1c2430;--muted:#5b6675;--line:#d9dee5;--high:#f8d7da;--medium:#fff3cd;--ok:#d1e7dd;--bg:#ffffff}}
body{{font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif;color:var(--ink);background:var(--bg);margin:0;padding:24px 16px}}
main{{max-width:1100px;margin:auto}} h1{{font-size:22px;margin:0 0 4px}} .sub{{color:var(--muted);margin:0 0 20px}}
.tiles{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:28px}}
.tile{{border:1px solid var(--line);border-radius:8px;padding:12px;display:flex;flex-direction:column}}
.tile .n{{font-size:28px;font-weight:700}} .tile.high{{background:var(--high)}} .tile.medium{{background:var(--medium)}}
.tile.ok{{background:var(--ok)}} .tile.zero{{color:var(--muted)}}
h2{{font-size:17px;margin:24px 0 8px}} .wrap{{overflow-x:auto}}
table{{border-collapse:collapse;width:100%;font-size:14px}} th,td{{border-bottom:1px solid var(--line);padding:6px 8px;text-align:left;vertical-align:top}}
th{{background:#1f3a5f;color:#fff}} td:nth-child(-n+3){{white-space:nowrap}} tr.high td{{background:var(--high)}} tr.medium td{{background:var(--medium)}}
code{{white-space:pre-wrap;font-size:12px}} footer{{color:var(--muted);font-size:13px;margin-top:28px}}
</style></head><body><main>
<h1>{esc(title)}</h1>
<p class="sub">{counts['comments']} comments · {counts['clauses_changed']} clauses changed · {counts['similar_groups']} group{'s' if counts['similar_groups'] != 1 else ''} of similar comments · generated {date.today().isoformat()}</p>
<div class="tiles">{tile_html}</div>
<h2>Comments needing attention</h2>
<div class="wrap"><table><tr><th>ID</th><th>Member</th><th>Clause</th><th>Decision</th><th>Trace status</th><th>Comment</th></tr>{trace_rows or '<tr><td colspan="6">None</td></tr>'}</table></div>
<h2>Changes needing attention</h2>
<div class="wrap"><table><tr><th>Clause</th><th>Title</th><th>Note</th><th>What changed</th></tr>{reg_rows or '<tr><td colspan="4">None</td></tr>'}</table></div>
<footer>DraftLoop prototype. Flags are prompts for human review, not findings; decisions remain with the convener and committee.</footer>
</main></body></html>"""
    path.write_text(page, encoding="utf-8")
    return path
