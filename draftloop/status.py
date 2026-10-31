"""Project status tracker: early warning for publication projects.

Reads a simple register of projects (one row per project) and flags those
at risk, using transparent rules rather than prediction:

  * RED    the target date has passed, or the current stage has run longer
           than its time limit;
  * AMBER  the target date is within the warning window, the stage has used
           most of its time limit, or the stage has no start date;
  * GREEN  otherwise.

Stage names and time limits are configurable, because every standards body
defines its own. The defaults are generic placeholders, not any body's rules.

Usage:
    draftloop-status --projects data/synthetic/projects.xlsx --out output
    draftloop-status --projects projects.csv --stages my_stages.json --as-of 2026-11-01
"""

from __future__ import annotations

import argparse
import html
import json
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Generic stages with an indicative time limit in months. Replace with your
# own directives' stages and limits (see docs/stages.example.json).
DEFAULT_STAGES: dict[str, int] = {
    "Proposal": 3,
    "Working draft": 12,
    "Committee draft": 12,
    "Enquiry": 6,
    "Final draft": 6,
    "Publication": 3,
}
WARN_DAYS = 60          # target date within this many days -> amber
WARN_FRACTION = 0.75    # stage has used this share of its limit -> amber
FAST_FRACTION = 0.5     # optimistic pace for the stages still to go

ALIASES = {
    "project": ["project", "project id", "reference", "ref", "id"],
    "title": ["title", "project title", "document"],
    "stage": ["stage", "current stage", "status"],
    "stage_start": ["stage start", "stage started", "stage_start", "since"],
    "target": ["target", "target date", "target publication", "due"],
    "lead": ["lead", "convener", "convenor", "project leader", "owner"],
    "next_action": ["next action", "next step", "next milestone", "action"],
}

RAG_FILL = {"RED": "F8D7DA", "AMBER": "FFF3CD", "GREEN": "D1E7DD"}
RAG_ORDER = {"RED": 0, "AMBER": 1, "GREEN": 2}


def _read(path: Path) -> pd.DataFrame:
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    return pd.read_csv(path)


def _columns(df: pd.DataFrame) -> dict[str, str]:
    lookup = {str(c).strip().lower(): c for c in df.columns}
    found = {}
    for field, names in ALIASES.items():
        for n in names:
            if n in lookup:
                found[field] = lookup[n]
                break
    missing = [f for f in ("project", "stage", "target") if f not in found]
    if missing:
        raise ValueError(f"Missing column(s) {missing}. Columns present: {list(df.columns)}")
    return found


def _to_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)) or value == "":
        return None
    if isinstance(value, (datetime, pd.Timestamp)):
        return value.date()
    if isinstance(value, date):
        return value
    try:
        return pd.to_datetime(str(value), dayfirst=False).date()
    except (ValueError, TypeError):
        return None


def assess(projects: pd.DataFrame, stages: dict[str, int] = DEFAULT_STAGES,
           as_of: date | None = None) -> pd.DataFrame:
    """Add RAG status and the reasons for it to each project."""
    as_of = as_of or date.today()
    cols = _columns(projects)
    stage_lookup = {s.lower(): (s, months) for s, months in stages.items()}
    order = list(stages)
    rows = []

    for _, p in projects.iterrows():
        raw_stage = str(p[cols["stage"]]).strip()
        stage_name, limit = stage_lookup.get(raw_stage.lower(), (raw_stage, None))
        start = _to_date(p[cols["stage_start"]]) if "stage_start" in cols else None
        target = _to_date(p[cols["target"]])
        reasons, rag = [], "GREEN"

        def flag(level: str, why: str) -> None:
            nonlocal rag
            reasons.append(why)
            if RAG_ORDER[level] < RAG_ORDER[rag]:
                rag = level

        months_in_stage = None
        if start:
            months_in_stage = round((as_of - start).days / 30.44, 1)
        if limit is None:
            flag("AMBER", f"stage '{raw_stage}' not in the stage list")
        elif start is None:
            flag("AMBER", "no stage start date")
        elif months_in_stage > limit:
            flag("RED", f"{months_in_stage:g} months in stage (limit {limit})")
        elif months_in_stage >= WARN_FRACTION * limit:
            flag("AMBER", f"{months_in_stage:g} of {limit} months used in stage")

        days_left = (target - as_of).days if target else None
        if target is None:
            flag("AMBER", "no target date")
        elif days_left < 0:
            flag("RED", f"target date passed {-days_left} days ago")
        elif days_left <= WARN_DAYS:
            flag("AMBER", f"target date in {days_left} days")

        # Early warning: even if the stages still to go run at a fraction of
        # their limits, is there enough time before the target date?
        if stage_name in order and days_left is not None and days_left >= 0:
            remaining = order[order.index(stage_name) + 1:]
            needed = FAST_FRACTION * sum(stages[s] for s in remaining)
            if remaining and needed * 30.44 > days_left:
                n = len(remaining)
                flag("AMBER", f"{n} stage{'s' if n > 1 else ''} still to go "
                              f"need{'' if n > 1 else 's'} about {needed:g} months even at half "
                              f"{'their limits' if n > 1 else 'its limit'}; "
                              f"{days_left / 30.44:.1f} months to target")

        rows.append({
            "project": p[cols["project"]],
            "title": p[cols["title"]] if "title" in cols else "",
            "stage": stage_name,
            "months_in_stage": months_in_stage,
            "stage_limit_months": limit,
            "target": target.isoformat() if target else "",
            "days_to_target": days_left,
            "lead": p[cols["lead"]] if "lead" in cols else "",
            "next_action": p[cols["next_action"]] if "next_action" in cols else "",
            "rag": rag,
            "reasons": "; ".join(reasons) if reasons else "on track",
        })

    out = pd.DataFrame(rows)
    out["_o"] = out["rag"].map(RAG_ORDER)
    return out.sort_values(["_o", "days_to_target"], na_position="first").drop(columns="_o").reset_index(drop=True)


HEADERS = {
    "rag": "Status", "project": "Project", "title": "Title", "stage": "Stage",
    "months_in_stage": "Months in stage", "stage_limit_months": "Stage limit (months)",
    "target": "Target date", "days_to_target": "Days to target", "lead": "Lead",
    "next_action": "Next action", "reasons": "Why",
}


def write_excel(path: Path, df: pd.DataFrame, as_of: date) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "Project status"
    ws.append([f"Project status as of {as_of.isoformat()} (DraftLoop prototype)"])
    ws["A1"].font = Font(bold=True, size=13)
    counts = df["rag"].value_counts()
    ws.append([f"Red {counts.get('RED', 0)} · Amber {counts.get('AMBER', 0)} · Green {counts.get('GREEN', 0)}"])
    ws.append([])
    ws.append(list(HEADERS.values()))
    for cell in ws[4]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3A5F")
    for row in df[list(HEADERS)].itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in row])
        r = ws.max_row
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=RAG_FILL[row[0]])
        for c in ws[r]:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    widths = [9, 12, 36, 16, 10, 10, 12, 10, 16, 36, 48]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"
    wb.save(path)
    return path


def write_html(path: Path, df: pd.DataFrame, as_of: date) -> Path:
    esc = html.escape
    counts = df["rag"].value_counts()
    tiles = "".join(
        f'<div class="tile {k.lower()}"><span class="n">{counts.get(k, 0)}</span><span>{label}</span></div>'
        for k, label in (("RED", "Red: act now"), ("AMBER", "Amber: watch"), ("GREEN", "Green: on track")))
    body = "".join(
        f'<tr><td><span class="pill {esc(r.rag.lower())}">{esc(r.rag)}</span></td><td>{esc(str(r.project))}</td>'
        f'<td>{esc(str(r.title))}</td><td>{esc(r.stage)}</td><td>{esc(r.target)}</td>'
        f'<td>{esc(str(r.lead))}</td><td>{esc(r.reasons)}</td><td>{esc(str(r.next_action))}</td></tr>'
        for r in df.itertuples())
    page = f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Project status</title>
<style>
body{{font:15px/1.5 system-ui,-apple-system,Segoe UI,sans-serif;color:#1c2430;background:#fff;margin:0;padding:24px 16px}}
main{{max-width:1150px;margin:auto}} h1{{font-size:22px;margin:0 0 4px}} .sub{{color:#5b6675;margin:0 0 20px}}
.tiles{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:10px;margin-bottom:24px}}
.tile{{border:1px solid #d9dee5;border-radius:8px;padding:12px;display:flex;flex-direction:column}}
.tile .n{{font-size:28px;font-weight:700}} .tile.red{{background:#f8d7da}} .tile.amber{{background:#fff3cd}} .tile.green{{background:#d1e7dd}}
.wrap{{overflow-x:auto}} table{{border-collapse:collapse;width:100%;font-size:14px}}
th,td{{border-bottom:1px solid #d9dee5;padding:6px 8px;text-align:left;vertical-align:top}} th{{background:#1f3a5f;color:#fff}}
td:nth-child(2),td:nth-child(5),td:nth-child(6){{white-space:nowrap}}
.pill{{border-radius:10px;padding:1px 8px;font-size:12px;font-weight:700}} .pill.red{{background:#f8d7da}}
.pill.amber{{background:#fff3cd}} .pill.green{{background:#d1e7dd}}
footer{{color:#5b6675;font-size:13px;margin-top:24px}}
</style></head><body><main>
<h1>Publication projects: status</h1><p class="sub">As of {as_of.isoformat()} · {len(df)} projects · rules: stage time limits and target dates</p>
<div class="tiles">{tiles}</div>
<div class="wrap"><table><tr><th>Status</th><th>Project</th><th>Title</th><th>Stage</th><th>Target</th><th>Lead</th><th>Why</th><th>Next action</th></tr>{body}</table></div>
<footer>DraftLoop prototype. Status follows transparent rules set by the user; it is a prompt for discussion with project leads, not a judgement.</footer>
</main></body></html>"""
    path.write_text(page, encoding="utf-8")
    return path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="draftloop-status", description="Flag publication projects at risk of slipping.")
    ap.add_argument("--projects", required=True, help="Project register (.xlsx or .csv)")
    ap.add_argument("--stages", help="JSON file of {stage: time limit in months}, in order")
    ap.add_argument("--as-of", help="Assessment date, YYYY-MM-DD (default: today)")
    ap.add_argument("--out", default="output")
    a = ap.parse_args(argv)

    stages = json.loads(Path(a.stages).read_text()) if a.stages else DEFAULT_STAGES
    as_of = date.fromisoformat(a.as_of) if a.as_of else date.today()
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    df = assess(_read(Path(a.projects)), stages, as_of)
    write_excel(out / "project_status.xlsx", df, as_of)
    write_html(out / "project_status.html", df, as_of)
    c = df["rag"].value_counts()
    print(f"{len(df)} projects: {c.get('RED', 0)} red, {c.get('AMBER', 0)} amber, {c.get('GREEN', 0)} green "
          f"-> {out / 'project_status.xlsx'}, {out / 'project_status.html'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
