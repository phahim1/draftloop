"""Tests run against the synthetic demo data, which has planted problems."""

import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

from draftloop.clauses import in_scope, normalise_clause, sort_key
from draftloop.compile import compile_comments
from draftloop.ingest import apply_decisions, load_comments, load_draft, normalise_decision
from draftloop.trace import trace

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data" / "synthetic"


@pytest.fixture(scope="module", autouse=True)
def synthetic_data():
    subprocess.run([sys.executable, str(ROOT / "scripts" / "make_synthetic.py")], check=True)


@pytest.mark.parametrize("raw, expected", [
    ("Clause 5.2", "5.2"), ("5.2 a)", "5.2"), ("§7.2", "7.2"), ("Annex A.2", "A.2"),
    ("annex b", "B"), ("A.1", "A.1"), ("General", "general"), ("", "general"),
    ("Sub-clause 10.3.1", "10.3.1"), ("cl. 4", "4"),
])
def test_normalise_clause(raw, expected):
    assert normalise_clause(raw) == expected


def test_sort_and_scope():
    keys = ["A.1", "10", "2.1", "general", "2"]
    assert sorted(keys, key=sort_key) == ["general", "2", "2.1", "10", "A.1"]
    assert in_scope("5.2.1", "5.2") and in_scope("5.2", "5.2")
    assert not in_scope("5.20", "5.2")


@pytest.mark.parametrize("raw, expected", [
    ("Accepted", "accepted"), ("Accepted in principle", "accepted in principle"),
    ("REJECTED", "rejected"), ("Noted", "noted"), ("", ""), ("maybe", "unrecognised"),
])
def test_normalise_decision(raw, expected):
    assert normalise_decision(raw) == expected


@pytest.fixture(scope="module")
def results():
    comments = apply_decisions(load_comments(DATA / "comments.xlsx"), DATA / "decisions.xlsx")
    compiled = compile_comments(comments)
    traced, register = trace(compiled, load_draft(DATA / "draft_v1.docx"), load_draft(DATA / "draft_v2.docx"))
    return compiled, traced, register


def status_of(traced: pd.DataFrame, cid: str) -> str:
    return traced.loc[traced["comment_id"] == cid, "trace_status"].item()


def test_draft_parsing():
    clauses = load_draft(DATA / "draft_v1.docx")
    assert {"1", "5.2", "7.4", "A", "A.1", "A.2"} <= set(clauses)
    assert "0.2 °C" in clauses["5.2"].text


def test_similar_comments_grouped(results):
    compiled, _, _ = results
    group = compiled.loc[compiled["comment_id"].isin(["C06", "C07", "C08"]), "similar_group"]
    assert group.nunique() == 1 and group.iloc[0] != ""
    assert compiled.loc[compiled["comment_id"] == "C09", "similar_group"].item() == ""


def test_planted_problems_found(results):
    _, traced, register = results
    not_done = set(traced.loc[traced["trace_status"] == "NOT_IMPLEMENTED", "comment_id"])
    assert not_done == {"C13", "C14", "C15"}
    assert status_of(traced, "C10") == "CHANGED_DESPITE_REJECTION"
    assert status_of(traced, "C18") == "NO_DECISION"
    assert status_of(traced, "C19") == "UNKNOWN_CLAUSE"
    assert status_of(traced, "C20") == "IMPLEMENTED"          # accepted deletion
    untraced = set(register.loc[register["trace_note"].str.startswith("UNTRACED"), "clause"])
    assert untraced == {"4", "7.4"}


def test_no_false_alarms(results):
    _, traced, _ = results
    for cid in ["C03", "C04", "C06", "C16"]:
        assert status_of(traced, cid) == "IMPLEMENTED"
    for cid in ["C05", "C11", "C12", "C17", "C21"]:
        assert status_of(traced, cid) == "NO_CHANGE_EXPECTED"


def test_renumbered_clauses_are_followed(tmp_path):
    """A new clause inserted at 5.2 pushes old 5.2 and 5.3 down by one."""
    sys.path.insert(0, str(ROOT / "scripts"))
    from make_synthetic import V1, write_draft

    old = [c for c in V1]
    new = []
    for key, heading, body in V1:
        if key == "5.2":
            new.append(("5.2", "Rated operating conditions", "Rated operating conditions shall be stated by the manufacturer."))
            new.append(("5.3", heading, body.replace("0.2 °C", "0.1 °C")))   # moved and changed
        elif key == "5.3":
            new.append(("5.4", heading, body))                              # moved, unchanged
        else:
            new.append((key, heading, body))
    write_draft(tmp_path / "v1.docx", "v1", old)
    write_draft(tmp_path / "v2.docx", "v2", new)

    comments = pd.DataFrame([
        {"comment_id": "R1", "member": "Member A", "clause_raw": "5.2", "clause": "5.2", "type": "technical",
         "comment": "Reduce class I MPE to 0.1 °C", "proposed_change": "0.1 °C", "decision": "accepted", "rationale": ""},
        {"comment_id": "R2", "member": "Member B", "clause_raw": "5.3", "clause": "5.3", "type": "technical",
         "comment": "Widen the temperature range", "proposed_change": "", "decision": "rejected", "rationale": ""},
        {"comment_id": "R3", "member": "Member C", "clause_raw": "5", "clause": "5", "type": "technical",
         "comment": "Add a clause on rated operating conditions", "proposed_change": "", "decision": "accepted", "rationale": ""},
    ])
    traced, register = trace(comments, load_draft(tmp_path / "v1.docx"), load_draft(tmp_path / "v2.docx"))
    assert status_of(traced, "R1") == "IMPLEMENTED"            # followed to new 5.3
    assert status_of(traced, "R2") == "NO_CHANGE_EXPECTED"     # moved to 5.4 but unchanged
    reg = register.set_index("clause")
    assert reg.loc["5.2", "new_number"] == "5.3"
    assert reg.loc["5.3", "new_number"] == "5.4" and reg.loc["5.3", "change"] == "unchanged"
    assert reg.loc["5.2 (new)", "change"] == "added"
    assert reg.loc["5.2 (new)", "trace_note"] == "Linked to an accepted comment"   # explained by R3


def test_project_status():
    from datetime import date
    from draftloop.status import assess

    df = assess(pd.read_excel(DATA / "projects.xlsx"), as_of=date(2026, 11, 1)).set_index("project")
    assert df.loc["DS 005", "rag"] == "RED"      # target passed and stage over its limit
    assert df.loc["DS 002", "rag"] == "RED"      # 16 months in a 12-month stage
    assert df.loc["DS 003", "rag"] == "AMBER"    # target in 19 days
    assert df.loc["DS 006", "rag"] == "GREEN"
    assert df.loc["DS 004", "rag"] == "GREEN"
    assert "target date passed" in df.loc["DS 005", "reasons"]


def test_cli_end_to_end(tmp_path):
    cmd = [sys.executable, "-m", "draftloop", "--comments", str(DATA / "comments.xlsx"),
           "--decisions", str(DATA / "decisions.xlsx"), "--old", str(DATA / "draft_v1.docx"),
           "--new", str(DATA / "draft_v2.docx"), "--out", str(tmp_path)]
    subprocess.run(cmd, check=True, cwd=ROOT)
    for name in ["disposition.docx", "draftloop_report.xlsx", "draftloop_report.html"]:
        assert (tmp_path / name).exists()
    assert not (tmp_path / "assist_audit.jsonl").exists()      # assistance is off by default
