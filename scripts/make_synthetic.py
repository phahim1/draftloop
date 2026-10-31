"""Generate the synthetic demo data.

Everything here is invented: a fictional draft standard ("DS 001, Example
digital thermometer"), fictional member bodies (Member A-H) and fictional
comments. Problems are planted deliberately so the demo has something to find:

  * 3 accepted comments that were NOT implemented (clauses 6.2, 6.3, 7.2)
  * 2 changes that no comment explains (clauses 4 and 7.4)
  * 2 comments (one rejected, one noted) whose clause changed anyway (clause 5.3)
  * 1 comment with no decision (clause 8), 1 unknown clause reference (9.4)
  * 1 clause deleted following an accepted comment (A.2)
  * a cluster of near-duplicate comments on clause 5.2

Run:  python scripts/make_synthetic.py  (writes to data/synthetic/)
"""

from pathlib import Path

import pandas as pd
from docx import Document

OUT = Path(__file__).resolve().parents[1] / "data" / "synthetic"

V1 = [
    ("1", "Scope", "This document specifies requirements and test procedures for example digital thermometers used for measurements subject to regulation."),
    ("2", "Normative references", "There are no normative references in this document."),
    ("3", "Terms and definitions", "For the purposes of this document, the following terms apply."),
    ("3.1", "maximum permissible error", "extreme value of measurement error permitted by this document for a given instrument"),
    ("3.2", "reference conditions", "conditions of use prescribed for evaluating the performance of an instrument"),
    ("4", "Units of measurement", "Temperature shall be indicated in degrees Celsius."),
    ("5", "Metrological requirements", "Instruments shall meet the requirements of this clause under rated operating conditions."),
    ("5.1", "Accuracy classes", "Instruments are classified into accuracy classes I and II."),
    ("5.2", "Maximum permissible errors", "The maximum permissible error for class I shall be 0.2 °C and for class II shall be 0.5 °C."),
    ("5.3", "Influence quantities", "The instrument shall meet the maximum permissible errors at ambient temperatures from 10 °C to 30 °C."),
    ("6", "Technical requirements", "Instruments shall be designed to prevent unauthorised adjustment."),
    ("6.1", "Indicating device", "The indication shall be clear and unambiguous, with a resolution of 0.1 °C."),
    ("6.2", "Software", "Legally relevant software shall be identified by a version number."),
    ("6.3", "Markings", "Each instrument shall bear the manufacturer's name and the accuracy class."),
    ("7", "Test procedures", "Tests shall be carried out under reference conditions unless stated otherwise."),
    ("7.1", "General", "The test equipment shall be calibrated and traceable to national standards."),
    ("7.2", "Accuracy test", "The error shall be determined at three test points across the measuring range."),
    ("7.3", "Temperature influence test", "The instrument shall be exposed to 10 °C and 30 °C for one hour before readings are taken."),
    ("7.4", "Durability test", "The instrument shall be operated for 1 000 cycles, after which the accuracy test is repeated."),
    ("8", "Documentation", "The manufacturer shall provide operating instructions."),
    ("A", "Test report format", "Annex A (informative) gives a template for test reports."),
    ("A.1", "Content of the report", "The report shall identify the instrument, the tests performed and the results."),
    ("A.2", "Report numbering", "Each report shall have a unique number."),
]

# Changes made for draft 2 (clause -> new body). Clauses not listed are unchanged.
V2_CHANGES = {
    "1": "This document specifies metrological and technical requirements and test procedures for example digital thermometers used for measurements subject to regulation.",
    "3.1": "extreme value of measurement error, with respect to a known reference value, permitted by this document for a given instrument",
    "4": "Temperature shall be indicated in degrees Celsius or kelvin.",                       # untraced
    "5.2": "The maximum permissible error for class I shall be 0.1 °C and for class II shall be 0.3 °C.",
    "5.3": "The instrument shall meet the maximum permissible errors at ambient temperatures from 5 °C to 40 °C.",  # changed despite rejection
    "7.3": "The instrument shall be exposed to 5 °C and 40 °C for two hours before readings are taken.",
    "7.4": "The instrument shall be operated for 2 000 cycles, after which the accuracy test is repeated.",  # untraced
}
V2_REMOVED = {"A.2"}

# (id, member, clause as written, type, comment, proposed change)
COMMENTS = [
    ("C01", "Member A", "General", "ge", "The draft is well structured and ready for the next stage once technical comments are resolved.", ""),
    ("C02", "Member D", "general", "ed", "Use the current drafting template throughout.", ""),
    ("C03", "Member B", "Clause 1", "ed", "Scope should mention both metrological and technical requirements.", "requirements and test procedures -> metrological and technical requirements and test procedures"),
    ("C04", "Member C", "3.1", "te", "Definition of maximum permissible error should state the reference value.", "with respect to a known reference value"),
    ("C05", "Member E", "3.2", "ed", "Definition is clear; no change needed.", ""),
    ("C06", "Member A", "5.2", "te", "The maximum permissible error for class I is too large for current technology; reduce to 0.1 °C.", "class I shall be 0.1 °C"),
    ("C07", "Member F", "Clause 5.2", "te", "Maximum permissible error for class I is too large given current technology; reduce to 0.1 °C.", "class I shall be 0.1 °C"),
    ("C08", "Member G", "5.2 a)", "te", "The class I maximum permissible error is too large for current technology, propose 0.1 °C.", "0.1 °C for class I"),
    ("C09", "Member B", "5.2", "te", "Class II maximum permissible error should be reduced to 0.3 °C.", "class II shall be 0.3 °C"),
    ("C10", "Member H", "5.3", "te", "Extend the ambient temperature range to 0 °C to 50 °C to reflect field use.", "0 °C to 50 °C"),
    ("C11", "Member C", "5.1", "ed", "Add a note explaining how accuracy classes are assigned.", ""),
    ("C12", "Member D", "6.1", "te", "A resolution of 0.1 °C is appropriate; keep as drafted.", ""),
    ("C13", "Member E", "6.2", "te", "Software identification should include a checksum as well as a version number.", "identified by a version number and a checksum"),
    ("C14", "Member A", "6.3", "ed", "Markings should also include the serial number.", "the manufacturer's name, the serial number and the accuracy class"),
    ("C15", "Member G", "§7.2", "te", "Three test points are too few; use five test points across the measuring range.", "at five test points"),
    ("C16", "Member F", "7.3", "te", "Align the influence test with the extended range and lengthen stabilisation to two hours.", "5 °C and 40 °C for two hours"),
    ("C17", "Member B", "7.1", "ed", "State that calibration certificates shall be available on request.", ""),
    ("C18", "Member H", "8", "ed", "Documentation should include installation instructions.", "operating and installation instructions"),
    ("C19", "Member C", "9.4", "te", "Clause 9.4 should refer to the conformity assessment procedure.", ""),
    ("C20", "Member D", "Annex A.2", "ed", "Report numbering is an internal matter for laboratories; delete A.2.", "Delete A.2"),
    ("C21", "Member E", "A.1", "ed", "The report should also state the measurement uncertainty.", ""),
    ("C22", "Member G", "5.3", "te", "Consider extending the ambient temperature range to cover outdoor use.", ""),
]

DECISIONS = [
    ("C01", "Noted", "Thank you."),
    ("C02", "Accepted", "Template will be applied at publication stage."),
    ("C03", "Accepted", "Scope amended."),
    ("C04", "Accepted", "Definition amended."),
    ("C05", "Noted", ""),
    ("C06", "Accepted", "See C07, C08."),
    ("C07", "Accepted", "Same as C06."),
    ("C08", "Accepted", "Same as C06."),
    ("C09", "Accepted", ""),
    ("C10", "Rejected", "Range already under discussion in the project group; no change at this stage."),
    ("C11", "Rejected", "Covered by clause 5."),
    ("C12", "Noted", ""),
    ("C13", "Accepted", "Checksum to be added."),
    ("C14", "Accepted in principle", "Serial number to be added; wording by editor."),
    ("C15", "Accepted", "Five test points."),
    ("C16", "Accepted", ""),
    ("C17", "Rejected", "Outside the scope of this document."),
    ("C18", "", ""),                      # no decision yet
    ("C19", "Noted", "There is no clause 9.4 in this draft."),
    ("C20", "Accepted", "A.2 deleted."),
    ("C21", "Rejected", "Covered by A.1 content requirements."),
    ("C22", "Noted", "See C10."),
]


# Fictional project register for the status tracker, assessed as of 2026-11-01.
PROJECTS = [
    ("DS 001", "Example digital thermometer", "Committee draft", "2026-03-01", "2027-06-30", "Lead A", "Circulate CD2 for comment"),
    ("DS 002", "Example water meter software", "Working draft", "2025-06-15", "2027-12-31", "Lead B", "Agree scope of software annex"),
    ("DS 003", "Example fuel dispenser", "Final draft", "2026-08-01", "2026-11-20", "Lead C", "Final editorial check"),
    ("DS 004", "Example weighing instrument", "Enquiry", "2026-09-10", "2027-09-30", "Lead D", "Ballot closes 10 Dec"),
    ("DS 005", "Example gas analyser", "Committee draft", "2025-02-01", "2026-09-30", "Lead E", "Resolve 140 comments"),
    ("DS 006", "Example taximeter", "Proposal", "2026-10-01", "2028-12-31", "Lead F", "Call for experts"),
    ("DS 007", "Example energy meter", "Publication", "2026-10-15", "2027-01-31", "Lead G", "Typesetting"),
    ("DS 008", "Example speed meter", "Working draft", "2026-07-01", "2027-03-31", "Lead H", "WD2 meeting"),
]


def write_draft(path: Path, title: str, clauses) -> None:
    doc = Document()
    doc.add_heading(title, level=0)
    doc.add_paragraph("Fictional example for demonstration only. Not a real standard.")
    for key, heading, body in clauses:
        level = 1 if "." not in key else 2
        label = f"Annex {key}" if key == "A" else key
        doc.add_heading(f"{label} {heading}", level=level)
        doc.add_paragraph(body)
    doc.save(path)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    write_draft(OUT / "draft_v1.docx", "DS 001 Example digital thermometer - Committee draft 1", V1)
    v2 = [(k, h, V2_CHANGES.get(k, b)) for k, h, b in V1 if k not in V2_REMOVED]
    write_draft(OUT / "draft_v2.docx", "DS 001 Example digital thermometer - Committee draft 2", v2)

    pd.DataFrame(COMMENTS, columns=["ID", "MB", "Clause/Subclause", "Type of comment", "Comments", "Proposed change"]
                 ).to_excel(OUT / "comments.xlsx", index=False)
    pd.DataFrame(DECISIONS, columns=["ID", "Decision", "Observations"]).to_excel(OUT / "decisions.xlsx", index=False)
    pd.DataFrame(PROJECTS, columns=["Project", "Title", "Stage", "Stage start", "Target date", "Lead", "Next action"]
                 ).to_excel(OUT / "projects.xlsx", index=False)
    print(f"Synthetic data written to {OUT}")


if __name__ == "__main__":
    main()
