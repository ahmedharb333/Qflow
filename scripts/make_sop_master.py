"""
Qflow Consulting — SOP Master workbook generator.

Builds ONE workbook with:
  - "Procedure - EXAMPLE (Recruitment)" — a filled demo SOP, full structure
  - "Procedure_BLANK" (hidden)          — the same structure, empty, duplicated
                                           by the Apps Script menu for each new
                                           department/process SOP
  - "ReadMe"                            — instructions

Row layout is IDENTICAL and FIXED across every "Procedure - *" tab so the
Apps Script generators can use hardcoded row numbers safely.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
NAVY = "1F3864"
LIGHT = "DCE6F1"
INPUT_FILL = "FFFDE0"
BAND_FILL = "E8EEF7"
thin = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# ---- fixed row map (shared by every Procedure tab + both Apps Script generators) ----
ROW_TITLE = 2
HDR_ROWS = [4, 5, 6, 7, 8]          # 5 rows x 2 fields each
PURPOSE_LABEL, PURPOSE_INPUT = 10, 11
SCOPE_LABEL, SCOPE_INPUT = 13, 14
EXCEPT_LABEL, EXCEPT_INPUT = 16, 17
DEFS_LABEL, DEFS_HEADER, DEFS_FIRST, DEFS_ROWS = 19, 20, 21, 6          # 21-26
REFS_LABEL, REFS_FIRST, REFS_ROWS = 28, 29, 4                          # 29-32
POLICY_LABEL, POLICY_FIRST, POLICY_ROWS = 34, 35, 6                    # 35-40
STEPS_LABEL, STEPS_HEADER, STEPS_FIRST, STEPS_ROWS = 42, 43, 44, 16    # 44-59
ESC_LABEL, ESC_HEADER, ESC_FIRST, ESC_ROWS = 61, 62, 63, 4             # 63-66
REV_LABEL, REV_HEADER, REV_FIRST, REV_ROWS = 68, 69, 70, 3             # 70-72
FOOTER_ROW = 74

STEP_COLS = ["B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
STEP_HEADERS = ["Step\nNo.", "Phase", "Activity / Task", "Responsible Role\n(Swimlane)",
                 "Duration\n(Time)", "Input", "Output / Deliverable", "Tool / System",
                 "Reference /\nClause", "Notes", "Type", "Branch:\nIf Yes -> Step",
                 "Branch:\nIf No -> Step", "Consulted\n(C)", "Informed\n(I)"]

COL_WIDTHS = [3, 7, 15, 26, 15, 9, 14, 16, 12, 12, 14, 9, 9, 9, 12, 12, 3]  # A..Q


def label(ws, cell_ref, text, bold=True, fill=NAVY, color="FFFFFF", size=10):
    c = ws[cell_ref]
    c.value = text
    c.font = Font(name=FONT, size=size, bold=bold, color=color)
    if fill:
        c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(vertical="center", wrap_text=True)


def input_cell(ws, rng, value="", size=10, align="left", fill=INPUT_FILL):
    ws.merge_cells(rng)
    c = ws[rng.split(":")[0]]
    c.value = value
    c.font = Font(name=FONT, size=size)
    c.fill = PatternFill("solid", fgColor=fill)
    c.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
    for row in ws[rng]:
        for cell in row:
            cell.border = border


def build_sheet(wb, sheet_name, sample=None, hidden=False):
    """sample: dict with keys header, purpose, scope, exceptions, definitions (list of (term,def)),
    references (list of str), policy (list of str), steps (list of dicts), escalation (list of tuples),
    revision (list of tuples). Pass None for a fully blank template."""
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    for i, w in enumerate(COL_WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells(f"B{ROW_TITLE}:P{ROW_TITLE}")
    ws[f"B{ROW_TITLE}"] = "QFLOW CONSULTING — SOP TEMPLATE"
    ws[f"B{ROW_TITLE}"].font = Font(name=FONT, size=16, bold=True, color=NAVY)

    h = sample["header"] if sample else {}
    hdr_fields = [
        ("Document Name", h.get("doc_name", "[e.g. Recruitment & Onboarding SOP]"), "Doc No.", h.get("doc_no", "[HR-SOP-XXX]")),
        ("Version", h.get("version", "[v1.0]"), "Department / Process", h.get("department", "[e.g. Human Resources]")),
        ("Issue Date", h.get("issue_date", "[DD-MMM-YYYY]"), "Effective Date", h.get("eff_date", "[DD-MMM-YYYY]")),
        ("Next Review Date", h.get("review_date", "[DD-MMM-YYYY]"), "Owner", h.get("owner", "[Role/Name]")),
        ("Approved By", h.get("approved_by", "[Role/Name]"), "Applicable Standard/Clause", h.get("standard", "[ISO 9001 Cl. X]")),
    ]
    for row, (l1, v1, l2, v2) in zip(HDR_ROWS, hdr_fields):
        label(ws, f"B{row}", l1, fill=LIGHT, color="000000")
        input_cell(ws, f"D{row}:F{row}", value=v1)
        label(ws, f"G{row}", l2, fill=LIGHT, color="000000")
        input_cell(ws, f"H{row}:J{row}", value=v2)

    label(ws, f"B{PURPOSE_LABEL}", "PURPOSE"); ws.merge_cells(f"B{PURPOSE_LABEL}:P{PURPOSE_LABEL}")
    input_cell(ws, f"B{PURPOSE_INPUT}:P{PURPOSE_INPUT}", value=sample["purpose"] if sample else "[What this SOP achieves and why it exists]")

    label(ws, f"B{SCOPE_LABEL}", "SCOPE"); ws.merge_cells(f"B{SCOPE_LABEL}:P{SCOPE_LABEL}")
    input_cell(ws, f"B{SCOPE_INPUT}:P{SCOPE_INPUT}", value=sample["scope"] if sample else "[Where this SOP starts/ends, what/who it applies to]")

    label(ws, f"B{EXCEPT_LABEL}", "EXCEPTIONS"); ws.merge_cells(f"B{EXCEPT_LABEL}:P{EXCEPT_LABEL}")
    input_cell(ws, f"B{EXCEPT_INPUT}:P{EXCEPT_INPUT}", value=sample["exceptions"] if sample else "[Cases this SOP does NOT cover, if any]")

    label(ws, f"B{DEFS_LABEL}", "DEFINITIONS & TERMS"); ws.merge_cells(f"B{DEFS_LABEL}:P{DEFS_LABEL}")
    label(ws, f"C{DEFS_HEADER}", "Term", fill=LIGHT, color="000000"); ws.merge_cells(f"C{DEFS_HEADER}:C{DEFS_HEADER}")
    label(ws, f"D{DEFS_HEADER}", "Definition", fill=LIGHT, color="000000"); ws.merge_cells(f"D{DEFS_HEADER}:P{DEFS_HEADER}")
    defs = (sample["definitions"] if sample else []) + [("", "")] * DEFS_ROWS
    for i in range(DEFS_ROWS):
        r = DEFS_FIRST + i
        term, definition = defs[i]
        input_cell(ws, f"C{r}:C{r}", value=term)
        input_cell(ws, f"D{r}:P{r}", value=definition)

    label(ws, f"B{REFS_LABEL}", "REFERENCES & RELATED DOCUMENTS"); ws.merge_cells(f"B{REFS_LABEL}:P{REFS_LABEL}")
    refs = (sample["references"] if sample else []) + [""] * REFS_ROWS
    for i in range(REFS_ROWS):
        r = REFS_FIRST + i
        input_cell(ws, f"B{r}:P{r}", value=refs[i])

    label(ws, f"B{POLICY_LABEL}", "POLICY & GOVERNING PRINCIPLES"); ws.merge_cells(f"B{POLICY_LABEL}:P{POLICY_LABEL}")
    pol = (sample["policy"] if sample else []) + [""] * POLICY_ROWS
    for i in range(POLICY_ROWS):
        r = POLICY_FIRST + i
        input_cell(ws, f"B{r}:P{r}", value=pol[i])

    label(ws, f"B{STEPS_LABEL}", "PROCEDURE STEPS  (Type = Process / Decision / Start / End. For Decision rows, fill Branch Yes/No with target Step No.)")
    ws.merge_cells(f"B{STEPS_LABEL}:P{STEPS_LABEL}")
    for col, htext in zip(STEP_COLS, STEP_HEADERS):
        label(ws, f"{col}{STEPS_HEADER}", htext, fill=LIGHT, color="000000", size=9)
    steps = (sample["steps"] if sample else [])
    for i in range(STEPS_ROWS):
        r = STEPS_FIRST + i
        if i < len(steps):
            s = steps[i]
            vals = [s.get(k, "") for k in ("step_no", "phase", "activity", "role", "duration", "input", "output",
                                             "tool", "reference", "notes", "type", "branch_yes", "branch_no",
                                             "consulted", "informed")]
        else:
            vals = [""] * len(STEP_COLS)
        for col, val in zip(STEP_COLS, vals):
            c = ws[f"{col}{r}"]
            c.value = val
            c.border = border
            c.font = Font(name=FONT, size=9)
            c.alignment = Alignment(wrap_text=True, vertical="center",
                                     horizontal=("center" if col in ("B", "F", "L", "M", "N") else "left"))
            if col != "B":
                c.fill = PatternFill("solid", fgColor=INPUT_FILL)

    lane_list = '"Process Owner,Quality Manager,Internal Auditor,Department Head,Employee,HR,IT/Systems,External Auditor,Manager,Finance"'
    dv = DataValidation(type="list", formula1=lane_list, allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add(f"E{STEPS_FIRST}:E{STEPS_FIRST + STEPS_ROWS - 1}")

    type_list = '"Process,Decision,Start,End"'
    dv2 = DataValidation(type="list", formula1=type_list, allow_blank=True, showDropDown=False)
    ws.add_data_validation(dv2)
    dv2.add(f"L{STEPS_FIRST}:L{STEPS_FIRST + STEPS_ROWS - 1}")

    label(ws, f"B{ESC_LABEL}", "ESCALATION & EXCEPTION HANDLING"); ws.merge_cells(f"B{ESC_LABEL}:P{ESC_LABEL}")
    for c1, c2, t in [("B", "E", "Condition"), ("F", "I", "Action / Resolution"), ("J", "K", "Timeframe / SLA"), ("L", "P", "Escalate To")]:
        label(ws, f"{c1}{ESC_HEADER}", t, fill=LIGHT, color="000000", size=9)
        ws.merge_cells(f"{c1}{ESC_HEADER}:{c2}{ESC_HEADER}")
    esc = (sample["escalation"] if sample else []) + [("", "", "", "")] * ESC_ROWS
    for i in range(ESC_ROWS):
        r = ESC_FIRST + i
        cond, act, sla, to = esc[i]
        input_cell(ws, f"B{r}:E{r}", value=cond)
        input_cell(ws, f"F{r}:I{r}", value=act)
        input_cell(ws, f"J{r}:K{r}", value=sla)
        input_cell(ws, f"L{r}:P{r}", value=to)

    label(ws, f"B{REV_LABEL}", "REVISION HISTORY"); ws.merge_cells(f"B{REV_LABEL}:P{REV_LABEL}")
    for c1, c2, t in [("B", "C", "Version"), ("D", "E", "Date"), ("F", "L", "Description of Change"), ("M", "P", "Approved By")]:
        label(ws, f"{c1}{REV_HEADER}", t, fill=LIGHT, color="000000")
        ws.merge_cells(f"{c1}{REV_HEADER}:{c2}{REV_HEADER}")
    rev = (sample["revision"] if sample else []) + [("", "", "", "")] * REV_ROWS
    for i in range(REV_ROWS):
        r = REV_FIRST + i
        v, d, desc, ap = rev[i]
        input_cell(ws, f"B{r}:C{r}", value=v)
        input_cell(ws, f"D{r}:E{r}", value=d)
        input_cell(ws, f"F{r}:L{r}", value=desc)
        input_cell(ws, f"M{r}:P{r}", value=ap)

    ws.merge_cells(f"B{FOOTER_ROW}:P{FOOTER_ROW}")
    ws[f"B{FOOTER_ROW}"] = ("Yellow cells = fields to edit. Use Qflow Tools menu: \"New SOP for Department/Process\" to "
                             "create another tab like this one; \"Generate SOP Document\" for the written Word/PDF; "
                             "\"Generate Swimlane Diagram\" for the Appendix A flowchart. See the ReadMe tab.")
    ws[f"B{FOOTER_ROW}"].font = Font(name=FONT, size=8, italic=True, color="808080")
    ws[f"B{FOOTER_ROW}"].alignment = Alignment(wrap_text=True, vertical="center")

    ws.freeze_panes = "C2"  # just the title row + row-label column; don't freeze the whole preamble
    if hidden:
        ws.sheet_state = "hidden"
    return ws


def recruitment_demo():
    return {
        "header": {
            "doc_name": "Recruitment & Onboarding SOP",
            "doc_no": "HR-SOP-REC-001",
            "version": "1.0",
            "department": "Human Resources",
            "issue_date": "01-Sep-2026",
            "eff_date": "01-Sep-2026",
            "review_date": "01-Sep-2027",
            "owner": "HR Manager",
            "approved_by": "Executive Office",
            "standard": "ISO 9001 Cl. 7.2 / 8.1",
        },
        "purpose": ("To standardize and control the recruitment and hiring process, ensuring candidates are "
                    "selected against objective, published criteria and in line with UAE labour law and approved HR policy."),
        "scope": ("Applies to all job vacancies across all departments, from the moment a hiring request is raised "
                   "until the new employee starts and their file is activated in the HR system."),
        "exceptions": "Temporary contractors, consultants, and internship placements are excluded unless stated otherwise.",
        "definitions": [
            ("Vacancy", "An approved position in the org structure with budget allocated."),
            ("Evaluation Gate", "A documented Yes/No decision point at which a candidate is passed or rejected."),
            ("Qualified Candidate", "A candidate meeting the core position requirements after initial screening."),
            ("HRIS", "The HR information system used to record employee files (e.g. ZenHR)."),
        ],
        "references": [
            "UAE Labour Law and amendments.",
            "Company HR Policy Manual.",
            "Appendix A: Recruitment process swimlane diagram.",
        ],
        "policy": [
            "Candidates go through a published evaluation sequence; no candidate is rejected without a documented gate decision.",
            "Hiring decisions are based on joint assessment between HR and the hiring manager.",
            "A written employment contract is signed before the employee's start date, per UAE labour law.",
            "Candidate data and evaluation results are kept confidential.",
            "No single manager may finalize a hiring decision without HR and Executive Office approval.",
        ],
        "steps": [
            dict(step_no=1, phase="Phase 1 - Requisition & Posting", activity="Raise a formal hiring request with required competencies", role="Manager", duration="1 day", input="Approved headcount", output="Hiring request", tool="Email / HRIS", reference="-", notes="", type="Process", consulted="HR", informed=""),
            dict(step_no=2, phase="Phase 1 - Requisition & Posting", activity="Verify budget and org structure for the position", role="HR", duration="1 day", input="Hiring request", output="Verified requisition", tool="HRIS", reference="-", notes="", type="Process", consulted="Finance", informed=""),
            dict(step_no=3, phase="Phase 1 - Requisition & Posting", activity="Post the vacancy through approved internal/external channels", role="HR", duration="2 days", input="Verified requisition", output="Published vacancy", tool="Job boards", reference="-", notes="", type="Process", consulted="", informed="Manager"),
            dict(step_no=4, phase="Phase 2 - Screening & HR Interview", activity="Receive and log applications", role="HR", duration="Ongoing", input="Applications", output="Candidate log", tool="HRIS", reference="-", notes="", type="Process", consulted="", informed=""),
            dict(step_no=5, phase="Phase 2 - Screening & HR Interview", activity="Screen CVs against position requirements", role="HR", duration="3 days", input="Candidate log", output="Shortlist", tool="HRIS", reference="-", notes="", type="Process", consulted="", informed=""),
            dict(step_no=6, phase="Phase 2 - Screening & HR Interview", activity="Conduct HR interview (traits, culture fit)", role="HR", duration="45 min", input="Shortlist", output="HR interview result", tool="-", reference="-", notes="", type="Process", consulted="", informed=""),
            dict(step_no=7, phase="Phase 2 - Screening & HR Interview", activity="Gate: did the candidate pass the HR interview?", role="HR", duration="10 min", input="HR interview result", output="Gate decision", tool="-", reference="-", notes="", type="Decision", branch_yes=8, branch_no=15, consulted="", informed=""),
            dict(step_no=8, phase="Phase 3 - Technical Interview & Decision", activity="Conduct technical interview", role="Manager", duration="1 hr", input="Gate decision (Yes)", output="Technical interview result", tool="-", reference="-", notes="", type="Process", consulted="", informed="HR"),
            dict(step_no=9, phase="Phase 3 - Technical Interview & Decision", activity="Gate: did the candidate pass the technical interview?", role="Manager", duration="10 min", input="Technical interview result", output="Gate decision", tool="-", reference="-", notes="", type="Decision", branch_yes=10, branch_no=15, consulted="", informed=""),
            dict(step_no=10, phase="Phase 3 - Technical Interview & Decision", activity="Prepare evaluation report and recommendation", role="HR", duration="1 day", input="Interview results", output="Evaluation report", tool="HRIS", reference="-", notes="", type="Process", consulted="Manager", informed=""),
            dict(step_no=11, phase="Phase 3 - Technical Interview & Decision", activity="Submit report to Executive Office for final approval", role="Process Owner", duration="2 days", input="Evaluation report", output="Approval decision", tool="Email", reference="-", notes="", type="Decision", branch_yes=12, branch_no=10, consulted="", informed=""),
            dict(step_no=12, phase="Phase 4 - Contracting & Documentation", activity="Notify candidate and issue written offer", role="HR", duration="1 day", input="Approval decision (Yes)", output="Signed offer", tool="Email", reference="-", notes="", type="Process", consulted="", informed="Manager"),
            dict(step_no=13, phase="Phase 4 - Contracting & Documentation", activity="Collect documents, sign contract, register in HRIS", role="HR", duration="3 days", input="Signed offer", output="Employee file", tool="HRIS", reference="-", notes="", type="Process", consulted="", informed=""),
            dict(step_no=14, phase="Phase 4 - Contracting & Documentation", activity="Employee starts; file activated in the system", role="HR", duration="1 day", input="Employee file", output="Active employee record", tool="HRIS", reference="ISO 9001 7.2", notes="End of procedure", type="End", consulted="", informed="Manager"),
            dict(step_no=15, phase="Rejected Candidates (from either gate)", activity="Reject candidate and notify formally", role="HR", duration="1 day", input="Gate decision (No)", output="Rejection notice", tool="Email", reference="-", notes="End of path", type="End", consulted="", informed="Manager"),
        ],
        "escalation": [
            ("Disagreement between Manager and HR on a candidate", "Escalate to Executive Office for a final ruling", "3 business days", "Executive Office"),
            ("Executive Office does not approve the recommendation", "New recommendation prepared; evaluation results reviewed again", "5 business days", "HR Manager"),
            ("Any exception to the evaluation sequence (skipped stage, extended deadline)", "Requires prior documented approval", "Before proceeding", "HR Manager + Executive Office"),
        ],
        "revision": [("1.0", "01-Sep-2026", "Initial release", "Executive Office")],
    }


def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    build_sheet(wb, "EXAMPLE - Recruitment SOP", sample=recruitment_demo(), hidden=False)
    build_sheet(wb, "Procedure_BLANK", sample=None, hidden=True)

    ws2 = wb.create_sheet("ReadMe")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 2
    ws2.column_dimensions["B"].width = 105

    def note(ws, row, text, bold=False, size=10, color="000000"):
        c = ws[f"B{row}"]
        c.value = text
        c.font = Font(name=FONT, size=size, bold=bold, color=color)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        return row + 1

    r = 2
    r = note(ws2, r, "Qflow SOP Master — how this workbook works", bold=True, size=13, color=NAVY)
    r += 1
    r = note(ws2, r, "This one Google Sheet is the master for every SOP you write — recruitment, termination, "
                     "training, customer service, whatever department needs one. Each SOP lives on its own tab, "
                     "same structure every time, so the generators below always know where to look.")
    r += 1
    r = note(ws2, r, "1. Create a new SOP", bold=True, size=12, color=NAVY)
    r = note(ws2, r, "Qflow Tools > \"New SOP for Department/Process...\" — type a name (e.g. \"Termination\", "
                     "\"Training\", \"Customer Service\"). A new tab is created with the full blank structure: "
                     "document control, purpose, scope, exceptions, definitions, references, policy, procedure "
                     "steps (with phases + decision gates), escalation rules, revision history.")
    r += 1
    r = note(ws2, r, "2. Fill it in", bold=True, size=12, color=NAVY)
    r = note(ws2, r, "Yellow cells are editable. In the step table: Phase groups steps under a heading (matches "
                     "colored bands in the diagram). Type = Decision needs Branch Yes/No step numbers. Consulted "
                     "(C) / Informed (I) are role names — used to build the RACI table in the generated document "
                     "automatically (Responsible Role = R/A by default).")
    r += 1
    r = note(ws2, r, "3. Generate the outputs", bold=True, size=12, color=NAVY)
    r = note(ws2, r, "With the SOP's tab active (open), use Qflow Tools:")
    r = note(ws2, r, "  - \"Generate SOP Document\" -> a written Google Doc + PDF: all sections above, RACI table, "
                     "numbered procedure steps by phase.")
    r = note(ws2, r, "  - \"Generate Swimlane Diagram (.drawio)\" -> Appendix-A style diagram: one column per role, "
                     "colored phase bands, decision diamonds, branching arrows. Open at app.diagrams.net, "
                     "File > Export as > PDF.")
    r += 1
    r = note(ws2, r, "Keep in mind", bold=True, size=12, color=NAVY)
    r = note(ws2, r, "- Don't insert/delete rows inside the template sections — the generators use fixed row "
                     "numbers. There's room for 16 steps per SOP; if you need more, ask to have the template "
                     "regenerated with a taller step table rather than inserting rows by hand.")
    r = note(ws2, r, "- \"Procedure_BLANK\" is the hidden master copy the menu duplicates from — don't fill it in "
                     "directly or unhide/rename it.")
    r = note(ws2, r, "- \"EXAMPLE - Recruitment SOP\" is a fully filled sample — duplicate the pattern, "
                     "or use New SOP to start clean.")
    r = note(ws2, r, "- Phase bands appear in the diagram in ROW order, and a new band starts whenever the Phase "
                     "text changes — so keep rows of the same phase together. A rejection/dead-end step that "
                     "physically comes after a later phase in the row list (see the EXAMPLE's \"Rejected "
                     "Candidates\" step) will get its own small band rather than merging back into an earlier one "
                     "— that's expected, not a bug.")

    ws2.sheet_view.showGridLines = False
    wb.move_sheet("ReadMe", offset=-2)  # put ReadMe first

    wb.save("Qflow_SOP_Master.xlsx")
    print("written")


if __name__ == "__main__":
    main()
