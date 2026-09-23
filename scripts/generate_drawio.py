"""
Qflow Consulting — SOP Master (Procedure tab) -> draw.io swimlane generator.

Reads a filled-in "Procedure - *" (or "EXAMPLE - *") tab of Qflow_SOP_Master.xlsx
and writes a .drawio (mxGraph XML) file: one column per Responsible Role, colored
phase bands, decision diamonds with Yes/No branches.

Usage:
    python generate_drawio.py "Qflow_SOP_Master.xlsx" "Recruitment SOP" "Recruitment.drawio"

    (second arg = exact sheet/tab name; omit to use the active/first non-hidden
    "Procedure_BLANK"-excluded sheet)

Then open the .drawio file at https://app.diagrams.net (or the desktop app):
File > Open, then File > Export as > PDF.
"""
import sys
import html
import openpyxl

DOC_NAME_CELL = "D4"
STEPS_FIRST_ROW = 44
STEPS_ROWS = 16
COL = {"step": "B", "phase": "C", "activity": "D", "role": "E", "duration": "F",
       "type": "L", "branch_yes": "M", "branch_no": "N"}

LANE_WIDTH = 220
HEADER_HEIGHT = 50
TITLE_HEIGHT = 50
ROW_HEIGHT = 120
PHASE_BAND_HEIGHT = 40
NODE_W, NODE_H = 180, 74
DECISION_W, DECISION_H = 170, 90

PALETTE = [
    ("#DCEEFB", "#4472C4"), ("#FCE4D6", "#ED7D31"), ("#E2D5F1", "#7030A0"),
    ("#D9EAD3", "#38761D"), ("#F4CCCC", "#CC0000"), ("#D0E0E3", "#134F5C"), ("#FFF2CC", "#BF9000"),
]
PHASE_PALETTE = ["#E2EFDA", "#FFF2CC", "#FCE4D6", "#E2D5F1", "#D0E0E3", "#F4CCCC"]
DECISION_FILL, DECISION_STROKE = "#FFF2CC", "#BF9000"
TERMINAL_FILL, TERMINAL_STROKE = "#D9EAD3", "#38761D"
NAVY = "#1F3864"


def esc(s):
    return html.escape(str(s), quote=True)


def read_steps(xlsx_path, sheet_name=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if sheet_name:
        ws = wb[sheet_name]
    else:
        candidates = [n for n in wb.sheetnames if n not in ("ReadMe", "Procedure_BLANK")]
        ws = wb[candidates[0]]
    title = ws[DOC_NAME_CELL].value or ws.title
    steps = []
    for i in range(STEPS_ROWS):
        r = STEPS_FIRST_ROW + i
        activity = ws[f'{COL["activity"]}{r}'].value
        role = ws[f'{COL["role"]}{r}'].value
        step_no = ws[f'{COL["step"]}{r}'].value
        if not activity and not role and step_no in (None, ""):
            continue
        steps.append({
            "step_no": step_no,
            "phase": str(ws[f'{COL["phase"]}{r}'].value or "").strip(),
            "activity": str(activity or "").strip(),
            "role": str(role or "Unassigned").strip(),
            "duration": str(ws[f'{COL["duration"]}{r}'].value or "").strip(),
            "type": (str(ws[f'{COL["type"]}{r}'].value).strip() if ws[f'{COL["type"]}{r}'].value else "Process"),
            "branch_yes": ws[f'{COL["branch_yes"]}{r}'].value,
            "branch_no": ws[f'{COL["branch_no"]}{r}'].value,
        })
    return str(title), steps


def build_drawio(title, steps):
    lanes = []
    for s in steps:
        if s["role"] not in lanes:
            lanes.append(s["role"])
    lane_index = {name: i for i, name in enumerate(lanes)}
    lane_color = {name: PALETTE[i % len(PALETTE)] for i, name in enumerate(lanes)}
    by_step_no = {s["step_no"]: s for s in steps if s["step_no"] is not None}

    total_width = max(len(lanes), 1) * LANE_WIDTH
    cells = ['<mxCell id="0" />', '<mxCell id="1" parent="0" />']
    node_id = {}

    # --- pass 1: walk steps, inserting a phase band whenever the phase changes ---
    y = TITLE_HEIGHT + HEADER_HEIGHT
    current_phase = object()  # sentinel, never equals a real phase string
    phase_i = -1
    for idx, s in enumerate(steps):
        if s["phase"] and s["phase"] != current_phase:
            phase_i += 1
            fill = PHASE_PALETTE[phase_i % len(PHASE_PALETTE)]
            cells.append(
                f'<mxCell id="phase{phase_i}" value="{esc(s["phase"])}" style="rounded=0;whiteSpace=wrap;html=1;'
                f'fillColor={fill};strokeColor=#BFBFBF;fontStyle=1;fontSize=11;align=left;spacingLeft=10;verticalAlign=middle;" '
                f'vertex="1" parent="1"><mxGeometry x="0" y="{y}" width="{total_width}" height="{PHASE_BAND_HEIGHT}" as="geometry" /></mxCell>'
            )
            y += PHASE_BAND_HEIGHT
            current_phase = s["phase"]

        li = lane_index[s["role"]]
        x_lane = li * LANE_WIDTH
        step_type = s["type"].lower()
        label_txt = f'{s["step_no"]}. {s["activity"]}' if s["step_no"] is not None else s["activity"]
        if s["duration"]:
            label_txt += f'&#10;({s["duration"]})'

        if step_type == "decision":
            w, h = DECISION_W, DECISION_H
            style = f"rhombus;whiteSpace=wrap;html=1;fillColor={DECISION_FILL};strokeColor={DECISION_STROKE};fontSize=10;"
        elif step_type in ("start", "end"):
            w, h = NODE_W - 20, NODE_H - 14
            style = f"rounded=1;arcSize=50;whiteSpace=wrap;html=1;fillColor={TERMINAL_FILL};strokeColor={TERMINAL_STROKE};fontSize=10;"
        else:
            fill, stroke = lane_color[s["role"]]
            w, h = NODE_W, NODE_H
            style = f"rounded=1;whiteSpace=wrap;html=1;fillColor={fill};strokeColor={stroke};fontSize=10;"

        x = x_lane + (LANE_WIDTH - w) / 2
        node_y = y + (ROW_HEIGHT - h) / 2
        cid = f"n{idx}"
        node_id[s["step_no"]] = cid
        cells.append(
            f'<mxCell id="{cid}" value="{esc(label_txt)}" style="{style}" vertex="1" parent="1">'
            f'<mxGeometry x="{x:.1f}" y="{node_y:.1f}" width="{w}" height="{h}" as="geometry" /></mxCell>'
        )
        y += ROW_HEIGHT

    total_height = y + 20

    # lane header + column borders (added after total_height is known)
    lane_cells = []
    for i, name in enumerate(lanes):
        _, stroke = lane_color[name]
        x = i * LANE_WIDTH
        lane_cells.append(
            f'<mxCell id="lanehdr{i}" value="{esc(name)}" style="rounded=0;whiteSpace=wrap;html=1;'
            f'fillColor={stroke};strokeColor=none;fontColor=#FFFFFF;fontStyle=1;fontSize=11;align=center;verticalAlign=middle;" '
            f'vertex="1" parent="1"><mxGeometry x="{x}" y="{TITLE_HEIGHT}" width="{LANE_WIDTH}" height="{HEADER_HEIGHT}" as="geometry" /></mxCell>'
        )
        lane_cells.append(
            f'<mxCell id="lanecol{i}" value="" style="rounded=0;whiteSpace=wrap;html=1;fillColor=none;'
            f'strokeColor=#BFBFBF;dashed=1;" vertex="1" parent="1">'
            f'<mxGeometry x="{x}" y="{TITLE_HEIGHT + HEADER_HEIGHT}" width="{LANE_WIDTH}" height="{total_height - TITLE_HEIGHT - HEADER_HEIGHT}" as="geometry" /></mxCell>'
        )
    cells.append(
        f'<mxCell id="title" value="{esc(title)}" style="rounded=0;whiteSpace=wrap;html=1;'
        f'fillColor={NAVY};strokeColor=none;fontColor=#FFFFFF;fontStyle=1;fontSize=16;align=center;verticalAlign=middle;" '
        f'vertex="1" parent="1"><mxGeometry x="0" y="0" width="{total_width}" height="{TITLE_HEIGHT}" as="geometry" /></mxCell>'
    )
    cells = cells[:2] + [cells[-1]] + lane_cells + cells[2:-1]

    # edges
    edge_i = 0

    def add_edge(src, dst, label=""):
        nonlocal edge_i
        edge_i += 1
        lbl = f' value="{esc(label)}"' if label else ""
        cells.append(
            f'<mxCell id="e{edge_i}"{lbl} style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;'
            f'strokeColor=#595959;fontSize=9;fontColor=#595959;" edge="1" parent="1" source="{src}" target="{dst}">'
            f'<mxGeometry relative="1" as="geometry" /></mxCell>'
        )

    for idx, s in enumerate(steps):
        step_type = s["type"].lower()
        cid = node_id[s["step_no"]]
        if step_type == "decision":
            yes_target = by_step_no.get(s["branch_yes"])
            no_target = by_step_no.get(s["branch_no"])
            if yes_target:
                add_edge(cid, node_id[yes_target["step_no"]], "Yes")
            if no_target:
                add_edge(cid, node_id[no_target["step_no"]], "No")
        elif step_type == "end":
            pass
        else:
            if idx + 1 < len(steps):
                nxt = steps[idx + 1]
                add_edge(cid, node_id[nxt["step_no"]])

    body = "\n        ".join(cells)
    xml = f'''<mxfile host="app.diagrams.net">
  <diagram name="{esc(title)}" id="qflow-procedure">
    <mxGraphModel dx="1000" dy="700" grid="1" gridSize="10" guides="1" tooltips="1" connect="1"
        arrows="1" fold="1" page="1" pageScale="1" pageWidth="{total_width}" pageHeight="{total_height}"
        math="0" shadow="0">
      <root>
        {body}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>
'''
    return xml


def main():
    if len(sys.argv) < 2:
        print("Usage: python generate_drawio.py <SOP_Master.xlsx> [sheet name] [output.drawio]")
        sys.exit(1)
    xlsx_path = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
    out_path = sys.argv[3] if len(sys.argv) > 3 else "diagram.drawio"
    title, steps = read_steps(xlsx_path, sheet_name)
    if not steps:
        print("No steps found.")
        sys.exit(1)
    xml = build_drawio(title, steps)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(xml)
    print(f"Wrote {out_path} — {len(steps)} steps, open it at app.diagrams.net")


if __name__ == "__main__":
    main()
