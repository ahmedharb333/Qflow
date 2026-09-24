/**
 * Qflow Consulting — SOP Master workbook tools.
 *
 * Menu:
 *   - New SOP for Department/Process...   -> duplicates the hidden "Procedure_BLANK"
 *                                             tab as "Procedure - <name>"
 *   - Generate SOP Document (Doc + PDF)    -> writes the full narrative SOP
 *                                             (purpose, scope, definitions, references,
 *                                             policy, phased procedure steps, RACI,
 *                                             escalation, revision history)
 *   - Generate Swimlane Diagram (.drawio)  -> Appendix-A style diagram: columns per
 *                                             role, colored phase bands, decision
 *                                             diamonds, branching arrows
 *
 * INSTALL (one paste):
 * 1. Extensions > Apps Script.
 * 2. Delete any starter code, paste this entire file, save (Ctrl+S).
 * 3. Close the Apps Script tab, reload the Google Sheet.
 * 4. Menu "Qflow Tools" appears next to Help.
 *
 * Both generators act on the ACTIVE tab — open the "Procedure - X" tab you want
 * before running them.
 */

// ---- fixed row/col map — must match make_sop_master.py exactly ----
const DOC_NAME_CELL = "D4", DOC_NO_CELL = "H4", VERSION_CELL = "D5", DEPT_CELL = "H5";
const ISSUE_DATE_CELL = "D6", EFF_DATE_CELL = "H6", REVIEW_DATE_CELL = "D7", OWNER_CELL = "H7";
const APPROVED_BY_CELL = "D8", STANDARD_CELL = "H8";
const PURPOSE_CELL = "B11", SCOPE_CELL = "B14", EXCEPT_CELL = "B17";
const DEFS_FIRST = 21, DEFS_ROWS = 6;
const REFS_FIRST = 29, REFS_ROWS = 4;
const POLICY_FIRST = 35, POLICY_ROWS = 6;
const STEPS_FIRST = 44, STEPS_ROWS = 16;
const ESC_FIRST = 63, ESC_ROWS = 4;
const REV_FIRST = 70, REV_ROWS = 3;
const SCOL = { step: 2, phase: 3, activity: 4, role: 5, duration: 6, input: 7, output: 8,
               tool: 9, reference: 10, notes: 11, type: 12, branchYes: 13, branchNo: 14,
               consulted: 15, informed: 16 }; // B..P

const RESERVED_SHEETS = ["ReadMe", "Procedure_BLANK"];

const LANE_WIDTH = 220, HEADER_HEIGHT = 50, TITLE_HEIGHT = 50, ROW_HEIGHT = 120, PHASE_BAND_HEIGHT = 40;
const NODE_W = 180, NODE_H = 74, DEC_W = 170, DEC_H = 90;
const PALETTE = [
  ["#DCEEFB", "#4472C4"], ["#FCE4D6", "#ED7D31"], ["#E2D5F1", "#7030A0"],
  ["#D9EAD3", "#38761D"], ["#F4CCCC", "#CC0000"], ["#D0E0E3", "#134F5C"], ["#FFF2CC", "#BF9000"]
];
const PHASE_PALETTE = ["#E2EFDA", "#FFF2CC", "#FCE4D6", "#E2D5F1", "#D0E0E3", "#F4CCCC"];
const DECISION_FILL = "#FFF2CC", DECISION_STROKE = "#BF9000";
const TERMINAL_FILL = "#D9EAD3", TERMINAL_STROKE = "#38761D";
const NAVY = "#1F3864";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Qflow Tools")
    .addItem("New SOP for Department/Process...", "newDepartmentProcedure")
    .addSeparator()
    .addItem("1. Generate SOP Document (Doc + PDF)", "generateSopDocument")
    .addItem("2. Generate Swimlane Diagram (.drawio)", "generateDrawio")
    .addToUi();
}

function esc_(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// ---------------------------------------------------------------------------
// New department/process tab
// ---------------------------------------------------------------------------
function newDepartmentProcedure() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "New SOP",
    "Department / process name (e.g. Recruitment, Termination, Training, Customer Service):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const name = resp.getResponseText().trim();
  if (!name) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blank = ss.getSheetByName("Procedure_BLANK");
  if (!blank) { ui.alert('Template sheet "Procedure_BLANK" not found — was it renamed or deleted?'); return; }

  let sheetName = ("Procedure - " + name).substring(0, 100);
  if (ss.getSheetByName(sheetName)) { ui.alert('A tab named "' + sheetName + '" already exists.'); return; }

  const newSheet = blank.copyTo(ss);
  newSheet.setName(sheetName);
  newSheet.showSheet();
  newSheet.getRange(DEPT_CELL).setValue(name);
  newSheet.getRange(DOC_NAME_CELL).setValue(name + " SOP");
  ss.setActiveSheet(newSheet);
  ui.alert('Created "' + sheetName + '". Fill in Purpose, Scope, Definitions, Policy and the step table, ' +
           "then use Qflow Tools to generate the SOP document or swimlane diagram.");
}

function getActiveSopSheet_() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (RESERVED_SHEETS.indexOf(sheet.getName()) !== -1) {
    SpreadsheetApp.getUi().alert('Open a real SOP tab first (not "' + sheet.getName() + '").');
    return null;
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// Shared readers
// ---------------------------------------------------------------------------
function readHeader_(sheet) {
  return {
    docName: sheet.getRange(DOC_NAME_CELL).getValue() || sheet.getName(),
    docNo: sheet.getRange(DOC_NO_CELL).getValue(),
    version: sheet.getRange(VERSION_CELL).getValue(),
    department: sheet.getRange(DEPT_CELL).getValue(),
    issueDate: sheet.getRange(ISSUE_DATE_CELL).getValue(),
    effDate: sheet.getRange(EFF_DATE_CELL).getValue(),
    reviewDate: sheet.getRange(REVIEW_DATE_CELL).getValue(),
    owner: sheet.getRange(OWNER_CELL).getValue(),
    approvedBy: sheet.getRange(APPROVED_BY_CELL).getValue(),
    standard: sheet.getRange(STANDARD_CELL).getValue(),
    purpose: sheet.getRange(PURPOSE_CELL).getValue(),
    scope: sheet.getRange(SCOPE_CELL).getValue(),
    exceptions: sheet.getRange(EXCEPT_CELL).getValue(),
  };
}

function readTableRows_(sheet, firstRow, rowCount, colStart, colCount) {
  const vals = sheet.getRange(firstRow, colStart, rowCount, colCount).getValues();
  return vals.filter((row) => row.some((v) => String(v).trim() !== ""));
}

function readSteps_(sheet) {
  const steps = [];
  for (let i = 0; i < STEPS_ROWS; i++) {
    const r = STEPS_FIRST + i;
    const activity = sheet.getRange(r, SCOL.activity).getValue();
    const role = sheet.getRange(r, SCOL.role).getValue();
    const stepNo = sheet.getRange(r, SCOL.step).getValue();
    if (!activity && !role && (stepNo === "" || stepNo === null)) continue;
    steps.push({
      stepNo: stepNo,
      phase: String(sheet.getRange(r, SCOL.phase).getValue() || "").trim(),
      activity: String(activity || "").trim(),
      role: String(role || "Unassigned").trim(),
      duration: String(sheet.getRange(r, SCOL.duration).getValue() || "").trim(),
      input: String(sheet.getRange(r, SCOL.input).getValue() || "").trim(),
      output: String(sheet.getRange(r, SCOL.output).getValue() || "").trim(),
      tool: String(sheet.getRange(r, SCOL.tool).getValue() || "").trim(),
      reference: String(sheet.getRange(r, SCOL.reference).getValue() || "").trim(),
      notes: String(sheet.getRange(r, SCOL.notes).getValue() || "").trim(),
      type: String(sheet.getRange(r, SCOL.type).getValue() || "Process").trim() || "Process",
      branchYes: sheet.getRange(r, SCOL.branchYes).getValue(),
      branchNo: sheet.getRange(r, SCOL.branchNo).getValue(),
      consulted: String(sheet.getRange(r, SCOL.consulted).getValue() || "").trim(),
      informed: String(sheet.getRange(r, SCOL.informed).getValue() || "").trim(),
    });
  }
  return steps;
}

function driveFolderForSpreadsheet_(ss) {
  const file = DriveApp.getFileById(ss.getId());
  const parents = file.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

// ---------------------------------------------------------------------------
// 1. SOP written document
// ---------------------------------------------------------------------------
function generateSopDocument() {
  const sheet = getActiveSopSheet_();
  if (!sheet) return;
  const h = readHeader_(sheet);
  const defs = readTableRows_(sheet, DEFS_FIRST, DEFS_ROWS, 3, 2);       // C:D
  const refs = readTableRows_(sheet, REFS_FIRST, REFS_ROWS, 2, 1);       // B
  const policy = readTableRows_(sheet, POLICY_FIRST, POLICY_ROWS, 2, 1); // B
  const steps = readSteps_(sheet);
  const escalation = readTableRows_(sheet, ESC_FIRST, ESC_ROWS, 2, 4);   // B,F,J,L approx (merged, read raw anchor cols)
  const revision = readTableRows_(sheet, REV_FIRST, REV_ROWS, 2, 4);    // B,D,F,M approx

  if (steps.length === 0) {
    SpreadsheetApp.getUi().alert("No procedure steps found. Fill in the step table first.");
    return;
  }

  const doc = DocumentApp.create(h.docName + " - SOP");
  const body = doc.getBody();
  body.setMarginTop(50).setMarginBottom(50).setMarginLeft(60).setMarginRight(60);

  body.appendParagraph(h.docName).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph("Qflow Consulting — ISO & Quality Management").setItalic(true);

  const headerTable = body.appendTable([
    ["Document No.", String(h.docNo || ""), "Version", String(h.version || "")],
    ["Department / Process", String(h.department || ""), "Owner", String(h.owner || "")],
    ["Issue Date", fmtDate_(h.issueDate), "Effective Date", fmtDate_(h.effDate)],
    ["Next Review Date", fmtDate_(h.reviewDate), "Approved By", String(h.approvedBy || "")],
    ["Applicable Standard/Clause", String(h.standard || ""), "", ""],
  ]);
  styleTable_(headerTable, [0, 2]);

  addHeading_(body, "1. Purpose");
  body.appendParagraph(String(h.purpose || ""));

  addHeading_(body, "2. Scope");
  body.appendParagraph(String(h.scope || ""));
  if (h.exceptions) {
    body.appendParagraph("Exceptions: " + String(h.exceptions));
  }

  if (defs.length) {
    addHeading_(body, "3. Definitions & Terms");
    const dTable = body.appendTable([["Term", "Definition"]].concat(defs.map((r) => [String(r[0]), String(r[1])])));
    styleTable_(dTable, [0]);
  }

  if (refs.length) {
    addHeading_(body, "4. References & Related Documents");
    refs.forEach((r) => body.appendListItem(String(r[0])).setGlyphType(DocumentApp.GlyphType.BULLET));
  }

  if (policy.length) {
    addHeading_(body, "5. Policy & Governing Principles");
    policy.forEach((r) => body.appendListItem(String(r[0])).setGlyphType(DocumentApp.GlyphType.BULLET));
  }

  addHeading_(body, "6. Procedure");
  let currentPhase = null;
  steps.forEach((s) => {
    if (s.phase && s.phase !== currentPhase) {
      body.appendParagraph(s.phase).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      currentPhase = s.phase;
    }
    const durSuffix = s.duration ? " (" + s.duration + ")" : "";
    if (String(s.type).toLowerCase() === "decision") {
      const p = body.appendParagraph(s.stepNo + ". [" + s.role + "] Gate: " + s.activity + durSuffix);
      p.editAsText().setBold(0, String(s.stepNo).length, true);
      if (s.branchYes !== "" && s.branchYes !== null) {
        body.appendListItem("If Yes -> go to step " + s.branchYes).setNestingLevel(0).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
      }
      if (s.branchNo !== "" && s.branchNo !== null) {
        body.appendListItem("If No -> go to step " + s.branchNo).setNestingLevel(0).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET);
      }
    } else {
      const tag = String(s.type).toLowerCase() === "end" ? " (end of procedure)" : "";
      body.appendParagraph(s.stepNo + ". [" + s.role + "] " + s.activity + durSuffix + tag);
    }
  });

  // RACI — Responsible role = R/A by default, plus Consulted/Informed from the sheet
  const roles = [];
  steps.forEach((s) => { if (roles.indexOf(s.role) === -1) roles.push(s.role); });
  if (roles.length) {
    addHeading_(body, "7. Roles & Responsibilities (RACI)");
    const raciHeader = ["Activity"].concat(roles);
    const raciRows = steps.map((s) => {
      const row = [s.stepNo + ". " + s.activity];
      roles.forEach((role) => {
        let mark = "";
        if (role === s.role) mark = "R/A";
        else if (s.consulted && s.consulted.toLowerCase().indexOf(role.toLowerCase()) !== -1) mark = "C";
        else if (s.informed && s.informed.toLowerCase().indexOf(role.toLowerCase()) !== -1) mark = "I";
        row.push(mark);
      });
      return row;
    });
    const raciTable = body.appendTable([raciHeader].concat(raciRows));
    styleTable_(raciTable, [0]);
    body.appendParagraph("R = Responsible, A = Accountable, C = Consulted, I = Informed.").setItalic(true);
  }

  if (escalation.length) {
    addHeading_(body, "8. Escalation & Exception Handling");
    const eTable = body.appendTable(
      [["Condition", "Action / Resolution", "Timeframe / SLA", "Escalate To"]].concat(
        escalation.map((r) => [String(r[0]), String(r[1]), String(r[2]), String(r[3])])
      )
    );
    styleTable_(eTable, [0]);
  }

  addHeading_(body, "9. Revision History");
  const revRows = revision.length ? revision : [["", "", "", ""]];
  const rTable = body.appendTable(
    [["Version", "Date", "Description of Change", "Approved By"]].concat(
      revRows.map((r) => [String(r[0]), fmtDate_(r[1]), String(r[2]), String(r[3])])
    )
  );
  styleTable_(rTable, [0]);

  addHeading_(body, "Appendix A: Process Swimlane Diagram");
  body.appendParagraph("Generate via Qflow Tools > Generate Swimlane Diagram (.drawio), export as PDF, and insert/attach here.").setItalic(true);

  doc.saveAndClose();

  const folder = driveFolderForSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet());
  const docFile = DriveApp.getFileById(doc.getId());
  folder.addFile(docFile);
  DriveApp.getRootFolder().removeFile(docFile);

  const pdfBlob = docFile.getAs("application/pdf").setName(h.docName + " - SOP.pdf");
  const pdfFile = folder.createFile(pdfBlob);

  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">SOP document created from ${steps.length} step(s).</p>
      <p><a href="${docFile.getUrl()}" target="_blank" style="display:inline-block;background:#1F3864;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open Google Doc</a></p>
      <p><a href="${pdfFile.getUrl()}" target="_blank" style="display:inline-block;background:#14B8A6;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open PDF (Drive)</a></p>
      <p style="font-size:11px;color:#666;">Both saved in the same Drive folder as this spreadsheet.</p>
    </div>`;
  const output = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(output, "SOP Document Ready");
}

function addHeading_(body, text) {
  body.appendParagraph(text).setHeading(DocumentApp.ParagraphHeading.HEADING1);
}

function styleTable_(table, boldCols) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(6).setPaddingRight(6);
      if (r === 0) {
        cell.setBackgroundColor("#1F3864");
        cell.editAsText().setForegroundColor("#FFFFFF").setBold(true).setFontSize(10);
      } else {
        cell.editAsText().setFontSize(10);
        if (boldCols && boldCols.indexOf(c) !== -1) cell.editAsText().setBold(true);
      }
    }
  }
}

function fmtDate_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]") {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd-MMM-yyyy");
  }
  return String(v);
}

// ---------------------------------------------------------------------------
// 2. Swimlane diagram (.drawio)
// ---------------------------------------------------------------------------
function buildDrawioXml_(title, steps) {
  const lanes = [];
  steps.forEach((s) => { if (lanes.indexOf(s.role) === -1) lanes.push(s.role); });
  const laneIndex = {};
  lanes.forEach((n, i) => (laneIndex[n] = i));
  const laneColor = {};
  lanes.forEach((n, i) => (laneColor[n] = PALETTE[i % PALETTE.length]));
  const byStepNo = {};
  steps.forEach((s) => { if (s.stepNo !== "" && s.stepNo !== null) byStepNo[s.stepNo] = s; });

  const totalWidth = Math.max(lanes.length, 1) * LANE_WIDTH;
  const cells = [];
  const nodeId = {};

  let y = TITLE_HEIGHT + HEADER_HEIGHT;
  let currentPhase = null;
  let phaseI = -1;
  steps.forEach((s, idx) => {
    if (s.phase && s.phase !== currentPhase) {
      phaseI++;
      const fill = PHASE_PALETTE[phaseI % PHASE_PALETTE.length];
      cells.push(
        `<mxCell id="phase${phaseI}" value="${esc_(s.phase)}" style="rounded=0;whiteSpace=wrap;html=1;fillColor=${fill};strokeColor=#BFBFBF;fontStyle=1;fontSize=11;align=left;spacingLeft=10;verticalAlign=middle;" vertex="1" parent="1"><mxGeometry x="0" y="${y}" width="${totalWidth}" height="${PHASE_BAND_HEIGHT}" as="geometry" /></mxCell>`
      );
      y += PHASE_BAND_HEIGHT;
      currentPhase = s.phase;
    }

    const li = laneIndex[s.role];
    const xLane = li * LANE_WIDTH;
    const type = s.type.toLowerCase();
    let label = s.stepNo !== "" && s.stepNo !== null ? `${s.stepNo}. ${s.activity}` : s.activity;
    if (s.duration) label += `&#10;(${s.duration})`;

    let w, h, style;
    if (type === "decision") {
      w = DEC_W; h = DEC_H;
      style = `rhombus;whiteSpace=wrap;html=1;fillColor=${DECISION_FILL};strokeColor=${DECISION_STROKE};fontSize=10;`;
    } else if (type === "start" || type === "end") {
      w = NODE_W - 20; h = NODE_H - 14;
      style = `rounded=1;arcSize=50;whiteSpace=wrap;html=1;fillColor=${TERMINAL_FILL};strokeColor=${TERMINAL_STROKE};fontSize=10;`;
    } else {
      const [fill, stroke] = laneColor[s.role];
      w = NODE_W; h = NODE_H;
      style = `rounded=1;whiteSpace=wrap;html=1;fillColor=${fill};strokeColor=${stroke};fontSize=10;`;
    }
    const x = xLane + (LANE_WIDTH - w) / 2;
    const nodeY = y + (ROW_HEIGHT - h) / 2;
    const cid = "n" + idx;
    nodeId[s.stepNo] = cid;
    cells.push(
      `<mxCell id="${cid}" value="${esc_(label)}" style="${style}" vertex="1" parent="1"><mxGeometry x="${x.toFixed(1)}" y="${nodeY.toFixed(1)}" width="${w}" height="${h}" as="geometry" /></mxCell>`
    );
    y += ROW_HEIGHT;
  });
  const totalHeight = y + 20;

  const laneCells = [];
  lanes.forEach((name, i) => {
    const [, stroke] = laneColor[name];
    const x = i * LANE_WIDTH;
    laneCells.push(
      `<mxCell id="lanehdr${i}" value="${esc_(name)}" style="rounded=0;whiteSpace=wrap;html=1;fillColor=${stroke};strokeColor=none;fontColor=#FFFFFF;fontStyle=1;fontSize=11;align=center;verticalAlign=middle;" vertex="1" parent="1"><mxGeometry x="${x}" y="${TITLE_HEIGHT}" width="${LANE_WIDTH}" height="${HEADER_HEIGHT}" as="geometry" /></mxCell>`
    );
    laneCells.push(
      `<mxCell id="lanecol${i}" value="" style="rounded=0;whiteSpace=wrap;html=1;fillColor=none;strokeColor=#BFBFBF;dashed=1;" vertex="1" parent="1"><mxGeometry x="${x}" y="${TITLE_HEIGHT + HEADER_HEIGHT}" width="${LANE_WIDTH}" height="${totalHeight - TITLE_HEIGHT - HEADER_HEIGHT}" as="geometry" /></mxCell>`
    );
  });
  const titleCell = `<mxCell id="title" value="${esc_(title)}" style="rounded=0;whiteSpace=wrap;html=1;fillColor=${NAVY};strokeColor=none;fontColor=#FFFFFF;fontStyle=1;fontSize=16;align=center;verticalAlign=middle;" vertex="1" parent="1"><mxGeometry x="0" y="0" width="${totalWidth}" height="${TITLE_HEIGHT}" as="geometry" /></mxCell>`;

  let edgeI = 0;
  const edgeCells = [];
  function addEdge(src, dst, label) {
    edgeI++;
    const val = label ? ` value="${esc_(label)}"` : "";
    edgeCells.push(
      `<mxCell id="e${edgeI}"${val} style="edgeStyle=orthogonalEdgeStyle;rounded=0;html=1;strokeColor=#595959;fontSize=9;fontColor=#595959;" edge="1" parent="1" source="${src}" target="${dst}"><mxGeometry relative="1" as="geometry" /></mxCell>`
    );
  }
  steps.forEach((s, idx) => {
    const type = s.type.toLowerCase();
    const cid = nodeId[s.stepNo];
    if (type === "decision") {
      const yesTarget = byStepNo[s.branchYes];
      const noTarget = byStepNo[s.branchNo];
      if (yesTarget) addEdge(cid, nodeId[yesTarget.stepNo], "Yes");
      if (noTarget) addEdge(cid, nodeId[noTarget.stepNo], "No");
    } else if (type === "end") {
      // no outgoing edge
    } else if (idx + 1 < steps.length) {
      addEdge(cid, nodeId[steps[idx + 1].stepNo]);
    }
  });

  const body = ['<mxCell id="0" />', '<mxCell id="1" parent="0" />', titleCell]
    .concat(laneCells).concat(cells).concat(edgeCells).join("\n        ");

  return `<mxfile host="app.diagrams.net">
  <diagram name="${esc_(title)}" id="qflow-procedure">
    <mxGraphModel dx="1000" dy="700" grid="1" gridSize="10" guides="1" tooltips="1" connect="1" arrows="1" fold="1" page="1" pageScale="1" pageWidth="${totalWidth}" pageHeight="${totalHeight}" math="0" shadow="0">
      <root>
        ${body}
      </root>
    </mxGraphModel>
  </diagram>
</mxfile>`;
}

function generateDrawio() {
  const sheet = getActiveSopSheet_();
  if (!sheet) return;
  const h = readHeader_(sheet);
  const steps = readSteps_(sheet);
  if (steps.length === 0) {
    SpreadsheetApp.getUi().alert("No steps found. Fill in the step table first.");
    return;
  }
  const xml = buildDrawioXml_(h.docName || sheet.getName(), steps);
  const fileName = String(h.docName || sheet.getName()).replace(/[^\w\- ]+/g, "").trim() + ".drawio";

  const folder = driveFolderForSpreadsheet_(SpreadsheetApp.getActiveSpreadsheet());
  const blob = Utilities.newBlob(xml, "application/xml", fileName);
  const file = folder.createFile(blob);

  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">Diagram built from ${steps.length} step(s), ${new Set(steps.map(s=>s.role)).size} lane(s).</p>
      <p><a href="${file.getUrl()}" target="_blank" style="display:inline-block;background:#1F3864;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open ${esc_(fileName)} in Drive</a></p>
      <p style="font-size:12px;color:#666;margin-top:10px;">
        From that Drive preview, use the Download icon (top-right) to save the .drawio file, then open it at
        <a href="https://app.diagrams.net" target="_blank">app.diagrams.net</a> (File &gt; Open from &gt; Device)
        and File &gt; Export as &gt; PDF.
      </p>
    </div>`;
  const output = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(output, "Swimlane Diagram Ready");
}
