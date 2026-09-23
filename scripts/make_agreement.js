const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, PageOrientation, LevelFormat, convertInchesToTwip } = require("docx");
const fs = require("fs");

const PAGE_W = 12240, PAGE_H = 15840; // US Letter

const H = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, children: [new TextRun({ text, bold: true, color: "1F3864", size: 26 })] });
const H2 = (text) => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 100 }, children: [new TextRun({ text, bold: true, color: "2E5395", size: 22 })] });
const P = (text, opts={}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 21, ...opts })] });
const Bold = (text) => new TextRun({ text, bold: true, size: 21 });
const Bullet = (text) => new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 }, children: [new TextRun({ text, size: 21 })] });

function cell(text, opts={}) {
  return new TableCell({
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shade ? { type: ShadingType.CLEAR, fill: opts.shade } : undefined,
    children: [new Paragraph({ children: [new TextRun({ text, bold: !!opts.bold, size: 20 })] })],
  });
}

const partiesTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [4680, 4680],
  rows: [
    new TableRow({ children: [cell("CONSULTANT", { bold: true, shade: "DCE6F1", width: 4680 }), cell("CLIENT", { bold: true, shade: "DCE6F1", width: 4680 })] }),
    new TableRow({ children: [
      cell("Qflow Consulting [FZ-LLC / LLC — insert legal form]\nTrade License No.: [•]\nAddress: [•], Abu Dhabi, United Arab Emirates\nRepresented by: [•], [Title]", { width: 4680 }),
      cell("[Client Legal Name]\nTrade License No.: [•]\nAddress: [•]\nRepresented by: [•], [Title]", { width: 4680 }),
    ]}),
  ],
});

const scopeTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3120, 6240],
  rows: [
    new TableRow({ children: [cell("Service", { bold: true, shade: "DCE6F1", width: 3120 }), cell("Description", { bold: true, shade: "DCE6F1", width: 6240 })] }),
    new TableRow({ children: [cell("Audit Readiness", { width: 3120 }), cell("Gap assessment, mock audit, corrective action plan prior to certification/surveillance audit.", { width: 6240 })] }),
    new TableRow({ children: [cell("Internal Audit Program", { width: 3120 }), cell("Design of audit schedule, checklists, and execution of internal audits per agreed scope.", { width: 6240 })] }),
    new TableRow({ children: [cell("CAPA Support", { width: 3120 }), cell("Root cause analysis facilitation and CAPA process design/review.", { width: 6240 })] }),
    new TableRow({ children: [cell("Documentation & Process Control", { width: 3120 }), cell("Development/review of procedures, work instructions, and document control system.", { width: 6240 })] }),
    new TableRow({ children: [cell("ISO Standard(s) in scope", { width: 3120 }), cell("[ISO 9001 / ISO 27001 / other — specify]", { width: 6240 })] }),
    new TableRow({ children: [cell("Other", { width: 3120 }), cell("[•]", { width: 6240 })] }),
  ],
});

const feeTable = new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [3120, 3120, 3120],
  rows: [
    new TableRow({ children: [cell("Item", { bold: true, shade: "DCE6F1", width: 3120 }), cell("Basis", { bold: true, shade: "DCE6F1", width: 3120 }), cell("Amount (AED)", { bold: true, shade: "DCE6F1", width: 3120 })] }),
    new TableRow({ children: [cell("Professional Fee", { width: 3120 }), cell("[Fixed project fee / Monthly retainer / Daily rate]", { width: 3120 }), cell("[•]", { width: 3120 })] }),
    new TableRow({ children: [cell("VAT (5%)", { width: 3120 }), cell("As applicable under UAE VAT law", { width: 3120 }), cell("[•]", { width: 3120 })] }),
    new TableRow({ children: [cell("Expenses", { width: 3120 }), cell("Reimbursed at cost, pre-approved, with receipts", { width: 3120 }), cell("At cost", { width: 3120 })] }),
    new TableRow({ children: [cell("Payment Terms", { width: 3120 }), cell("[•] % upon signature, balance on [milestone/net 15 days from invoice]", { width: 3120 }), cell("—", { width: 3120 })] }),
  ],
});

const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "QFLOW CONSULTING", bold: true, size: 32, color: "1F3864" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "ISO & Quality Management Consulting — Abu Dhabi, UAE", size: 20, color: "595959" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.TITLE, spacing: { after: 240 }, children: [new TextRun({ text: "CONSULTING SERVICES AGREEMENT", bold: true, size: 30 })] }),
      P("This Consulting Services Agreement (\"Agreement\") is entered into on [Date] (\"Effective Date\") by and between:"),
      partiesTable,
      P("(each a \"Party\", collectively the \"Parties\")", { italics: true }),

      H("1. Scope of Services"),
      P("The Consultant shall provide the following services (\"Services\") to the Client:"),
      scopeTable,
      P("Any service not listed above requires a signed Change Order per Clause 8 before work begins.", { italics: true }),

      H("2. Term"),
      P("This Agreement commences on the Effective Date and continues until [completion of the Services / for a period of [•] months], unless terminated earlier under Clause 9."),

      H("3. Fees & Payment"),
      feeTable,
      P("Invoices are payable within [15] days of the invoice date. Late payment accrues interest at [•]% per month on overdue amounts. Reasonable pre-approved expenses (travel, accommodation, third-party tools) are reimbursed at cost against receipts."),

      H("4. Client Responsibilities"),
      Bullet("Provide timely access to relevant personnel, records, systems, and facilities required for the Services."),
      Bullet("Designate a single point of contact for day-to-day coordination."),
      Bullet("Review and respond to Consultant deliverables within [5] business days, unless otherwise agreed."),
      Bullet("Ensure the accuracy of information, documents, and data supplied to the Consultant."),

      H("5. Deliverables & Acceptance"),
      P("Deliverables are listed in Clause 1 / attached Statement of Work. A deliverable is deemed accepted if the Client raises no written objection within [10] business days of delivery."),

      H("6. Confidentiality"),
      P("Each Party shall keep confidential all non-public information disclosed by the other Party in connection with this Agreement and use it solely to perform its obligations hereunder. This obligation survives termination of this Agreement for [3] years. This clause is supplemented, where executed, by a separate Non-Disclosure Agreement between the Parties."),

      H("7. Intellectual Property"),
      P("Pre-existing Consultant materials (methodologies, templates, checklists, tools) remain the Consultant's property. Upon full payment, the Client receives a non-exclusive, perpetual license to use deliverables created specifically for the Client for its internal business purposes. The Consultant retains the right to reuse generic, non-client-identifying methodologies and know-how developed during the engagement."),

      H("8. Change Orders"),
      P("Any change to scope, timeline, or fees must be documented in a written Change Order signed by both Parties before the change takes effect."),

      H("9. Term & Termination"),
      Bullet("Either Party may terminate for convenience with [30] days' written notice."),
      Bullet("Either Party may terminate immediately for material breach not cured within [15] days of written notice."),
      Bullet("Upon termination, the Client pays for Services performed and expenses incurred up to the termination date."),

      H("10. Limitation of Liability"),
      P("Neither Party's aggregate liability under this Agreement shall exceed the total fees paid by the Client in the [6] months preceding the claim. Neither Party is liable for indirect, incidental, or consequential damages. Nothing in this clause limits liability that cannot be limited under applicable law."),

      H("11. Independent Contractor"),
      P("The Consultant performs the Services as an independent contractor, not as an employee, agent, or partner of the Client. Nothing in this Agreement creates a joint venture or partnership."),

      H("12. Governing Law & Dispute Resolution"),
      P("This Agreement is governed by the laws of the United Arab Emirates [/ DIFC, as applicable]. The Parties shall first attempt to resolve any dispute through good-faith negotiation; failing resolution within [30] days, the dispute shall be referred to [arbitration under the ADCCAC Rules seated in Abu Dhabi / the competent courts of Abu Dhabi], as elected by the Parties."),

      H("13. General"),
      Bullet("Notices shall be in writing and delivered to the addresses stated above (or by email with read confirmation)."),
      Bullet("This Agreement, together with any signed Statement of Work and NDA, constitutes the entire agreement between the Parties and supersedes all prior discussions."),
      Bullet("If any provision is held invalid, the remaining provisions remain in full force."),
      Bullet("This Agreement may be executed in counterparts, including electronic signature."),

      new Paragraph({ spacing: { before: 400, after: 200 }, children: [new TextRun({ text: "SIGNATURES", bold: true, size: 22, color: "1F3864" })] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [cell("For and on behalf of Qflow Consulting", { bold: true, width: 4680 }), cell("For and on behalf of the Client", { bold: true, width: 4680 })] }),
          new TableRow({ children: [cell("\n\nName: [•]\nTitle: [•]\nSignature: ____________________\nDate: [•]", { width: 4680 }), cell("\n\nName: [•]\nTitle: [•]\nSignature: ____________________\nDate: [•]", { width: 4680 })] }),
        ],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "Consulting_Services_Agreement.docx", buf);
  console.log("written");
});
