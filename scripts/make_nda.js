const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell,
  WidthType, ShadingType, AlignmentType, LevelFormat } = require("docx");
const fs = require("fs");

const PAGE_W = 12240, PAGE_H = 15840;

const H = (text) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 120 }, children: [new TextRun({ text, bold: true, color: "1F3864", size: 24 })] });
const P = (text, opts={}) => new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 21, ...opts })] });
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
    new TableRow({ children: [cell("PARTY A", { bold: true, shade: "DCE6F1", width: 4680 }), cell("PARTY B", { bold: true, shade: "DCE6F1", width: 4680 })] }),
    new TableRow({ children: [
      cell("Qflow Consulting [FZ-LLC / LLC]\nTrade License No.: [•]\nAddress: [•], Abu Dhabi, United Arab Emirates", { width: 4680 }),
      cell("[Counterparty Legal Name]\nTrade License No.: [•]\nAddress: [•]", { width: 4680 }),
    ]}),
  ],
});

const doc = new Document({
  numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 420, hanging: 240 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } },
    children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "QFLOW CONSULTING", bold: true, size: 32, color: "1F3864" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "ISO & Quality Management Consulting — Abu Dhabi, UAE", size: 20, color: "595959" })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, heading: HeadingLevel.TITLE, spacing: { after: 240 }, children: [new TextRun({ text: "MUTUAL NON-DISCLOSURE AGREEMENT", bold: true, size: 30 })] }),
      P("This Mutual Non-Disclosure Agreement (\"Agreement\") is entered into on [Date] (\"Effective Date\") by and between:"),
      partiesTable,
      P("(each a \"Party\", collectively the \"Parties\"), in connection with a potential or ongoing business relationship, including without limitation ISO/Quality Management consulting services (\"Purpose\").", { italics: true }),

      H("1. Definition of Confidential Information"),
      P("\"Confidential Information\" means any non-public information disclosed by one Party (\"Disclosing Party\") to the other (\"Receiving Party\"), whether oral, written, or electronic, and whether or not marked confidential, including but not limited to:"),
      Bullet("Business processes, quality management systems, audit findings, and internal audit reports"),
      Bullet("Financial information, pricing, and business plans"),
      Bullet("Customer, supplier, and employee information"),
      Bullet("Technical data, methodologies, templates, and proprietary tools"),
      Bullet("Any information reasonably understood to be confidential given its nature or the circumstances of disclosure"),

      H("2. Exclusions"),
      P("Confidential Information does not include information that:"),
      Bullet("Was already lawfully known to the Receiving Party prior to disclosure, without a duty of confidentiality"),
      Bullet("Is or becomes publicly available through no fault of the Receiving Party"),
      Bullet("Is independently developed by the Receiving Party without use of the Confidential Information"),
      Bullet("Is lawfully received from a third party without breach of any confidentiality obligation"),
      Bullet("Is required to be disclosed by law, regulation, or court order, provided the Receiving Party gives prompt notice (where legally permitted) to allow the Disclosing Party to seek protection"),

      H("3. Obligations of the Receiving Party"),
      Bullet("Use the Confidential Information solely for the Purpose"),
      Bullet("Protect it with at least the same degree of care used for its own confidential information, and no less than reasonable care"),
      Bullet("Not disclose it to any third party without prior written consent, except to employees, officers, or advisors who need to know it for the Purpose and are bound by confidentiality obligations at least as protective as this Agreement"),
      Bullet("Not reverse-engineer, copy, or use the Confidential Information to compete with the Disclosing Party"),

      H("4. Term"),
      P("This Agreement takes effect on the Effective Date and remains in force for [2] years, unless terminated earlier by either Party on [30] days' written notice. Confidentiality obligations with respect to Confidential Information disclosed during the term survive termination for a period of [3] years, except for trade secrets, which remain protected for as long as they qualify as trade secrets under applicable law."),

      H("5. Return or Destruction of Information"),
      P("Upon written request of the Disclosing Party, or upon termination of this Agreement, the Receiving Party shall promptly return or destroy all Confidential Information and confirm such action in writing, except for archival copies retained solely to comply with legal or regulatory obligations, which remain subject to this Agreement."),

      H("6. No License or Warranty"),
      P("Nothing in this Agreement grants the Receiving Party any right, license, or interest in the Disclosing Party's Confidential Information beyond the limited use stated in Clause 3. Confidential Information is provided \"as is\" without warranty of accuracy or completeness."),

      H("7. Remedies"),
      P("The Parties acknowledge that a breach of this Agreement may cause irreparable harm for which monetary damages alone would be inadequate. The Disclosing Party is therefore entitled to seek injunctive relief, in addition to any other remedies available at law or equity, without the need to post a bond."),

      H("8. No Obligation to Proceed"),
      P("Nothing in this Agreement obligates either Party to proceed with the Purpose, enter into any further agreement, or continue discussions, and either Party may terminate discussions at any time."),

      H("9. Governing Law & Dispute Resolution"),
      P("This Agreement is governed by the laws of the United Arab Emirates [/ DIFC, as applicable]. Any dispute arising from this Agreement shall first be addressed through good-faith negotiation; failing resolution within [30] days, it shall be referred to [arbitration under the ADCCAC Rules seated in Abu Dhabi / the competent courts of Abu Dhabi], as elected by the Parties."),

      H("10. General"),
      Bullet("This Agreement constitutes the entire understanding between the Parties regarding confidentiality and supersedes any prior discussions on the same subject."),
      Bullet("Neither Party may assign this Agreement without the other's prior written consent."),
      Bullet("If any provision is held invalid, the remaining provisions remain in full force."),
      Bullet("This Agreement may be executed in counterparts, including electronic signature."),

      new Paragraph({ spacing: { before: 400, after: 200 }, children: [new TextRun({ text: "SIGNATURES", bold: true, size: 22, color: "1F3864" })] }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [4680, 4680],
        rows: [
          new TableRow({ children: [cell("For and on behalf of Qflow Consulting", { bold: true, width: 4680 }), cell("For and on behalf of [Counterparty]", { bold: true, width: 4680 })] }),
          new TableRow({ children: [cell("\n\nName: [•]\nTitle: [•]\nSignature: ____________________\nDate: [•]", { width: 4680 }), cell("\n\nName: [•]\nTitle: [•]\nSignature: ____________________\nDate: [•]", { width: 4680 })] }),
        ],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(process.argv[2] || "NDA.docx", buf);
  console.log("written");
});
