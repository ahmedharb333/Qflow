# Qflow Consulting

Launch package for Qflow Consulting — an Abu Dhabi (UAE) consultancy specializing in ISO & Quality Management (Audit Readiness, Internal Audit, CAPA, Documentation & Process Control, ISO 9001 / 27001).

## Contents

| Folder | What's in it |
|---|---|
| `LinkedIn/` | Company Page setup guide + a 14-day launch content calendar |
| `Contracts/` | Consulting Services Agreement + NDA (`.docx`, UAE/Abu Dhabi law, ADCCAC arbitration) |
| `Templates/` | Working templates — see below |
| `Logo/` | Logo mark (300px/1000px, transparent + white), full lockup, LinkedIn cover banner |
| `scripts/` | Source generators for every deliverable above (Python/openpyxl, Node/docx-js, PIL) |

### Templates/

- **`Qflow_SOP_Master.xlsx`** — master SOP workbook. One tab per procedure (`Procedure - <Department>`), each with document control, purpose/scope, definitions, references, policy, a phased step table (with decision gates and RACI columns), escalation rules, and revision history.
  - `Apps_Script_Swimlane_Generator.gs` — paste into Extensions > Apps Script. Adds a **Qflow Tools** menu: create a new department SOP tab, generate the written SOP (Google Doc + PDF), and generate a swimlane diagram (`.drawio`) with colored phase bands and decision diamonds — matching a standard drawio.net swimlane, ready for `File > Export as > PDF`.
  - `generate_drawio.py` — same diagram generator, runnable locally against a downloaded copy of the workbook if you'd rather not use Apps Script.
  - `Recruitment_Demo.drawio` — sample output (15-step Recruitment & Onboarding SOP with a decision gate) you can open directly at [app.diagrams.net](https://app.diagrams.net).
- **`Qflow_Invoice_Template.xlsx`** — invoice with a hidden `Config` tab (auto-incrementing invoice number, prefix) and an `Invoice Log` tab.
  - `Invoice_Tools.gs` — paste into Extensions > Apps Script. Adds a **Qflow Invoice Tools** menu: assign the next invoice number, generate a PDF (auto-filed into `Invoices/<Client Name>/` in Drive and logged), insert the company stamp, email the client via Gmail, and mark an invoice paid.
  - `Qflow_Company_Stamp.png` — upload into the same Drive folder as the invoice spreadsheet; "Insert Company Stamp" looks it up by that filename.
- **`Qflow_Process_Template.xlsx`** — lightweight generic process-map template (inputs/outputs, steps, revision history) for processes that don't need the full SOP structure.

**Note on e-signature:** the stamp menu item is a visual stamp only, not a legally binding signature. For real e-signing, run the generated PDF through DocuSign or Dropbox Sign.

## Using the SOP / Invoice Apps Script tools

1. Open the `.xlsx` file in Google Sheets (import if needed — the tools only run on native Sheets, not on an un-converted `.xlsx` sitting in Drive).
2. Extensions > Apps Script, delete the starter code, paste the matching `.gs` file, save.
3. Reload the spreadsheet — the new menu appears next to Help.
4. First run of any Drive/Gmail action prompts a one-time Google authorization — approve it (it's your own script acting on your own account).

## Regenerating deliverables

Everything in this package is generated from `scripts/`, not edited by hand — that's the point of keeping this in git. To change wording, branding, or layout, edit the relevant script and re-run it, rather than editing the `.docx`/`.xlsx`/`.png` output directly.

```bash
cd scripts
npm install docx            # for make_agreement.js / make_nda.js
python -m pip install openpyxl pillow matplotlib   # for the .py generators

node make_agreement.js "../Contracts/Qflow_Consulting_Services_Agreement.docx"
node make_nda.js "../Contracts/Qflow_Consulting_NDA.docx"
python make_sop_master.py    # -> Qflow_SOP_Master.xlsx
python make_invoice.py       # -> Qflow_Invoice_Template.xlsx
python make_process.py       # -> Qflow_Process_Template.xlsx
python make_stamp.py         # -> Qflow_Company_Stamp.png
python make_banner.py        # -> Qflow_LinkedIn_Banner_1128x191.png
python make_logo2.py         # -> logo mark + lockup PNGs
```

Copy the regenerated file(s) into `Contracts/`, `Templates/`, or `Logo/` as appropriate.
