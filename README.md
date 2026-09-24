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
  - `Invoice_Tools.gs` (canonical copy: `apps-script-invoice/Code.gs`) — adds a **Qflow Invoice Tools** menu built around an explicit state machine (`DRAFT -> ISSUED -> SENT -> PAID`): **Preview PDF** for a look without touching numbering, **Finalize Invoice** to assign the number once (lock-protected against double-assignment) and protect client info/line items, **Regenerate PDF** to save a new version without deleting prior ones, **Insert Company Stamp**, **Email Invoice to Client**, **Unlock for Correction**, and **Mark as Paid** (only from ISSUED/SENT). Every finalized invoice is filed under `Invoices/<Client Name>/` in Drive and logged on the Invoice Log tab from one single status-writing function, so the sheet and the log can't drift apart.
  - `Qflow_Company_Stamp.png` — upload into the same Drive folder as the invoice spreadsheet; "Insert Company Stamp" looks it up by that filename.
- **`Qflow_Process_Template.xlsx`** — lightweight generic process-map template (inputs/outputs, steps, revision history) for processes that don't need the full SOP structure.

**Note on e-signature:** the stamp menu item is a visual stamp only, not a legally binding signature. For real e-signing, run the generated PDF through DocuSign or Dropbox Sign.

## Using the SOP / Invoice Apps Script tools

**One-time, manual paste** (fine for a single sheet, or if you don't want the `clasp` setup below):
1. Open the `.xlsx` file in Google Sheets (import if needed — the tools only run on native Sheets, not on an un-converted `.xlsx` sitting in Drive).
2. Extensions > Apps Script, delete the starter code, paste the matching `.gs` file, save.
3. Reload the spreadsheet — the new menu appears next to Help.
4. First run of any Drive/Gmail action prompts a one-time Google authorization — approve it (it's your own script acting on your own account).

**No more copy-paste — deploy from this repo instead:** see "Deploying via clasp / GitHub" below. Once set up, editing `Templates/apps-script-invoice/Code.gs` (or `apps-script-swimlane/Code.gs`) and pushing to `main` updates the live Apps Script automatically.

## Deploying via clasp / GitHub

[`clasp`](https://github.com/google/clasp) is Google's CLI for pushing local files straight into a bound Apps Script project — that's what replaces copy-pasting into the Apps Script editor.

### One-time setup (local machine)

```bash
npm install                     # installs clasp as a dev dependency (see package.json)
npx clasp login                 # opens a browser — sign in as the Google account that owns the Sheets
```

For each project, point clasp at the **existing** bound script (don't run `clasp create` — that makes a new, disconnected script):

1. Open the spreadsheet > Extensions > Apps Script > ⚙️ Project Settings > copy the **Script ID**.
2. Create the pointer file (this is intentionally not committed — see `.gitignore`):
   ```bash
   echo '{"scriptId":"<paste the Script ID>","rootDir":"."}' > Templates/apps-script-invoice/.clasp.json
   ```
   Repeat for `Templates/apps-script-swimlane/.clasp.json` with that project's Script ID.
3. First push (this OVERWRITES whatever is currently live with this repo's code — expected, since the repo is now the source of truth):
   ```bash
   npm run push:invoice
   npm run push:swimlane
   ```

From then on: edit `Code.gs` → `npm run push:invoice` (or `push:swimlane`, or `push:all`) → the live menu updates immediately. No Apps Script editor required.

### Fully automatic — push to GitHub deploys it

`.github/workflows/deploy-apps-script.yml` runs `clasp push` automatically whenever `Templates/apps-script-invoice/**` or `Templates/apps-script-swimlane/**` changes on `main`. One more one-time setup step, in the GitHub repo itself:

1. **Secret** (Settings > Secrets and variables > Actions > *Secrets* tab > New repository secret):
   `CLASP_CREDENTIALS` = the full contents of `~/.clasprc.json` (created by `clasp login` above — this is the OAuth token, treat it like a password).
2. **Variables** (same page, *Variables* tab):
   `INVOICE_SCRIPT_ID` and `SWIMLANE_SCRIPT_ID` = the two Script IDs from step 1 above (these aren't secret, just identifiers).

After that: edit code, `git push`, and GitHub Actions pushes it live — the whole "copy the file into Apps Script" step is gone for good.

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
