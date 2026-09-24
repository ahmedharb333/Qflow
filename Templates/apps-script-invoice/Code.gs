/**
 * Qflow Consulting — Invoice workbook tools.
 *
 * Multi-client: every tab named "Invoice" or "Invoice - <anything>" is a live
 * invoice (excluding the hidden Invoice_BLANK template). All actions act on
 * whichever invoice tab is ACTIVE when you run them.
 *
 * STATE MACHINE — each invoice moves through exactly these statuses:
 *   DRAFT -> ISSUED -> SENT -> PAID
 *   - DRAFT:   being edited, no invoice number yet, nothing saved to Drive.
 *   - ISSUED:  "Finalize Invoice" has run — number assigned, v1 PDF saved,
 *              line items protected (warning-only) against silent edits.
 *   - SENT:    "Email Invoice to Client" has run.
 *   - PAID:    "Mark as Paid" has run.
 *   setStatus_() is the ONLY place that writes status — it updates the sheet
 *   cell and the Invoice Log row together, so they can't drift apart.
 *
 * Menu (Qflow Invoice Tools):
 *   New Invoice for Client...   — duplicate a blank invoice as a new tab.
 *   Preview PDF (Draft)         — render the current sheet to a PDF for a
 *                                 look, WITHOUT assigning a number, logging,
 *                                 or touching Drive's numbered invoice trail.
 *                                 Safe to run as many times as you like.
 *   Finalize Invoice            — assigns the next invoice number (once,
 *                                 lock-protected against double-assignment),
 *                                 saves the v1 PDF, and protects the client
 *                                 info / line items with a warning so you
 *                                 don't silently edit an issued invoice.
 *   Insert Company Stamp        — drops Qflow_Company_Stamp.png onto the
 *                                 signature block. This is a visual stamp
 *                                 only, not a legally binding e-signature —
 *                                 for that, run the PDF through DocuSign or
 *                                 Dropbox Sign.
 *   Regenerate PDF              — only after Finalize. Saves a NEW version
 *                                 file (v2, v3, ...) instead of overwriting
 *                                 the previous one, so every PDF you ever
 *                                 actually issued stays on record.
 *   Email Invoice to Client     — finalizes first if needed, emails the
 *                                 latest PDF version via Gmail, sets SENT.
 *   Unlock for Correction       — removes the edit warning so you can fix a
 *                                 mistake. Regenerate PDF afterwards to
 *                                 record the corrected version.
 *   Mark as Paid                — requires the invoice to be ISSUED or SENT.
 *   Open Invoice Log
 *
 * INSTALL: see README.md in this repo for the clasp-based, no-copy-paste
 * deployment; for a one-off manual paste, Extensions > Apps Script > paste
 * this file > save > reload the sheet.
 */

const SHEET_LOG = "Invoice Log";
const SHEET_CONFIG = "Config";
const SHEET_BLANK = "Invoice_BLANK";
const RESERVED_SHEETS = [SHEET_LOG, SHEET_CONFIG, SHEET_BLANK];

const CELL_INVOICE_NO = "E3";
const CELL_DATE = "E4";
const CELL_DUE_DATE = "E5";
const CELL_STATUS = "E6";
const CELL_CLIENT_NAME = "B8";
const CELL_CLIENT_EMAIL = "B12";
const CELL_TOTAL = "F23";
const STAMP_ANCHOR_ROW = 32;
const STAMP_ANCHOR_COL = 3;
const STAMP_FILE_NAME = "Qflow_Company_Stamp.png";

const PROTECTED_A1_RANGES = ["B8:C12", "E3:F6", "C15:E19"]; // client info, invoice meta, line items
const PROTECTION_TAG = "QFLOW_FINALIZED_INVOICE";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Qflow Invoice Tools")
    .addItem("New Invoice for Client...", "newClientInvoice")
    .addSeparator()
    .addItem("Preview PDF (Draft)", "previewPdf")
    .addItem("Finalize Invoice", "finalizeInvoiceMenu")
    .addItem("Insert Company Stamp", "insertStamp")
    .addItem("Regenerate PDF", "regeneratePdfMenu")
    .addItem("Email Invoice to Client", "emailInvoice")
    .addSeparator()
    .addItem("Unlock for Correction", "unlockForCorrection")
    .addItem("Mark as Paid", "markPaid")
    .addItem("Open Invoice Log", "openLog")
    .addToUi();
}

// ---------------------------------------------------------------------------
// Sheet lookup / field helpers
// ---------------------------------------------------------------------------
function getInvoiceSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = sheet.getName();
  if (RESERVED_SHEETS.indexOf(name) !== -1 || name.toLowerCase().indexOf("invoice") !== 0) {
    SpreadsheetApp.getUi().alert('Open an invoice tab first (a tab named "Invoice" or "Invoice - <Client>"), not "' + name + '".');
    return null;
  }
  return sheet;
}

function pad4_(n) {
  return ("0000" + n).slice(-4);
}

function extractInvoiceNo_(sheet) {
  return String(sheet.getRange(CELL_INVOICE_NO).getValue()).replace("Invoice No.:", "").trim();
}

function extractClientName_(sheet) {
  return String(sheet.getRange(CELL_CLIENT_NAME).getValue()).trim() || "Unknown Client";
}

function extractClientEmail_(sheet) {
  return String(sheet.getRange(CELL_CLIENT_EMAIL).getValue()).replace("Client Email:", "").trim();
}

function extractStatus_(sheet) {
  return String(sheet.getRange(CELL_STATUS).getValue()).replace("Status:", "").trim() || "DRAFT";
}

function hasRealInvoiceNumber_(sheet) {
  const no = extractInvoiceNo_(sheet);
  return !!no && no.indexOf("[") === -1;
}

// ---------------------------------------------------------------------------
// Single source of truth for status — sheet cell + Invoice Log row together
// ---------------------------------------------------------------------------
function getOrCreateLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(SHEET_LOG);
  if (log) return log;
  log = ss.insertSheet(SHEET_LOG);
  log.getRange(1, 1, 1, 6).setValues([["Invoice No.", "Date Generated", "Client", "Total (AED)", "Status", "PDF Link"]]);
  log.getRange(1, 1, 1, 6).setFontWeight("bold").setFontColor("#FFFFFF").setBackground("#1F3864");
  log.setFrozenRows(1);
  return log;
}

function setStatus_(sheet, status, pdfLink) {
  sheet.getRange(CELL_STATUS).setValue("Status: " + status);
  const invoiceNo = extractInvoiceNo_(sheet);
  if (!invoiceNo || invoiceNo.indexOf("[") !== -1) return; // not finalized yet — nothing to log
  const client = extractClientName_(sheet);
  const total = sheet.getRange(CELL_TOTAL).getValue();
  const log = getOrCreateLogSheet_();
  const data = log.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy");
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) === invoiceNo) {
      const link = pdfLink || data[r][5];
      log.getRange(r + 1, 2, 1, 5).setValues([[today, client, total, status, link]]);
      return;
    }
  }
  log.appendRow([invoiceNo, today, client, total, status, pdfLink || ""]);
}

// ---------------------------------------------------------------------------
// New Invoice for Client — strips any carried-over stamp image
// ---------------------------------------------------------------------------
function clearInvoiceFields_(sheet) {
  sheet.getRange(CELL_INVOICE_NO).setValue("Invoice No.: [•]");
  sheet.getRange(CELL_DATE).setValue("Invoice Date: [DD-MMM-YYYY]");
  sheet.getRange(CELL_DUE_DATE).setValue("Due Date: [DD-MMM-YYYY]");
  sheet.getRange(CELL_STATUS).setValue("Status: DRAFT");
  sheet.getRange(CELL_CLIENT_NAME).setValue("[Client Legal Name]");
  sheet.getRange(CELL_CLIENT_EMAIL).setValue("Client Email: [name@client.com]");
  for (let r = 15; r <= 19; r++) {
    sheet.getRange(r, 3).setValue("");
    sheet.getRange(r, 4).setValue("");
    sheet.getRange(r, 5).setValue("");
  }
  // strip any images (e.g. a stamp) carried over from the sheet this was duplicated from
  sheet.getImages().forEach(function (img) { img.remove(); });
  // remove any finalize-protection carried over from the source sheet
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === PROTECTION_TAG) p.remove();
  });
}

function newClientInvoice() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("New Invoice", "Client name (used for the tab name and the Drive folder):", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const client = resp.getResponseText().trim();
  if (!client) return;

  let sheetName = ("Invoice - " + client).substring(0, 100);
  if (ss.getSheetByName(sheetName)) {
    ui.alert('A tab named "' + sheetName + '" already exists. Open it directly, or use a different name.');
    return;
  }

  let source = ss.getSheetByName(SHEET_BLANK);
  if (!source) {
    source = getInvoiceSheet_(); // older workbooks without Invoice_BLANK — duplicate the active tab instead
    if (!source) return;
  }

  const newSheet = source.copyTo(ss);
  newSheet.setName(sheetName);
  newSheet.showSheet();
  clearInvoiceFields_(newSheet);
  newSheet.getRange(CELL_CLIENT_NAME).setValue(client);
  ss.setActiveSheet(newSheet);
  ui.alert('Created "' + sheetName + '". Fill in the line items and client details, then Finalize Invoice when ready to issue it.');
}

// ---------------------------------------------------------------------------
// Drive folder helper — Invoices/<Client Name>/ next to this spreadsheet
// ---------------------------------------------------------------------------
function getOrCreateClientFolder_(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  const configuredRootId = cfg ? String(cfg.getRange("B3").getValue() || "").trim() : "";

  let rootFolder;
  if (configuredRootId) {
    rootFolder = DriveApp.getFolderById(configuredRootId);
  } else {
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    const existing = parent.getFoldersByName("Invoices");
    rootFolder = existing.hasNext() ? existing.next() : parent.createFolder("Invoices");
  }

  const safeName = clientName.replace(/[\\/:*?"<>|]+/g, "-").trim() || "Unknown Client";
  const existingClient = rootFolder.getFoldersByName(safeName);
  return existingClient.hasNext() ? existingClient.next() : rootFolder.createFolder(safeName);
}

// ---------------------------------------------------------------------------
// PDF rendering (shared by Preview, Finalize, Regenerate, Email)
// ---------------------------------------------------------------------------
function renderSheetToPdfBlob_(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() +
    "/export?format=pdf&gid=" + sheet.getSheetId() +
    "&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false" +
    "&pagenumbers=false&horizontal_alignment=CENTER&size=A4";
  const resp = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } });
  return resp.getBlob();
}

// ---------------------------------------------------------------------------
// Preview PDF — never touches numbering, status, or the Invoice Log
// ---------------------------------------------------------------------------
function previewPdf() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  const client = extractClientName_(sheet);
  const blob = renderSheetToPdfBlob_(sheet).setName("PREVIEW - " + client + ".pdf");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  // previews are ephemeral by design — overwrite the last one, don't accumulate copies
  const olds = folder.getFilesByName(blob.getName());
  while (olds.hasNext()) olds.next().setTrashed(true);
  const outFile = folder.createFile(blob);

  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">Draft preview for ${client} — no invoice number assigned, nothing logged.</p>
      <p><a href="${outFile.getUrl()}" target="_blank" style="display:inline-block;background:#595959;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open Preview PDF</a></p>
      <p style="font-size:11px;color:#666;">Run this as many times as you like while drafting. Run "Finalize Invoice" when it's ready to issue.</p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(210), "Preview Ready");
}

// ---------------------------------------------------------------------------
// Finalize Invoice — assign number (lock-guarded), save v1 PDF, protect fields
// ---------------------------------------------------------------------------
function assignInvoiceNumberLocked_(sheet) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONFIG);
    const next = cfg.getRange("B1").getValue() || 1;
    const prefix = cfg.getRange("B2").getValue() || "INV-";
    const invoiceNo = prefix + pad4_(next);
    sheet.getRange(CELL_INVOICE_NO).setValue("Invoice No.: " + invoiceNo);
    if (!sheet.getRange(CELL_DATE).getValue() || String(sheet.getRange(CELL_DATE).getValue()).indexOf("[") !== -1) {
      sheet.getRange(CELL_DATE).setValue("Invoice Date: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy"));
    }
    cfg.getRange("B1").setValue(next + 1);
    return invoiceNo;
  } finally {
    lock.releaseLock();
  }
}

function protectFinalized_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === PROTECTION_TAG) p.remove();
  });
  PROTECTED_A1_RANGES.forEach(function (a1) {
    const p = sheet.getRange(a1).protect().setDescription(PROTECTION_TAG);
    p.setWarningOnly(true);
  });
}

function removeFinalizedProtection_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (p) {
    if (p.getDescription() === PROTECTION_TAG) p.remove();
  });
}

function nextVersionNumber_(folder, invoiceNo) {
  const files = folder.getFiles();
  let maxV = 0;
  const re = new RegExp("^" + invoiceNo.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + " v(\\d+) - ");
  while (files.hasNext()) {
    const m = re.exec(files.next().getName());
    if (m) maxV = Math.max(maxV, parseInt(m[1], 10));
  }
  return maxV + 1;
}

function saveInvoicePdfVersion_(sheet) {
  const invoiceNo = extractInvoiceNo_(sheet);
  const client = extractClientName_(sheet);
  const folder = getOrCreateClientFolder_(client);
  const version = nextVersionNumber_(folder, invoiceNo);
  const blob = renderSheetToPdfBlob_(sheet).setName(invoiceNo + " v" + version + " - " + client + ".pdf");
  const file = folder.createFile(blob); // never trashes prior versions — full issued history stays on record
  return { file: file, version: version };
}

function finalizeInvoice_(sheet) {
  if (!hasRealInvoiceNumber_(sheet)) {
    assignInvoiceNumberLocked_(sheet);
  }
  const saved = saveInvoicePdfVersion_(sheet);
  protectFinalized_(sheet);
  setStatus_(sheet, "ISSUED", saved.file.getUrl());
  return saved;
}

function finalizeInvoiceMenu() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  if (hasRealInvoiceNumber_(sheet) && extractStatus_(sheet) !== "DRAFT") {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert("Already Finalized", extractInvoiceNo_(sheet) + " is already " + extractStatus_(sheet) +
      ". Use \"Regenerate PDF\" to save an updated version instead. Finalize again anyway?", ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }
  const saved = finalizeInvoice_(sheet);
  const client = extractClientName_(sheet);
  const invoiceNo = extractInvoiceNo_(sheet);
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">${invoiceNo} for ${client} issued (v${saved.version}). Line items and client info are now protected — edits show a warning.</p>
      <p><a href="${saved.file.getUrl()}" target="_blank" style="display:inline-block;background:#1F3864;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open PDF</a></p>
      <p style="font-size:11px;color:#666;">Filed under Invoices/${client}/ and logged on the Invoice Log tab.</p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(220), "Invoice Issued");
}

// ---------------------------------------------------------------------------
// Regenerate PDF — only after Finalize; always a new version, never overwrites
// ---------------------------------------------------------------------------
function regeneratePdfMenu() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  if (!hasRealInvoiceNumber_(sheet)) {
    SpreadsheetApp.getUi().alert('This invoice hasn\'t been finalized yet. Use "Finalize Invoice" first.');
    return;
  }
  const saved = saveInvoicePdfVersion_(sheet);
  setStatus_(sheet, extractStatus_(sheet), saved.file.getUrl()); // keep current status, refresh the logged PDF link
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">Saved v${saved.version} of ${extractInvoiceNo_(sheet)}. The earlier version(s) are kept in Drive, not replaced.</p>
      <p><a href="${saved.file.getUrl()}" target="_blank" style="display:inline-block;background:#1F3864;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open PDF</a></p>
    </div>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(380).setHeight(200), "New Version Saved");
}

// ---------------------------------------------------------------------------
// Insert Company Stamp
// ---------------------------------------------------------------------------
function insertStamp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getInvoiceSheet_();
  const ui = SpreadsheetApp.getUi();
  if (!sheet) return;

  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const found = parent.getFilesByName(STAMP_FILE_NAME);
  if (!found.hasNext()) {
    ui.alert("Stamp image not found",
      'Upload "' + STAMP_FILE_NAME + '" into the same Drive folder as this spreadsheet, then run this again.', ui.ButtonSet.OK);
    return;
  }
  const blob = found.next().getBlob();
  sheet.insertImage(blob, STAMP_ANCHOR_COL, STAMP_ANCHOR_ROW).setWidth(110).setHeight(110);
  ui.alert('Stamp inserted. Run "Finalize Invoice" or "Regenerate PDF" to bake it into a saved PDF.');
}

// ---------------------------------------------------------------------------
// Email Invoice to Client — finalizes first if needed, then sends latest PDF
// ---------------------------------------------------------------------------
function emailInvoice() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const email = extractClientEmail_(sheet);
  if (!email || email.indexOf("[") !== -1 || email.indexOf("@") === -1) {
    ui.alert('Set a real "Client Email" (cell B12) before emailing.');
    return;
  }

  const saved = hasRealInvoiceNumber_(sheet) ? saveInvoicePdfVersion_(sheet) : finalizeInvoice_(sheet);
  const client = extractClientName_(sheet);
  const invoiceNo = extractInvoiceNo_(sheet);
  const total = sheet.getRange(CELL_TOTAL).getValue();
  const dueDate = String(sheet.getRange(CELL_DUE_DATE).getValue()).replace("Due Date:", "").trim();

  const subject = "Invoice " + invoiceNo + " — Qflow Consulting";
  const body =
    "Dear " + client + ",\n\n" +
    "Please find attached invoice " + invoiceNo + " for AED " + total + ", due " + dueDate + ".\n\n" +
    "Thank you for your business.\n\nQflow Consulting";

  GmailApp.sendEmail(email, subject, body, {
    attachments: [saved.file.getAs(MimeType.PDF)],
    name: "Qflow Consulting",
  });

  setStatus_(sheet, "SENT", saved.file.getUrl());
  ui.alert("Emailed " + invoiceNo + " to " + email + ".");
}

// ---------------------------------------------------------------------------
// Unlock for Correction
// ---------------------------------------------------------------------------
function unlockForCorrection() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert("Unlock for Correction",
    "This removes the edit warning on client info and line items for " + (extractInvoiceNo_(sheet) || "this invoice") +
    ". After correcting, run \"Regenerate PDF\" to save the corrected version — the original stays on file. Continue?",
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;
  removeFinalizedProtection_(sheet);
  ui.alert("Unlocked. Remember to Regenerate PDF after making your correction.");
}

// ---------------------------------------------------------------------------
// Mark as Paid / Open Log
// ---------------------------------------------------------------------------
function markPaid() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  const ui = SpreadsheetApp.getUi();
  if (!hasRealInvoiceNumber_(sheet)) {
    ui.alert('This invoice hasn\'t been finalized yet — nothing to mark as paid. Use "Finalize Invoice" first.');
    return;
  }
  const status = extractStatus_(sheet);
  if (status !== "ISSUED" && status !== "SENT") {
    ui.alert('Status is "' + status + '" — only ISSUED or SENT invoices can be marked PAID.');
    return;
  }
  setStatus_(sheet, "PAID");
  ui.alert("Marked " + extractInvoiceNo_(sheet) + " as PAID.");
}

function openLog() {
  getOrCreateLogSheet_().activate();
}
