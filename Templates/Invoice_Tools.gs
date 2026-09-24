/**
 * Qflow Consulting — Invoice workbook tools.
 *
 * Supports multiple clients at once: every tab named "Invoice" or
 * "Invoice - <anything>" is a live invoice. Use "New Invoice for Client..."
 * to spin up a fresh tab per client instead of overwriting one shared sheet.
 * Every action below (assign number, generate PDF, stamp, email, mark paid)
 * acts on whichever invoice tab is ACTIVE when you run it.
 *
 * Menu (Qflow Invoice Tools):
 *   New Invoice for Client...  — duplicates a blank invoice as a new tab
 *                                 named "Invoice - <Client>" and switches to it.
 *   1. Assign Invoice Number   — pulls the next number from the hidden Config
 *                                 tab and writes it into the invoice header.
 *   2. Generate PDF            — exports the active invoice tab as a PDF,
 *                                 files it in Drive under Invoices/<Client>/,
 *                                 and logs it on the "Invoice Log" tab.
 *   3. Insert Company Stamp    — drops Qflow_Company_Stamp.png onto the
 *                                 signature block. (Real e-signature — a
 *                                 legally binding signed document — is not
 *                                 something Apps Script can do. For that, use
 *                                 DocuSign or Dropbox Sign on the generated
 *                                 PDF; this menu only adds a visual stamp.)
 *   4. Email Invoice to Client — generates a fresh PDF and emails it via
 *                                 Gmail to the Client Email cell, then marks
 *                                 status SENT.
 *   Mark as Paid               — flips status to PAID (sheet + log).
 *   Open Invoice Log           — jumps to the Invoice Log tab.
 *
 * INSTALL (one paste):
 * 1. Open the Invoice workbook in Google Sheets.
 * 2. Extensions > Apps Script, delete any starter code, paste this file, save.
 * 3. Close the Apps Script tab, reload the spreadsheet.
 * 4. Upload Qflow_Company_Stamp.png (included in this package) into the SAME
 *    Drive folder as this spreadsheet — "Insert Company Stamp" looks for it
 *    by that exact file name.
 * 5. Menu "Qflow Invoice Tools" appears next to Help.
 * 6. First run of any action will prompt you to authorize Gmail/Drive access
 *    — that's expected, approve it (it's your own script on your own sheet).
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
const STAMP_ANCHOR_ROW = 32;   // inside the "AUTHORIZED SIGNATURE & STAMP" box
const STAMP_ANCHOR_COL = 3;   // column C
const STAMP_FILE_NAME = "Qflow_Company_Stamp.png";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Qflow Invoice Tools")
    .addItem("New Invoice for Client...", "newClientInvoice")
    .addSeparator()
    .addItem("1. Assign Invoice Number", "assignInvoiceNumber")
    .addItem("2. Generate PDF", "generatePdfMenu")
    .addItem("3. Insert Company Stamp", "insertStamp")
    .addItem("4. Email Invoice to Client", "emailInvoice")
    .addSeparator()
    .addItem("Mark as Paid", "markPaid")
    .addItem("Open Invoice Log", "openLog")
    .addToUi();
}

function getInvoiceSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const name = sheet.getName();
  if (RESERVED_SHEETS.indexOf(name) !== -1 || name.toLowerCase().indexOf("invoice") !== 0) {
    SpreadsheetApp.getUi().alert('Open an invoice tab first (a tab named "Invoice" or "Invoice - <Client>"), not "' + name + '".');
    return null;
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// New Invoice for Client
// ---------------------------------------------------------------------------
function clearInvoiceFields_(sheet) {
  sheet.getRange(CELL_INVOICE_NO).setValue("Invoice No.: [•]");
  sheet.getRange(CELL_DATE).setValue("Invoice Date: [DD-MMM-YYYY]");
  sheet.getRange(CELL_DUE_DATE).setValue("Due Date: [DD-MMM-YYYY]");
  sheet.getRange(CELL_STATUS).setValue("Status: DRAFT");
  sheet.getRange(CELL_CLIENT_NAME).setValue("[Client Legal Name]");
  sheet.getRange(CELL_CLIENT_EMAIL).setValue("Client Email: [name@client.com]");
  // clear the 5 line-item rows (row 15-19: # column stays, Description/Qty/Rate clear)
  for (let r = 15; r <= 19; r++) {
    sheet.getRange(r, 3).setValue("");  // C: Description
    sheet.getRange(r, 4).setValue("");  // D: Qty/Hrs
    sheet.getRange(r, 5).setValue("");  // E: Unit Rate
  }
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
    // older workbooks (imported before Invoice_BLANK existed) — duplicate whatever
    // invoice tab is active and wipe it, rather than failing outright
    source = getInvoiceSheet_();
    if (!source) return;
  }

  const newSheet = source.copyTo(ss);
  newSheet.setName(sheetName);
  newSheet.showSheet();
  clearInvoiceFields_(newSheet);
  newSheet.getRange(CELL_CLIENT_NAME).setValue(client);
  ss.setActiveSheet(newSheet);
  ui.alert('Created "' + sheetName + '". Fill in the line items and client details, then use Qflow Invoice Tools to assign a number and generate the PDF.');
}

function pad4_(n) {
  return ("0000" + n).slice(-4);
}

// ---------------------------------------------------------------------------
// 1. Assign Invoice Number
// ---------------------------------------------------------------------------
function assignInvoiceNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  const ui = SpreadsheetApp.getUi();

  const current = String(sheet.getRange(CELL_INVOICE_NO).getValue());
  if (current.indexOf("[") === -1 && current.trim() !== "Invoice No.:" && current.trim() !== "") {
    const resp = ui.alert("Invoice Number Already Set", current + "\n\nAssign a new number anyway? (This will consume the next counter value.)", ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return;
  }

  const next = cfg.getRange("B1").getValue() || 1;
  const prefix = cfg.getRange("B2").getValue() || "INV-";
  const invoiceNo = prefix + pad4_(next);
  sheet.getRange(CELL_INVOICE_NO).setValue("Invoice No.: " + invoiceNo);
  if (!sheet.getRange(CELL_DATE).getValue() || String(sheet.getRange(CELL_DATE).getValue()).indexOf("[") !== -1) {
    sheet.getRange(CELL_DATE).setValue("Invoice Date: " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy"));
  }
  cfg.getRange("B1").setValue(next + 1);
  ui.alert("Assigned " + invoiceNo);
  return invoiceNo;
}

function extractInvoiceNo_(sheet) {
  const raw = String(sheet.getRange(CELL_INVOICE_NO).getValue());
  return raw.replace("Invoice No.:", "").trim();
}

function extractClientName_(sheet) {
  const raw = String(sheet.getRange(CELL_CLIENT_NAME).getValue()).trim();
  return raw || "Unknown Client";
}

function extractClientEmail_(sheet) {
  const raw = String(sheet.getRange(CELL_CLIENT_EMAIL).getValue());
  return raw.replace("Client Email:", "").trim();
}

// ---------------------------------------------------------------------------
// Drive folder helper — Invoices/<Client Name>/ next to this spreadsheet
// ---------------------------------------------------------------------------
function getOrCreateClientFolder_(clientName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(SHEET_CONFIG);
  const configuredRootId = String(cfg.getRange("B3").getValue() || "").trim();

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
// 2. Generate PDF
// ---------------------------------------------------------------------------
function buildInvoicePdfBlob_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getInvoiceSheet_();
  if (!sheet) throw new Error("No invoice tab active.");
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() +
    "/export?format=pdf&gid=" + sheet.getSheetId() +
    "&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false" +
    "&pagenumbers=false&horizontal_alignment=CENTER&size=A4";
  const resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
  });
  return resp.getBlob();
}

function logInvoice_(invoiceNo, client, total, status, link) {
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  const data = log.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]) === invoiceNo) {
      log.getRange(r + 1, 2, 1, 5).setValues([[Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy"), client, total, status, link]]);
      return;
    }
  }
  log.appendRow([invoiceNo, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd-MMM-yyyy"), client, total, status, link]);
}

function generatePdf_() {
  const sheet = getInvoiceSheet_();
  if (!sheet) throw new Error("No invoice tab active.");
  let invoiceNo = extractInvoiceNo_(sheet);
  if (!invoiceNo || invoiceNo.indexOf("[") !== -1) {
    invoiceNo = assignInvoiceNumber();
    if (!invoiceNo) throw new Error("Invoice number not assigned.");
  }
  const client = extractClientName_(sheet);
  const total = sheet.getRange(CELL_TOTAL).getValue();
  const blob = buildInvoicePdfBlob_().setName(invoiceNo + " - " + client + ".pdf");
  const folder = getOrCreateClientFolder_(client);

  // replace an existing PDF with the same name so re-generating doesn't pile up copies
  const existing = folder.getFilesByName(blob.getName());
  while (existing.hasNext()) existing.next().setTrashed(true);

  const file = folder.createFile(blob);
  const status = String(sheet.getRange(CELL_STATUS).getValue()).replace("Status:", "").trim() || "DRAFT";
  logInvoice_(invoiceNo, client, total, status, file.getUrl());
  return { file: file, invoiceNo: invoiceNo, client: client };
}

function generatePdfMenu() {
  const result = generatePdf_();
  const html = `
    <div style="font-family:Arial,sans-serif;padding:16px;">
      <p style="font-size:13px;">${result.invoiceNo} for ${result.client} generated.</p>
      <p><a href="${result.file.getUrl()}" target="_blank" style="display:inline-block;background:#1F3864;color:#fff;
          padding:9px 14px;border-radius:4px;text-decoration:none;font-size:13px;">Open PDF</a></p>
      <p style="font-size:11px;color:#666;">Filed under Invoices/${result.client}/ and logged on the Invoice Log tab.</p>
    </div>`;
  const output = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(output, "Invoice PDF Ready");
}

// ---------------------------------------------------------------------------
// 3. Insert Company Stamp
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
      'Upload "' + STAMP_FILE_NAME + '" (included in the Templates package) into the same Drive folder ' +
      "as this spreadsheet, then run this again.", ui.ButtonSet.OK);
    return;
  }
  const blob = found.next().getBlob();
  sheet.insertImage(blob, STAMP_ANCHOR_COL, STAMP_ANCHOR_ROW)
    .setWidth(110).setHeight(110);
  ui.alert("Stamp inserted. Drag it into place if needed, then re-run Generate PDF to bake it into the export.");
}

// ---------------------------------------------------------------------------
// 4. Email Invoice to Client
// ---------------------------------------------------------------------------
function emailInvoice() {
  const sheet = getInvoiceSheet_();
  const ui = SpreadsheetApp.getUi();
  if (!sheet) return;
  const email = extractClientEmail_(sheet);
  if (!email || email.indexOf("[") !== -1 || email.indexOf("@") === -1) {
    ui.alert('Set a real "Client Email" (cell B12) before emailing.');
    return;
  }

  const result = generatePdf_();
  const client = result.client;
  const invoiceNo = result.invoiceNo;
  const total = sheet.getRange(CELL_TOTAL).getValue();
  const dueDate = String(sheet.getRange(CELL_DUE_DATE).getValue()).replace("Due Date:", "").trim();

  const subject = "Invoice " + invoiceNo + " — Qflow Consulting";
  const body =
    "Dear " + client + ",\n\n" +
    "Please find attached invoice " + invoiceNo + " for AED " + total + ", due " + dueDate + ".\n\n" +
    "Thank you for your business.\n\nQflow Consulting";

  GmailApp.sendEmail(email, subject, body, {
    attachments: [result.file.getAs(MimeType.PDF)],
    name: "Qflow Consulting",
  });

  sheet.getRange(CELL_STATUS).setValue("Status: SENT");
  logInvoice_(invoiceNo, client, total, "SENT", result.file.getUrl());
  ui.alert("Emailed " + invoiceNo + " to " + email + ".");
}

// ---------------------------------------------------------------------------
// Mark as Paid / Open Log
// ---------------------------------------------------------------------------
function markPaid() {
  const sheet = getInvoiceSheet_();
  if (!sheet) return;
  sheet.getRange(CELL_STATUS).setValue("Status: PAID");
  const invoiceNo = extractInvoiceNo_(sheet);
  const client = extractClientName_(sheet);
  const total = sheet.getRange(CELL_TOTAL).getValue();
  if (invoiceNo && invoiceNo.indexOf("[") === -1) {
    const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
    const data = log.getDataRange().getValues();
    let found = false;
    for (let r = 1; r < data.length; r++) {
      if (String(data[r][0]) === invoiceNo) { log.getRange(r + 1, 5).setValue("PAID"); found = true; break; }
    }
    if (!found) logInvoice_(invoiceNo, client, total, "PAID", "");
  }
  SpreadsheetApp.getUi().alert("Marked " + (invoiceNo || "this invoice") + " as PAID.");
}

function openLog() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG).activate();
}
