/**
 * Qflow Consulting — Invoice workbook tools.
 *
 * Menu (Qflow Invoice Tools):
 *   1. Assign Invoice Number   — pulls the next number from the hidden Config
 *                                 tab and writes it into the invoice header.
 *   2. Generate PDF            — exports the Invoice tab as a PDF, files it
 *                                 in Drive under Invoices/<Client Name>/, and
 *                                 logs it on the "Invoice Log" tab.
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

const SHEET_INVOICE = "Invoice";
const SHEET_LOG = "Invoice Log";
const SHEET_CONFIG = "Config";

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
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_INVOICE);
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
