// === CONFIG ===
// Paste a Google Drive file ID here once your checklist PDF is ready.
// Leave blank for now — the email still sends, just without the attachment.
var PDF_DRIVE_FILE_ID = "";

var LEADS_SHEET_NAME = "Leads";
var EVENTS_SHEET_NAME = "Events";

function doGet(e) {
  return jsonResponse({ ok: true, message: "Rent vs Buy lead-capture endpoint is alive. POST to this URL to submit data." });
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);

    if (body.type === "lead") {
      saveLead(body);
      sendChecklistEmail(body.email);
    } else if (body.type === "event") {
      saveEvent(body);
    } else {
      return jsonResponse({ ok: false, error: "unknown type: " + body.type });
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function saveLead(body) {
  var sheet = getOrCreateSheet(LEADS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "created_at", "email", "price", "down_payment", "holding_years",
      "monthly_rent", "mortgage_rate_pct", "browser_language",
      "utm_source", "utm_medium", "utm_campaign"
    ]);
  }
  var inputs = body.calculation_inputs || {};
  sheet.appendRow([
    body.created_at || new Date().toISOString(),
    body.email || "",
    inputs.price || "",
    inputs.down_payment || "",
    inputs.holding_years || "",
    inputs.monthly_rent || "",
    inputs.mortgage_rate_pct || "",
    body.browser_language || "",
    body.utm_source || "",
    body.utm_medium || "",
    body.utm_campaign || ""
  ]);
}

function saveEvent(body) {
  var sheet = getOrCreateSheet(EVENTS_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["timestamp", "event_name", "payload"]);
  }
  sheet.appendRow([
    body.timestamp || new Date().toISOString(),
    body.event_name || "",
    JSON.stringify(body.payload || {})
  ]);
}

function sendChecklistEmail(email) {
  if (!email) return;
  var subject = "Your free Amsterdam Home Buyer's Checklist";
  var bodyText = "Hi,\n\n" +
    "Thanks for requesting the Amsterdam Home Buyer's Checklist!\n\n" +
    (PDF_DRIVE_FILE_ID
      ? "Your checklist is attached to this email."
      : "We're putting the finishing touches on the checklist and will send it to you very soon.") +
    "\n\nBest,\nRent vs Buy Calculator";

  var options = { name: "Rent vs Buy Calculator" };
  if (PDF_DRIVE_FILE_ID) {
    var pdfFile = DriveApp.getFileById(PDF_DRIVE_FILE_ID);
    options.attachments = [pdfFile.getAs(MimeType.PDF)];
  }

  MailApp.sendEmail(email, subject, bodyText, options);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
