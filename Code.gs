/** Sandhya Enterprises - Google Sheets API bridge.
 * Deploy as a Web App. Access can be anonymous only because the endpoint
 * receives a strong random shared key. Keep the URL and key private.
 */
const HEADERS = {
  Tenders: ["Tender_ID", "Source", "Title", "Department", "Deadline", "Estimated_Value", "Location", "Tender_URL", "Keywords_Matched", "First_Seen_UTC", "Last_Checked_UTC", "Status", "Notification_Status", "Reference_No", "Worksite_Address", "Worksite_PIN", "Value_INR", "EMD", "EMD_Exemption", "Tender_Fee", "Submission_Start", "Bid_Opening_Date", "Completion_Period", "NIT_Document", "BOQ_Document", "Tender_Document", "Corrigendum", "MES_Eligibility", "Technical_Eligibility", "Work_Category", "Manpower_Flag", "Bid_Decision", "Verification_Status", "Description", "Work_Scope", "Manpower_Assessment", "Profitability_Status", "Material_Availability", "Completion_Risk"],
  Compliance: ["Tender_ID", "Document_Name", "Check_Category", "Finding", "Severity", "Recommended_Action", "Review_Status"],
  Logs: ["Timestamp_UTC", "Module", "Status", "Message"]
};

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const desired = HEADERS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, desired.length).setValues([desired]);
    } else {
      const currentWidth = Math.max(sh.getLastColumn(), 1);
      const existing = sh.getRange(1, 1, 1, currentWidth).getDisplayValues()[0].map(String);
      const missing = desired.filter(h => existing.indexOf(h) === -1);
      if (missing.length) {
        sh.getRange(1, currentWidth + 1, 1, missing.length).setValues([missing]);
      }
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight("bold");
  });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("AGENT_SHARED_KEY")) {
    props.setProperty("AGENT_SHARED_KEY", Utilities.getUuid() + Utilities.getUuid());
  }
  Logger.log("Sheets ready. Copy AGENT_SHARED_KEY from Project Settings > Script Properties.");
}

function doGet() {
  return json_({ok: true, service: "Sandhya Tender Sheets Bridge", hint: "POST operation=ping to test write/API path."});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("Missing request body");
    if (e.postData.contents.length > 450000) throw new Error("Request body too large");
    const p = JSON.parse(e.postData.contents);
    const expected = PropertiesService.getScriptProperties().getProperty("AGENT_SHARED_KEY");
    if (!expected || !p.key || p.key !== expected) return json_({ok: false, error: "Unauthorized"});
    const operation = String(p.operation || "");
    lock.waitLock(20000);
    try {
      if (operation === "ping") return json_({ok: true, service: "Sheets bridge", time: new Date().toISOString()});
      if (operation === "upsert_tenders") return upsertTenders_(p.records || []);
      if (operation === "get_pending") return getPending_();
      if (operation === "mark_sent") return markSent_(String(p.tender_id || ""));
      if (operation === "add_compliance") return appendRows_("Compliance", p.records || []);
      if (operation === "add_log") return addLog_(String(p.status || "INFO"), String(p.message || ""));
      throw new Error("Unsupported operation");
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ok: false, error: String(err && err.message ? err.message : err)});
  }
}

function upsertTenders_(records) {
  if (!Array.isArray(records) || records.length > 100) throw new Error("records must be an array of maximum 100 rows");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenders");
  if (!sh) throw new Error("Tenders sheet missing; run setupSheets first");
  const existing = new Map();
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach((r, i) => existing.set(String(r[0]), i + 2));
  }
  const keys = HEADERS.Tenders;
  const toAppend = [];
  let updated = 0;
  records.forEach(r => {
    const id = String(r.Tender_ID || "");
    if (!id) return;
    if (existing.has(id)) {
      // Existing sheet row: refresh Last_Checked_UTC. A value of -1 means
      // the same ID was already queued for append in this incoming batch;
      // skip that duplicate instead of issuing an invalid row -1 update.
      const rowNumber = existing.get(id);
      if (rowNumber > 0) {
        sh.getRange(rowNumber, 11).setValue(safeCell_(r.Last_Checked_UTC || new Date().toISOString()));
        updated++;
      }
      return;
    }
    const row = keys.map(k => safeCell_(r[k] === undefined ? "" : r[k]));
    toAppend.push(row);
    existing.set(id, -1);
  });
  if (toAppend.length) {
    sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, keys.length).setValues(toAppend);
  }
  return json_({ok: true, inserted: toAppend.length, updated: updated, received: records.length});
}

function getPending_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenders");
  if (!sh) throw new Error("Tenders sheet missing");
  if (sh.getLastRow() < 2) return json_({ok: true, records: []});
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.Tenders.length).getValues();
  const rows = [];
  values.forEach(r => {
    if (String(r[12]).toUpperCase() === "PENDING") {
      const obj = {};
      HEADERS.Tenders.forEach((h, i) => obj[h] = r[i] instanceof Date ? r[i].toISOString() : String(r[i]));
      rows.push(obj);
    }
  });
  return json_({ok: true, records: rows.slice(0, 50)});
}

function markSent_(id) {
  if (!id) throw new Error("tender_id required");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tenders");
  if (!sh || sh.getLastRow() < 2) throw new Error("Tender row not found");
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sh.getRange(i + 2, 13).setValue("SENT");
      return json_({ok: true, tender_id: id, status: "SENT"});
    }
  }
  throw new Error("Tender_ID not found: " + id);
}

function appendRows_(sheetName, rows) {
  if (!Array.isArray(rows) || rows.length > 100) throw new Error("records must be an array of maximum 100 rows");
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) throw new Error(sheetName + " sheet missing");
  if (rows.length) {
    const width = HEADERS[sheetName].length;
    rows.forEach(r => { if (!Array.isArray(r) || r.length !== width) throw new Error("Invalid row width for " + sheetName); });
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, width).setValues(rows.map(r => r.map(safeCell_)));
  }
  return json_({ok: true, inserted: rows.length});
}

function addLog_(status, message) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
  if (!sh) throw new Error("Logs sheet missing");
  sh.appendRow([new Date().toISOString(), "Tender Agent", safeCell_(status), safeCell_(message)]);
  return json_({ok: true});
}

function safeCell_(value) {
  const text = String(value === null || value === undefined ? "" : value);
  // Prevent imported text from being interpreted as a spreadsheet formula.
  return /^[=+@]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
