/**
 * Vehicle Check – central database (Google Apps Script)
 * ------------------------------------------------------
 * Receives checks from the phone app and writes them into THIS Google Sheet:
 *   • "Checklists"      – one row per vehicle check
 *   • "Defects"         – one row per problem found (with Status column for the workshop)
 *   • "Service Records" – one row per service/repair record
 * Photos are saved in Google Drive in a folder called "Vehicle Check Photos".
 *
 * Setup (once): see the setup guide. In short — paste this file into
 * Extensions ▸ Apps Script, run `setup`, then Deploy ▸ New deployment ▸ Web app
 * (Execute as: Me, Who has access: Anyone) and copy the URL into the app's config.js.
 */

// Must match APP_KEY in the app's config.js. Change both if you want a different key.
const APP_KEY = "lbh-vehicles-2026";

const PHOTO_FOLDER_NAME = "Vehicle Check Photos";
const SHEETS = {
  Checklists: { color: "#0b3d5c" },
  Defects: { color: "#c62828" },
  "Service Records": { color: "#1e8e4e" },
  "Edit Log": { color: "#5f6368" }
};
const EDITLOG_HEADERS = ["Edited on", "Submission ID", "Sheet", "Vehicle", "Matrícula", "Edit no.", "Changes", "Photos added"];
const REMOVED_BY_EDIT = "Removed by driver edit";
const DEFECT_HEADERS = ["Received", "Submission ID", "Date", "Vehicle", "Matrícula", "Km", "Item", "Problem",
  "Photo", "Reported by", "Status", "Action taken", "Closed on"];

// ---------------------------------------------------------------- web endpoints
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === "fleet") {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!controllerOk_(ss, p.ckey)) return json_({ ok: false, error: "bad controller key" });
    try { return json_(controllerData_(ss)); } catch (err) { return json_({ ok: false, error: String(err) }); }
  }
  return json_({ ok: true, app: "vehicle-check", time: new Date().toISOString() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad json" }); }
  if (body.key !== APP_KEY) return json_({ ok: false, error: "bad key" });
  if (body.kind === "work") {   // controller page: work done, parts, costs
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!controllerOk_(ss, body.ckey)) return json_({ ok: false, error: "bad controller key" });
    if (!body.id || !body.row || !body.row["Matrícula"]) return json_({ ok: false, error: "bad payload" });
    const lk = LockService.getScriptLock(); lk.waitLock(30000);
    try { return json_(saveWork_(ss, body)); } catch (err) { return json_({ ok: false, error: String(err) }); } finally { lk.releaseLock(); }
  }
  if (!body.id || !body.sheet || !body.row || !SHEETS[body.sheet]) return json_({ ok: false, error: "bad payload" });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getSheet_(ss, body.sheet);
    const existing = findRow_(sh, body.id);
    // Already received? (the phone retries if the signal drops mid-send)
    if (existing && !body.edit) return json_({ ok: true, duplicate: true });

    const now = new Date();
    let old = null;
    if (existing) {
      old = readRow_(sh, existing);
      const sent = new Date(body.row["Submitted on phone"]);
      if (old["Submitted on phone"] instanceof Date && !isNaN(sent) && old["Submitted on phone"].getTime() === sent.getTime())
        return json_({ ok: true, duplicate: true }); // this exact edit was already saved
    }
    const row = Object.assign({ "Received": old ? old["Received"] : now }, body.row);

    // Photos → Drive (an edit adds to the same folder as the original check)
    const photos = body.photos || [];
    const byRef = {};
    let folder = old ? findPhotoFolder_(body.id) : null;
    if (photos.length) {
      folder = folder || photoFolder_(body.row["Matrícula"], body.row["Date"], body.id);
      photos.forEach(function (p) {
        const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || "image/jpeg",
          (body.row["Date"] || "") + " " + (body.row["Matrícula"] || "") + " " + p.name + ".jpg");
        const file = folder.createFile(blob);
        (byRef[p.ref] = byRef[p.ref] || []).push(file.getUrl());
      });
      const n = countFiles_(folder);
      row["Photos"] = '=HYPERLINK("' + folder.getUrl() + '","' + n + ' photo' + (n > 1 ? "s" : "") + '")';
    } else if (!old) {
      row["Photos"] = "";
    } // edit without new photos: the existing Photos link is kept

    if (old) {
      row["Edits"] = (Number(old["Edits"]) || 0) + 1;
      row["Last edited"] = now;
      logEdit_(ss, body, old, row, now, photos.length);
      writeRow_(sh, existing, row);
      updateDefects_(ss, body, byRef, folder, now);
    } else {
      appendRow_(sh, row);
      // Defects → their own sheet so the workshop can track them
      if (body.defects && body.defects.length) {
        const dsh = getSheet_(ss, "Defects");
        body.defects.forEach(function (d) {
          const r = Object.assign({ "Received": now }, d.row);
          r["Photo"] = photoLink_(byRef[d.ref], folder) || "";
          appendRow_(dsh, r);
        });
      }
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------- sheet helpers
function getSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const headers = name === "Defects" ? DEFECT_HEADERS : name === "Edit Log" ? EDITLOG_HEADERS : ["Received", "Submission ID"];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    styleHeader_(sh, SHEETS[name].color);
    if (name === "Defects") defectsFormatting_(sh);
  }
  return sh;
}

function styleHeader_(sh, color) {
  sh.setFrozenRows(1);
  const r = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1));
  r.setFontWeight("bold").setFontColor("#ffffff").setBackground(color).setWrap(true).setVerticalAlignment("middle");
  sh.setRowHeight(1, 42);
}

function headers_(sh) {
  const lc = sh.getLastColumn();
  return lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(String) : [];
}

function findRow_(sh, id) {
  const h = headers_(sh), col = (h.indexOf("Submission ID") >= 0 ? h.indexOf("Submission ID") : h.indexOf("Work ID")) + 1;
  if (!col || sh.getLastRow() < 2) return 0;
  const c = sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return c ? c.getRow() : 0;
}

function readRow_(sh, r) {
  const h = headers_(sh), rg = sh.getRange(r, 1, 1, h.length), v = rg.getValues()[0], f = rg.getFormulas()[0], o = {};
  h.forEach(function (k, i) { o[k] = f[i] || v[i]; });
  return o;
}

// Adds any new column headers the row needs (so new checklist items need no script change).
function ensureHeaders_(sh, obj) {
  let h = headers_(sh);
  const missing = Object.keys(obj).filter(function (k) { return h.indexOf(k) < 0; });
  if (missing.length) {
    sh.getRange(1, h.length + 1, 1, missing.length).setValues([missing]);
    h = h.concat(missing);
    styleHeader_(sh, (SHEETS[sh.getName()] || { color: "#1e8e4e" }).color);
  }
  return h;
}

// Writes an object as a new row, matching by column header.
function appendRow_(sh, obj) {
  const h = ensureHeaders_(sh, obj);
  const vals = h.map(function (k) { return k in obj ? toCell_(k, obj[k]) : ""; });
  sh.appendRow(vals);
  formatRow_(sh, sh.getLastRow(), h, vals, obj);
}

// Overwrites an existing row; columns not in obj keep what is there (e.g. notes the office added).
function writeRow_(sh, r, obj) {
  const h = ensureHeaders_(sh, obj), cur = readRow_(sh, r);
  const vals = h.map(function (k) { return k in obj ? toCell_(k, obj[k]) : (cur[k] === undefined ? "" : cur[k]); });
  const rg = sh.getRange(r, 1, 1, h.length);
  rg.setValues([vals]);
  rg.setBackground(null).setFontColor(null).setFontWeight("normal");
  formatRow_(sh, r, h, vals, Object.assign({}, cur, obj));
}

function formatRow_(sh, r, h, vals, obj) {
  h.forEach(function (k, i) {
    if (vals[i] instanceof Date) sh.getRange(r, i + 1).setNumberFormat(/^(Received|Submitted on phone|Last edited|Edited on|Closed on)$/.test(k) ? "dd/mm/yyyy hh:mm" : "dd/mm/yyyy");
  });
  const ri = h.indexOf("Result");
  if (ri >= 0 && obj["Result"]) sh.getRange(r, ri + 1).setBackground(obj["Result"] === "ALL OK" ? "#e3f5ea" : "#fdeaea").setFontWeight("bold");
  const wi = h.indexOf("Warnings");
  if (wi >= 0 && obj["Warnings"]) sh.getRange(r, wi + 1).setBackground("#fff4df");
  const ei = h.indexOf("Edits");
  if (ei >= 0 && Number(obj["Edits"]) > 0) sh.getRange(r, ei + 1).setBackground("#e8f0fe").setFontWeight("bold");
  if (sh.getName() === "Edit Log") { sh.getRange(r, h.indexOf("Changes") + 1).setWrap(true); return; }
  h.forEach(function (k, i) { if (vals[i] === "X" || (vals[i] === "Yes" && /damage|Noises/i.test(k))) sh.getRange(r, i + 1).setBackground("#fdeaea").setFontColor("#c62828").setFontWeight("bold"); });
}

function photoLink_(links, folder) {
  if (!links || !links.length) return null;
  return '=HYPERLINK("' + (links.length === 1 ? links[0] : folder.getUrl()) + '","View ' + (links.length === 1 ? "photo" : links.length + " photos") + '")';
}

// Edit: update this check's defect rows, keeping the workshop's Status / Action taken.
function updateDefects_(ss, body, byRef, folder, now) {
  const dsh = getSheet_(ss, "Defects"), h = headers_(dsh);
  const idCol = h.indexOf("Submission ID"), itemCol = h.indexOf("Item");
  const rows = {};
  if (dsh.getLastRow() > 1) {
    dsh.getRange(2, 1, dsh.getLastRow() - 1, h.length).getValues().forEach(function (v, i) {
      if (String(v[idCol]) === body.id) rows[v[itemCol]] = { r: i + 2, status: v[h.indexOf("Status")], action: v[h.indexOf("Action taken")] };
    });
  }
  const seen = {};
  (body.defects || []).forEach(function (d) {
    const item = d.row["Item"], link = photoLink_(byRef[d.ref], folder);
    seen[item] = true;
    if (rows[item]) {
      const r = Object.assign({}, d.row);
      delete r["Status"]; delete r["Action taken"]; delete r["Closed on"]; delete r["Received"];
      if (link) r["Photo"] = link;
      if (rows[item].status === "Closed" && rows[item].action === REMOVED_BY_EDIT) { r["Status"] = "Open"; r["Action taken"] = ""; r["Closed on"] = ""; }
      writeRow_(dsh, rows[item].r, r);
    } else {
      const r = Object.assign({ "Received": now }, d.row);
      r["Photo"] = link || "";
      appendRow_(dsh, r);
    }
  });
  Object.keys(rows).forEach(function (item) {
    if (!seen[item] && rows[item].status !== "Closed")
      writeRow_(dsh, rows[item].r, { "Status": "Closed", "Action taken": REMOVED_BY_EDIT, "Closed on": now });
  });
}

function logEdit_(ss, body, old, row, now, photosAdded) {
  const tz = ss.getSpreadsheetTimeZone();
  const show = function (k, v) {
    if (v instanceof Date) return Utilities.formatDate(v, tz, /Submitted|Received|edited/i.test(k) ? "dd/MM/yyyy HH:mm" : "dd/MM/yyyy");
    return v === null || v === undefined ? "" : String(v);
  };
  const changes = [];
  Object.keys(body.row).forEach(function (k) {
    if (/^(Submitted on phone|App language|Submission ID)$/.test(k)) return;
    const a = show(k, old[k]), b = show(k, toCell_(k, body.row[k])).replace(/^'/, "");
    if (a !== b) changes.push(k + ": " + (a || "(blank)") + " → " + (b || "(blank)"));
  });
  appendRow_(getSheet_(ss, "Edit Log"), {
    "Edited on": now, "Submission ID": body.id, "Sheet": body.sheet, "Vehicle": body.row["Vehicle"], "Matrícula": body.row["Matrícula"],
    "Edit no.": row["Edits"], "Changes": changes.length ? changes.join("\n") : "(no field changes)", "Photos added": photosAdded || ""
  });
}

function findPhotoFolder_(id) {
  const it = DriveApp.searchFolders('title contains "' + id + '" and trashed = false');
  return it.hasNext() ? it.next() : null;
}
function countFiles_(folder) {
  let n = 0; const it = folder.getFiles();
  while (it.hasNext()) { it.next(); n++; }
  return n;
}

function toCell_(key, v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") {
    let m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    if (/^\d{4}-\d{2}-\d{2}T/.test(v)) { const d = new Date(v); if (!isNaN(d)) return d; }
    if (/^[=+\-@]/.test(v) && v.indexOf("=HYPERLINK(") !== 0) return "'" + v; // never let typed text become a formula
  }
  return v;
}

function defectsFormatting_(sh) {
  const statusCol = DEFECT_HEADERS.indexOf("Status") + 1;
  const range = sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1);
  range.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Open", "In progress", "Closed"], true).build());
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Open").setBackground("#fdeaea").setFontColor("#c62828").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("In progress").setBackground("#fff4df").setFontColor("#b26a00").setRanges([range]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("Closed").setBackground("#e3f5ea").setFontColor("#1e8e4e").setRanges([range]).build());
  sh.setConditionalFormatRules(rules);
}

function photoFolder_(matricula, date, id) {
  const props = PropertiesService.getScriptProperties();
  let root;
  const rid = props.getProperty("PHOTO_FOLDER_ID");
  if (rid) { try { root = DriveApp.getFolderById(rid); } catch (e) { root = null; } }
  if (!root) { root = DriveApp.createFolder(PHOTO_FOLDER_NAME); props.setProperty("PHOTO_FOLDER_ID", root.getId()); }
  const vname = String(matricula || "UNKNOWN").toUpperCase();
  const it = root.getFoldersByName(vname);
  const vf = it.hasNext() ? it.next() : root.createFolder(vname);
  return vf.createFolder((date || "") + " " + id);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- menu + one-time setup
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Vehicle Check")
    .addItem("Email me an Excel copy now", "emailExcelCopy")
    .addItem("Send me an Excel copy every Monday", "scheduleWeeklyExcel")
    .addItem("Stop the weekly Excel email", "stopWeeklyExcel")
    .addSeparator()
    .addItem("Refresh Fleet & Dashboard now", "refreshFleet")
    .addItem("Send weekly reports now", "sendWeeklyReports")
    .addItem("Start weekly reports (Mondays) + daily refresh", "startWeeklyReports")
    .addItem("Stop weekly reports", "stopWeeklyReports")
    .addSeparator()
    .addItem("Run setup again", "setup")
    .addToUi();
}

/** Run this once from the Apps Script editor (select "setup" ▸ Run). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SHEETS).forEach(function (n) { getSheet_(ss, n); });
  const s1 = ss.getSheetByName("Sheet1");
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  photoFolder_root_();
  setupFleet_(ss);
  refreshFleet();
  try { SpreadsheetApp.getUi().alert("Setup complete. Now deploy as a Web app (Deploy ▸ New deployment)."); } catch (e) { Logger.log("Setup complete."); }
}
function photoFolder_root_() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PHOTO_FOLDER_ID")) props.setProperty("PHOTO_FOLDER_ID", DriveApp.createFolder(PHOTO_FOLDER_NAME).getId());
}

// ---------------------------------------------------------------- Excel copies
function emailExcelCopy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } })
    .getBlob().setName("Vehicle Checks " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") + ".xlsx");
  const to = Session.getEffectiveUser().getEmail();
  const defects = ss.getSheetByName("Defects");
  let open = 0;
  if (defects && defects.getLastRow() > 1) {
    const col = DEFECT_HEADERS.indexOf("Status") + 1;
    open = defects.getRange(2, col, defects.getLastRow() - 1, 1).getValues().filter(function (r) { return r[0] !== "Closed"; }).length;
  }
  MailApp.sendEmail({ to: to, subject: "Vehicle checks – Excel copy (" + open + " open defects)",
    body: "Attached is the latest copy of the vehicle check database.\n\nLive sheet: " + ss.getUrl(), attachments: [blob] });
  try { SpreadsheetApp.getActive().toast("Excel copy emailed to " + to); } catch (e) {}
}
function scheduleWeeklyExcel() {
  stopWeeklyExcel();
  ScriptApp.newTrigger("emailExcelCopy").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  try { SpreadsheetApp.getUi().alert("Done – you'll get an Excel copy every Monday around 07:00."); } catch (e) {}
}
function stopWeeklyExcel() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === "emailExcelCopy") ScriptApp.deleteTrigger(t); });
}

// =====================================================================================
//  FLEET MANAGEMENT
//  Tabs: Fleet (register) · Work Log · Parts Used · Supervisors · Settings · Dashboard
//  The controller records work from the PC page (controller/), drivers' checks come
//  from the phone app, and a weekly email goes to each site's supervisor.
// =====================================================================================

const FLEET_HEADERS = ["Matrícula", "Make", "Site", "Active", "Service every (km)", "Service every (months)",
  "Schedule last service km", "Schedule next service km",
  // ↓ filled in by the script – don't type in these columns
  "Last service date", "Last service km", "Next service km", "Next service date", "Km to service", "Service status",
  "Latest km", "Latest km date", "Last check date", "Inspection expires", "Imposto expires", "Insurance expires",
  "Open defects", "Notes"];
const FLEET_AUTO_FROM = "Last service date";
const WORK_HEADERS = ["Received", "Work ID", "Date", "Matrícula", "Vehicle", "Site", "Type", "Km", "Workshop",
  "Work done", "Parts", "Parts cost", "Labour cost", "Other cost", "Total cost", "Invoice no.", "Invoice",
  "Follow-up needed", "Defects closed", "Entered by", "Source", "Notes"];
const PARTS_HEADERS = ["Work ID", "Date", "Matrícula", "Site", "Part", "Part no.", "Qty", "Unit cost", "Line cost", "Workshop"];
const SUPERVISOR_HEADERS = ["Site", "Name", "Email", "Receives", "Active"];
const WORK_TYPES = ["Service", "Repair", "Service + Repair", "Tyres", "Inspection", "Bodywork", "Other"];
const INVOICE_FOLDER_NAME = "Vehicle Invoices";
const DEFAULT_SETTINGS = [
  ["Controller key", "", "Typed once into the controller page on the PC. Change it here to lock out old PCs."],
  ["Email test mode", "YES", "YES = every report goes only to the test address below, marked with who it would have gone to. Set to NO to email the real supervisors."],
  ["Test emails go to", "", "Your address for test-mode emails."],
  ["Due soon (km)", 1000, "A service counts as 'due soon' within this many km."],
  ["Due soon (days)", 30, "…or within this many days of its date limit."],
  ["Expiry warning (days)", 30, "Documents expiring within this many days are listed."],
  ["Currency", "MZN", "Shown on costs."]
];
const FLEET_SEED = [
  ["ACI 668 MC", "Mazda BT50", "Maputo", "Yes", 7500, 6, 168216, 175716, ""],
  ["ACJ 204 MC", "Mazda BT50", "Maputo", "Yes", 7500, 6, 186932, 194432, ""],
  ["ADG 339 MC", "Mazda BT50", "Maputo", "Yes", 7500, 6, 221595, 229095, ""],
  ["ACG 006 MC", "Mazda BT50", "Maputo", "Yes", 10000, 6, 195069, 205069, ""],
  ["AMA 597 MC", "Hyundai i20", "Maputo", "Yes", 10000, 6, 6000, 16000, ""],
  ["AMA 594 MC", "Hyundai i20", "Maputo", "Yes", 10000, 6, 30000, 40000, ""],
  ["AFU 441 MP", "Toyota Corolla", "Maputo", "Yes", 10000, 6, 167495, 177495, ""],
  ["AJK 012 MP", "Isuzu Traga", "Maputo", "Yes", 10000, 6, 16032, 26032, ""],
  ["ADC 278 MP", "Hyundai H1", "Maputo", "Yes", 10000, 6, 108912, 118912, ""],
  ["AMZ 692 MC", "Hyundai Staria", "Maputo", "Yes", 10000, 6, 4787, 14787, ""],
  ["AMT 614 MC", "Mitsubishi Triton", "Maputo", "Yes", 10000, 6, 23540, 33540, ""],
  ["AMU 845 MC", "Mitsubishi Triton", "Maputo", "Yes", 10000, 6, 9272, 19272, ""],
  ["AMU 846 MC", "Mitsubishi Triton", "Nacala", "Yes", 10000, 6, 20550, 30550, ""],
  ["AFI 225 MC", "Nissan Xtrail", "Maputo", "Yes", 7500, 6, 175273, 182773, ""],
  ["AMS 766 MC", "Mahindra", "Maputo", "Yes", 5000, 6, 98950, 103950, ""],
  ["AIY 639 MC", "Land Cruiser", "Maputo", "Yes", 10000, 6, 67299, 77299, ""],
  ["ALL 515 MC", "Mazda BT50", "Beira", "Yes", 10000, 6, 70129, 80129, ""],
  ["ALL 520 MC", "Mazda BT50", "Beira", "Yes", 10000, 6, 71713, 81713, ""],
  ["ABT 039 MC", "Mazda BT50", "Beira", "Yes", 7500, 6, 164484, 171984, ""],
  ["AEZ 478 MP", "Nissan NP300 SC", "Beira", "Yes", 7500, 6, 148404, 155904, ""],
  ["AEZ 477 MP", "Nissan NP300 DC", "Beira", "Yes", 7500, 6, 213870, 221370, ""],
  ["ALW 807 MC", "Mazda BT50", "Nacala", "Yes", 10000, 6, 59960, 69960, ""],
  ["ALO 479 MC", "Mazda BT50", "Nacala", "Yes", 10000, 6, 68721, 88721, ""],
  ["AHT 582 MC", "Mazda BT50", "Nacala", "Yes", 7500, 6, 140329, 147829, ""],
  ["AFJ 088 MC", "Mazda BT50", "Nacala", "Yes", 7500, 6, 120226, 127726, ""],
  ["ABQ 407 MC", "Mazda BT50 SC", "Nacala", "Yes", 7500, 6, 169143, 176643, ""],
  ["AEE 689 MC", "Mazda BT50", "Nacala", "Yes", 7500, 6, 174720, 182220, ""],
  ["ALO 450 MC", "Mazda BT50", "Pemba", "Yes", 10000, 6, 74039, 84039, ""],
  ["AHT 584 MC", "Mazda BT50", "Pemba", "Yes", 7500, 6, 108961, 116461, ""],
  ["AGP 241 MP", "Nissan NP300 SC", "Pemba", "Yes", 10000, 6, 139068, 149068, ""],
  ["AGQ 507 MP", "Nissan NP300 SC", "Pemba", "Yes", 10000, 6, 55667, 65667, ""],
  ["ADP 554 MC", "Peugeot Partner", "SB2 Maputo", "Yes", 10000, 6, 138648, 148648, ""],
  ["AHF 992 MP", "Mercedes Sprinter", "SB2 Maputo", "Yes", 5000, 6, 360322, 365322, ""],
  ["AGF 189 MC", "Nissan Sunny", "SB2 Maputo", "Yes", 7500, 6, 106234, 113734, ""],
  ["ABX 554 MC", "Mazda BT50", "Maputo", "No", "", 6, "", "", "In Maputo service-report folder but not in the service schedule – set Active to Yes if it is in use"],
  ["ACE 205 MC", "Mazda BT50", "Maputo", "No", "", 6, "", "", "In Maputo service-report folder but not in the service schedule – set Active to Yes if it is in use"],
  ["ACN 869 MC", "Mazda BT50", "Maputo", "No", "", 6, "", "", "In Maputo service-report folder but not in the service schedule – set Active to Yes if it is in use"],
  ["ADQ 380 MC", "Mazda BT50", "Maputo", "No", "", 6, "", "", "In Maputo service-report folder but not in the service schedule – set Active to Yes if it is in use"]
];

// ---------------------------------------------------------------- small helpers
function normReg_(s) { return String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); }
function prettyReg_(s) {
  const k = normReg_(s), m = k.match(/^([A-Z]{3})(\d{3})([A-Z]{2})$/);
  return m ? m[1] + " " + m[2] + " " + m[3] : String(s || "").trim().toUpperCase();
}
function num_(v) {
  if (typeof v === "number") return v;
  const s = String(v || "").replace(/[^\d.\-]/g, "");
  return s ? Number(s) : null;
}
function asDate_(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  const s = String(v || "");
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  return null;
}
function dayStart_(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function addDays_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function addMonths_(d, n) { const x = new Date(d); x.setMonth(x.getMonth() + n); return x; }
function daysBetween_(a, b) { return Math.round((dayStart_(b) - dayStart_(a)) / 86400000); }
function fmtDate_(d) { return d ? Utilities.formatDate(d, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), "dd MMM yyyy") : ""; }
function fmtNum_(n) { return n === null || n === undefined || n === "" ? "" : Number(n).toLocaleString("en-GB", { maximumFractionDigits: 0 }); }
function esc_(s) { return String(s === null || s === undefined ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

/** Reads a tab as objects keyed by header (plus _r = sheet row). Formulas are returned for HYPERLINK cells. */
function table_(ss, name) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return { sh: sh, h: sh ? headers_(sh) : [], rows: [] };
  const h = headers_(sh), rg = sh.getRange(2, 1, sh.getLastRow() - 1, h.length), v = rg.getValues();
  return { sh: sh, h: h, rows: v.map(function (r, i) { const o = { _r: i + 2 }; h.forEach(function (k, j) { o[k] = r[j]; }); return o; }) };
}

function settings_(ss) {
  const sh = ss.getSheetByName("Settings"), o = {};
  if (sh && sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { o[String(r[0]).trim()] = r[1]; });
  DEFAULT_SETTINGS.forEach(function (d) { if (o[d[0]] === undefined || o[d[0]] === "") o[d[0]] = d[1]; });
  return o;
}

function newSheet_(ss, name, headers, color) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setFontColor("#ffffff").setBackground(color).setWrap(true).setVerticalAlignment("middle");
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 42);
  }
  return sh;
}

// ---------------------------------------------------------------- setup of the fleet tabs
function setupFleet_(ss) {
  // Settings
  let st = ss.getSheetByName("Settings");
  if (!st) {
    st = newSheet_(ss, "Settings", ["Setting", "Value", "What it does"], "#5f6368");
    const rows = DEFAULT_SETTINGS.map(function (r) { return r.slice(); });
    rows[0][1] = Utilities.getUuid().replace(/-/g, "").slice(0, 10);
    rows[2][1] = Session.getEffectiveUser().getEmail();
    st.getRange(2, 1, rows.length, 3).setValues(rows);
    st.setColumnWidth(1, 200); st.setColumnWidth(2, 220); st.setColumnWidth(3, 520);
    st.getRange(2, 3, rows.length, 1).setWrap(true).setFontColor("#5d6b7a");
  }
  // Supervisors (dummy entries until the real ones are known)
  let sup = ss.getSheetByName("Supervisors");
  if (!sup) {
    sup = newSheet_(ss, "Supervisors", SUPERVISOR_HEADERS, "#5f6368");
    const sites = ["Maputo", "SB2 Maputo", "Beira", "Nacala", "Pemba"];
    const rows = sites.map(function (s) {
      const slug = s.toLowerCase().replace(/[^a-z0-9]+/g, ".");
      return [s, "Supervisor " + s + " (dummy)", slug + ".supervisor@example.com", "Own site", "Yes"];
    });
    rows.push(["All sites", "Fleet manager", Session.getEffectiveUser().getEmail(), "All sites", "Yes"]);
    sup.getRange(2, 1, rows.length, 5).setValues(rows);
    sup.getRange(2, 4, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Own site", "All sites"], true).build());
    sup.getRange(2, 5, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build());
    [120, 240, 280, 110, 70].forEach(function (w, i) { sup.setColumnWidth(i + 1, w); });
  }
  // Fleet register
  let fl = ss.getSheetByName("Fleet");
  if (!fl) {
    fl = newSheet_(ss, "Fleet", FLEET_HEADERS, "#0b3d5c");
    const auto = FLEET_HEADERS.indexOf(FLEET_AUTO_FROM) + 1;
    fl.getRange(1, auto, 1, FLEET_HEADERS.indexOf("Open defects") - auto + 2).setBackground("#5f6368");
    const rows = FLEET_SEED.map(function (r) {
      const o = {}; ["Matrícula", "Make", "Site", "Active", "Service every (km)", "Service every (months)", "Schedule last service km", "Schedule next service km", "Notes"]
        .forEach(function (k, i) { o[k] = r[i]; });
      return FLEET_HEADERS.map(function (k) { return k in o ? o[k] : ""; });
    });
    fl.getRange(2, 1, rows.length, FLEET_HEADERS.length).setValues(rows);
    fl.getRange(2, 4, 300, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(["Yes", "No"], true).build());
    fl.setFrozenColumns(1);
    fl.getRange(1, 1, 1, FLEET_HEADERS.length).setNote(null);
    fl.getRange(1, auto).setNote("Grey columns are filled in by the script from checks and work records. Edit the blue columns.");
  }
  // Work Log + Parts Used
  const wl = newSheet_(ss, "Work Log", WORK_HEADERS, "#1e8e4e");
  wl.getRange(2, WORK_HEADERS.indexOf("Type") + 1, 2000, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(WORK_TYPES, false).build());
  newSheet_(ss, "Parts Used", PARTS_HEADERS, "#1e8e4e");
  // Dashboard
  if (!ss.getSheetByName("Dashboard")) ss.insertSheet("Dashboard", 0);
}

// ---------------------------------------------------------------- the fleet picture
/** Works out everything about every vehicle from the register, checks, service records, work log and defects. */
function fleetState_(ss) {
  const set = settings_(ss), today = dayStart_(new Date());
  const fleet = table_(ss, "Fleet").rows.filter(function (v) { return normReg_(v["Matrícula"]); });
  const byReg = {};
  fleet.forEach(function (v) {
    byReg[normReg_(v["Matrícula"])] = {
      row: v, reg: prettyReg_(v["Matrícula"]), make: v["Make"], site: String(v["Site"] || "").trim(),
      active: String(v["Active"] || "Yes").toLowerCase() !== "no",
      interval: num_(v["Service every (km)"]), months: num_(v["Service every (months)"]),
      baseKm: num_(v["Schedule last service km"]), baseNext: num_(v["Schedule next service km"]),
      services: [], kms: [], checks: [], work: [], defects: [], exp: {}
    };
  });
  const veh = function (reg) { return byReg[normReg_(reg)]; };
  table_(ss, "Checklists").rows.forEach(function (c) {
    const v = veh(c["Matrícula"]); if (!v) return;
    const d = asDate_(c["Date"]) || asDate_(c["Received"]), km = num_(c["Km"]);
    v.checks.push({ date: d, km: km, type: c["Check type"], result: c["Result"] });
    if (km) v.kms.push({ km: km, date: d });
    ["Inspection expires", "Imposto expires", "Insurance expires"].forEach(function (k) {
      const e = asDate_(c[k]);
      if (e && d && (!v.exp[k] || d >= v.exp[k].from)) v.exp[k] = { date: e, from: d };
    });
  });
  table_(ss, "Service Records").rows.forEach(function (s) {
    const v = veh(s["Matrícula"]); if (!v) return;
    const d = asDate_(s["Date"]), km = num_(s["Km"]);
    v.services.push({ date: d, km: km });
    if (km) v.kms.push({ km: km, date: d });
  });
  table_(ss, "Work Log").rows.forEach(function (w) {
    const v = veh(w["Matrícula"]); if (!v) return;
    const d = asDate_(w["Date"]), km = num_(w["Km"]);
    v.work.push({ date: d, km: km, type: w["Type"], total: num_(w["Total cost"]) || 0, row: w });
    if (/service/i.test(w["Type"])) v.services.push({ date: d, km: km });
    if (km) v.kms.push({ km: km, date: d });
  });
  table_(ss, "Defects").rows.forEach(function (d) {
    const v = veh(d["Matrícula"]); if (!v) return;
    if (String(d["Status"]) !== "Closed") v.defects.push(d);
  });

  Object.keys(byReg).forEach(function (k) {
    const v = byReg[k];
    // latest km
    v.kms.sort(function (a, b) { return a.km - b.km; });
    const top = v.kms[v.kms.length - 1];
    v.latestKm = top ? top.km : v.baseKm;
    v.latestKmDate = top ? top.date : null;
    if (v.baseKm && (!v.latestKm || v.baseKm > v.latestKm)) { v.latestKm = v.baseKm; v.latestKmDate = null; }
    // last service: newest recorded service at/after the schedule's last service, else the schedule
    const svc = v.services.filter(function (s) { return s.km || s.date; })
      .sort(function (a, b) { return (a.km || 0) - (b.km || 0) || (a.date || 0) - (b.date || 0); });
    const last = svc.length ? svc[svc.length - 1] : null;
    if (last && last.km && (!v.baseKm || last.km >= v.baseKm)) {
      v.lastSvcKm = last.km; v.lastSvcDate = last.date;
      v.nextKm = v.interval ? last.km + v.interval : null;
    } else {
      v.lastSvcKm = v.baseKm;
      const same = svc.filter(function (s) { return s.km && v.baseKm && Math.abs(s.km - v.baseKm) < 300; }).pop();
      v.lastSvcDate = same ? same.date : null;
      v.nextKm = v.baseNext || (v.baseKm && v.interval ? v.baseKm + v.interval : null);
    }
    v.nextDate = v.lastSvcDate && v.months ? addMonths_(v.lastSvcDate, v.months) : null;
    v.kmTo = v.nextKm && v.latestKm ? v.nextKm - v.latestKm : null;
    v.daysTo = v.nextDate ? daysBetween_(today, v.nextDate) : null;
    if (v.kmTo === null && v.daysTo === null) v.status = "No data";
    else if ((v.kmTo !== null && v.kmTo < 0) || (v.daysTo !== null && v.daysTo < 0)) v.status = "OVERDUE";
    else if ((v.kmTo !== null && v.kmTo <= Number(set["Due soon (km)"])) || (v.daysTo !== null && v.daysTo <= Number(set["Due soon (days)"]))) v.status = "Due soon";
    else v.status = "OK";
    // why: which limit made it overdue / due soon (km, date or both), shown as e.g. "OVERDUE – date"
    const soonKm = Number(set["Due soon (km)"]), soonDays = Number(set["Due soon (days)"]);
    const byKm = v.kmTo !== null && (v.status === "OVERDUE" ? v.kmTo < 0 : v.kmTo <= soonKm);
    const byDate = v.daysTo !== null && (v.status === "OVERDUE" ? v.daysTo < 0 : v.daysTo <= soonDays);
    v.why = (v.status === "OVERDUE" || v.status === "Due soon") ? (byKm && byDate ? "km & date" : byKm ? "km" : byDate ? "date" : "") : "";
    v.label = v.status + (v.why ? " – " + v.why : "");
    v.checks.sort(function (a, b) { return (a.date || 0) - (b.date || 0); });
    v.lastCheck = v.checks.length ? v.checks[v.checks.length - 1].date : null;
  });
  return { set: set, today: today, list: fleet.map(function (f) { return byReg[normReg_(f["Matrícula"])]; }), byReg: byReg };
}

/** Writes the grey (automatic) columns of the Fleet tab. */
function writeFleet_(ss, state, valuesOnly) {
  const sh = ss.getSheetByName("Fleet"); if (!sh || !state.list.length) return;
  const h = headers_(sh), from = h.indexOf(FLEET_AUTO_FROM), to = h.indexOf("Open defects");
  const vals = state.list.map(function (v) {
    const o = {
      "Last service date": v.lastSvcDate || "", "Last service km": v.lastSvcKm || "", "Next service km": v.nextKm || "",
      "Next service date": v.nextDate || "", "Km to service": v.kmTo === null ? "" : v.kmTo, "Service status": v.label,
      "Latest km": v.latestKm || "", "Latest km date": v.latestKmDate || "", "Last check date": v.lastCheck || "",
      "Inspection expires": v.exp["Inspection expires"] ? v.exp["Inspection expires"].date : "",
      "Imposto expires": v.exp["Imposto expires"] ? v.exp["Imposto expires"].date : "",
      "Insurance expires": v.exp["Insurance expires"] ? v.exp["Insurance expires"].date : "",
      "Open defects": v.defects.length || ""
    };
    return h.slice(from, to + 1).map(function (k) { return k in o ? o[k] : ""; });
  });
  const rows = state.list.map(function (v) { return v.row._r; });
  // rows are contiguous from 2 in normal use; write row by row only if not
  const contiguous = rows.every(function (r, i) { return r === rows[0] + i; });
  if (contiguous) sh.getRange(rows[0], from + 1, vals.length, vals[0].length).setValues(vals);
  else vals.forEach(function (r, i) { sh.getRange(rows[i], from + 1, 1, r.length).setValues([r]); });
  if (valuesOnly) return;
  ["Last service date", "Next service date", "Latest km date", "Last check date", "Inspection expires", "Imposto expires", "Insurance expires"]
    .forEach(function (k) { sh.getRange(2, h.indexOf(k) + 1, sh.getLastRow() - 1, 1).setNumberFormat("dd/mm/yyyy"); });
  ["Last service km", "Next service km", "Km to service", "Latest km", "Schedule last service km", "Schedule next service km", "Service every (km)"]
    .forEach(function (k) { sh.getRange(2, h.indexOf(k) + 1, sh.getLastRow() - 1, 1).setNumberFormat("#,##0"); });
  const stc = h.indexOf("Service status") + 1, rng = sh.getRange(2, stc, Math.max(sh.getLastRow() - 1, 1), 1);
  const rules = sh.getConditionalFormatRules().filter(function (r) {
    return !r.getRanges().some(function (x) { return x.getColumn() === stc; });
  });
  [["OVERDUE", "#fdeaea", "#c62828"], ["Due soon", "#fff4df", "#b26a00"], ["OK", "#e3f5ea", "#1e8e4e"]].forEach(function (c) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith(c[0]).setBackground(c[1]).setFontColor(c[2]).setBold(true).setRanges([rng]).build());
  });
  sh.setConditionalFormatRules(rules);
}

// ---------------------------------------------------------------- reports
/** Previous Monday–Sunday (the week the weekly checks are judged on). */
function lastWeek_(today) {
  const dow = (today.getDay() + 6) % 7;           // Monday = 0
  const thisMon = addDays_(dayStart_(today), -dow);
  return { from: addDays_(thisMon, -7), to: addDays_(thisMon, -1) };
}

/** Everything a report needs, for one site (or all sites when site is null). */
function buildReport_(state, site) {
  const set = state.set, today = state.today, wk = lastWeek_(today);
  const inSite = function (v) { return !site || v.site.toLowerCase() === String(site).toLowerCase(); };
  const vs = state.list.filter(function (v) { return v.active && inSite(v); });
  const r = { site: site || "All sites", week: wk, vehicles: vs.length, currency: set["Currency"] };
  const order = { "OVERDUE": 0, "Due soon": 1 };
  r.services = vs.filter(function (v) { return v.status === "OVERDUE" || v.status === "Due soon"; })
    .sort(function (a, b) { return order[a.status] - order[b.status] || (a.kmTo === null ? 1e9 : a.kmTo) - (b.kmTo === null ? 1e9 : b.kmTo); });
  r.missing = vs.filter(function (v) {
    return !v.checks.some(function (c) { return c.date && c.date >= wk.from && c.date < addDays_(wk.to, 1); });
  });
  r.defects = [];
  vs.forEach(function (v) {
    v.defects.forEach(function (d) {
      const dd = asDate_(d["Date"]) || asDate_(d["Received"]);
      r.defects.push({ v: v, date: dd, item: d["Item"], problem: d["Problem"], status: d["Status"] || "Open", by: d["Reported by"], age: dd ? daysBetween_(dd, today) : null });
    });
  });
  r.defects.sort(function (a, b) { return (b.age || 0) - (a.age || 0); });
  r.expiry = [];
  const warn = Number(set["Expiry warning (days)"]);
  vs.forEach(function (v) {
    ["Inspection expires", "Imposto expires", "Insurance expires"].forEach(function (k) {
      if (!v.exp[k]) return;
      const left = daysBetween_(today, v.exp[k].date);
      if (left <= warn) r.expiry.push({ v: v, doc: k.replace(" expires", ""), date: v.exp[k].date, left: left });
    });
  });
  r.expiry.sort(function (a, b) { return a.left - b.left; });
  // costs
  const mStart = new Date(today.getFullYear(), today.getMonth(), 1), pmStart = new Date(today.getFullYear(), today.getMonth() - 1, 1),
    yStart = new Date(today.getFullYear(), 0, 1);
  r.costs = vs.map(function (v) {
    const c = { v: v, month: 0, lastMonth: 0, ytd: 0, jobs: 0 };
    v.work.forEach(function (w) {
      if (!w.date) return;
      if (w.date >= mStart) c.month += w.total;
      else if (w.date >= pmStart) c.lastMonth += w.total;
      if (w.date >= yStart) { c.ytd += w.total; c.jobs++; }
    });
    return c;
  }).filter(function (c) { return c.ytd || c.lastMonth; }).sort(function (a, b) { return b.ytd - a.ytd; });
  r.totals = r.costs.reduce(function (t, c) { t.month += c.month; t.lastMonth += c.lastMonth; t.ytd += c.ytd; return t; }, { month: 0, lastMonth: 0, ytd: 0 });
  r.work = [];
  vs.forEach(function (v) { v.work.forEach(function (w) { if (w.date && w.date >= wk.from && w.date < addDays_(wk.to, 1)) r.work.push({ v: v, w: w }); }); });
  r.work.sort(function (a, b) { return a.w.date - b.w.date; });
  r.overdue = r.services.filter(function (v) { return v.status === "OVERDUE"; }).length;
  return r;
}

/** Rebuilds the Dashboard tab and the automatic Fleet columns. Run from the menu or the daily trigger. */
function refreshFleet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = fleetState_(ss);
  writeFleet_(ss, state);
  writeDashboard_(ss, buildReport_(state, null));
  try { ss.toast("Fleet and Dashboard updated"); } catch (e) {}
  return state;
}

function writeDashboard_(ss, r) {
  let sh = ss.getSheetByName("Dashboard") || ss.insertSheet("Dashboard", 0);
  sh.clear(); sh.clearConditionalFormatRules();
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontFamily("Arial").setVerticalAlignment("top");
  let row = 1;
  const title = function (t, color, note) {
    row++;
    sh.getRange(row, 1, 1, 8).merge().setValue(t).setFontWeight("bold").setFontSize(12).setFontColor("#ffffff").setBackground(color);
    row++;
    if (note) { sh.getRange(row, 1, 1, 8).merge().setValue(note).setFontColor("#5d6b7a").setFontStyle("italic"); row++; }
  };
  const table = function (head, rows, empty) {
    sh.getRange(row, 1, 1, head.length).setValues([head]).setFontWeight("bold").setBackground("#e6eef4");
    row++;
    if (!rows.length) { sh.getRange(row, 1).setValue(empty).setFontColor("#1e8e4e"); row++; return; }
    sh.getRange(row, 1, rows.length, head.length).setValues(rows).setWrap(true);
    const start = row; row += rows.length; return start;
  };
  sh.getRange(1, 1, 1, 8).merge().setValue("FLEET DASHBOARD – updated " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd MMM yyyy HH:mm"))
    .setFontSize(15).setFontWeight("bold").setFontColor("#0b3d5c");
  row = 2;
  sh.getRange(row, 1, 1, 8).merge().setValue(
    r.vehicles + " active vehicles · " + r.overdue + " services overdue · " + (r.services.length - r.overdue) + " due soon · " +
    r.missing.length + " weekly checks missing (" + fmtDate_(r.week.from) + "–" + fmtDate_(r.week.to) + ") · " +
    r.defects.length + " open defects · spend this month " + fmtNum_(r.totals.month) + " " + r.currency).setFontWeight("bold");

  title("SERVICES OVERDUE / DUE SOON", "#c62828");
  let s = table_rows_services_(r);
  const sStart = table(["Site", "Matrícula", "Make", "Status", "Latest km", "Next service km", "Km to go", "Next service date"], s, "No services due ✓");
  if (sStart) s.forEach(function (x, i) { sh.getRange(sStart + i, 4).setFontColor(String(x[3]).indexOf("OVERDUE") === 0 ? "#c62828" : "#b26a00").setFontWeight("bold"); });

  title("WEEKLY CHECKS NOT DONE – " + fmtDate_(r.week.from) + " to " + fmtDate_(r.week.to), "#b26a00");
  table(["Site", "Matrícula", "Make", "Last check", "Days since"], r.missing.map(function (v) {
    return [v.site, v.reg, v.make, v.lastCheck || "never", v.lastCheck ? daysBetween_(v.lastCheck, new Date()) : ""];
  }), "Every active vehicle was checked ✓");

  title("WORK NEEDED – open defects and follow-ups", "#0b3d5c");
  table(["Site", "Matrícula", "Reported", "Item", "Problem", "Status", "Days open", "Reported by"], r.defects.map(function (d) {
    return [d.v.site, d.v.reg, d.date || "", d.item, d.problem, d.status, d.age, d.by];
  }), "No open defects ✓");

  title("DOCUMENTS EXPIRING", "#6a1b9a", "From the latest vehicle check. Blank if drivers haven't entered the dates.");
  table(["Site", "Matrícula", "Document", "Expires", "Days left"], r.expiry.map(function (e) {
    return [e.v.site, e.v.reg, e.doc, e.date, e.left < 0 ? "EXPIRED" : e.left];
  }), "Nothing expiring ✓");

  title("COSTS (" + r.currency + ")", "#1e8e4e");
  const cStart = table(["Site", "Matrícula", "Make", "This month", "Last month", "Year to date", "Jobs this year"], r.costs.map(function (c) {
    return [c.v.site, c.v.reg, c.v.make, c.month, c.lastMonth, c.ytd, c.jobs];
  }).concat(r.costs.length ? [["TOTAL", "", "", r.totals.month, r.totals.lastMonth, r.totals.ytd, ""]] : []), "No costs recorded yet");
  if (cStart) { sh.getRange(cStart, 4, r.costs.length + 1, 3).setNumberFormat("#,##0"); sh.getRange(row - 1, 1, 1, 7).setFontWeight("bold"); }

  [110, 110, 140, 150, 260, 110, 110, 140].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  // dates
  const dr = sh.getDataRange(), vals = dr.getValues();
  dr.setNumberFormats(vals.map(function (rr) { return rr.map(function (c, j) {
    return c instanceof Date ? "dd/mm/yyyy" : (typeof c === "number" && Math.abs(c) > 999 && j >= 3) ? "#,##0" : "General"; }); }));
  sh.setFrozenRows(2);
  sh.setTabColor("#0b3d5c");
}
function table_rows_services_(r) {
  return r.services.map(function (v) {
    return [v.site, v.reg, v.make, v.label, v.latestKm || "", v.nextKm || "", v.kmTo === null ? "" : v.kmTo, v.nextDate || ""];
  });
}

// ---------------------------------------------------------------- weekly email
function reportHtml_(r, ss) {
  const cur = r.currency;
  const th = function (cols) { return "<tr>" + cols.map(function (c) { return '<th style="text-align:left;background:#e6eef4;padding:6px 8px;font-size:13px">' + esc_(c) + "</th>"; }).join("") + "</tr>"; };
  const td = function (cols, style) { return "<tr>" + cols.map(function (c) { return '<td style="padding:5px 8px;border-bottom:1px solid #e3e8ed;font-size:13px;vertical-align:top;' + (style || "") + '">' + c + "</td>"; }).join("") + "</tr>"; };
  const tbl = function (head, rows, empty) {
    if (!rows.length) return '<p style="color:#1e8e4e;margin:4px 0 14px">' + esc_(empty) + "</p>";
    return '<table style="border-collapse:collapse;width:100%;margin-bottom:16px">' + th(head) + rows.join("") + "</table>";
  };
  const h2 = function (t, c) { return '<h3 style="margin:22px 0 6px;color:' + c + ';font-size:15px;border-bottom:2px solid ' + c + ';padding-bottom:3px">' + esc_(t) + "</h3>"; };
  const badge = function (t, c) { return '<b style="color:' + c + '">' + esc_(t) + "</b>"; };
  let h = '<div style="font-family:Arial,sans-serif;color:#1b2430;max-width:860px">';
  h += '<h2 style="color:#0b3d5c;margin:0 0 4px">Fleet weekly report – ' + esc_(r.site) + "</h2>";
  h += '<p style="margin:0 0 10px;color:#5d6b7a">Week ' + fmtDate_(r.week.from) + " – " + fmtDate_(r.week.to) + " · " + r.vehicles + " active vehicles</p>";
  h += '<table style="border-collapse:collapse;margin-bottom:6px"><tr>' + [
    [r.overdue, "services overdue", "#c62828"], [r.services.length - r.overdue, "due soon", "#b26a00"],
    [r.missing.length, "checks missing", "#b26a00"], [r.defects.length, "open defects", "#0b3d5c"],
    [fmtNum_(r.totals.month), cur + " this month", "#1e8e4e"]
  ].map(function (k) { return '<td style="padding:8px 14px;background:#f4f6f8;border:2px solid #fff;text-align:center"><div style="font-size:22px;font-weight:bold;color:' + k[2] + '">' + k[0] + '</div><div style="font-size:12px;color:#5d6b7a">' + k[1] + "</div></td>"; }).join("") + "</tr></table>";

  h += h2("Services overdue / due soon", "#c62828");
  h += tbl(["Site", "Matrícula", "Make", "Status", "Latest km", "Next service", "Km to go", "Date limit"], r.services.map(function (v) {
    return td([esc_(v.site), "<b>" + esc_(v.reg) + "</b>", esc_(v.make), badge(v.label, v.status === "OVERDUE" ? "#c62828" : "#b26a00"),
      fmtNum_(v.latestKm), fmtNum_(v.nextKm), v.kmTo === null ? "" : fmtNum_(v.kmTo), fmtDate_(v.nextDate)]);
  }), "No services due ✓");

  h += h2("Weekly checks not done (" + fmtDate_(r.week.from) + " – " + fmtDate_(r.week.to) + ")", "#b26a00");
  h += tbl(["Site", "Matrícula", "Make", "Last check"], r.missing.map(function (v) {
    return td([esc_(v.site), "<b>" + esc_(v.reg) + "</b>", esc_(v.make), v.lastCheck ? fmtDate_(v.lastCheck) : badge("never", "#c62828")]);
  }), "Every vehicle was checked ✓");

  h += h2("Work needed – open defects and follow-ups", "#0b3d5c");
  h += tbl(["Site", "Matrícula", "Reported", "Item", "Problem", "Status", "Days open"], r.defects.map(function (d) {
    return td([esc_(d.v.site), "<b>" + esc_(d.v.reg) + "</b>", fmtDate_(d.date), esc_(d.item), esc_(d.problem), esc_(d.status), d.age === null ? "" : d.age]);
  }), "No open defects ✓");

  h += h2("Documents expiring", "#6a1b9a");
  h += tbl(["Site", "Matrícula", "Document", "Expires", "Days left"], r.expiry.map(function (e) {
    return td([esc_(e.v.site), "<b>" + esc_(e.v.reg) + "</b>", esc_(e.doc), fmtDate_(e.date), e.left < 0 ? badge("EXPIRED", "#c62828") : e.left]);
  }), "Nothing expiring in the next weeks ✓");

  h += h2("Work done last week", "#1e8e4e");
  h += tbl(["Date", "Matrícula", "Type", "Workshop", "Work done", "Cost (" + cur + ")"], r.work.map(function (x) {
    return td([fmtDate_(x.w.date), "<b>" + esc_(x.v.reg) + "</b>", esc_(x.w.type), esc_(x.w.row["Workshop"]),
      esc_(String(x.w.row["Work done"] || "")).replace(/\n/g, "<br>"), fmtNum_(x.w.total)]);
  }), "No work recorded last week");

  h += h2("Costs (" + cur + ")", "#1e8e4e");
  h += tbl(["Matrícula", "Make", "This month", "Last month", "Year to date", "Jobs"], r.costs.map(function (c) {
    return td(["<b>" + esc_(c.v.reg) + "</b>", esc_(c.v.make), fmtNum_(c.month), fmtNum_(c.lastMonth), fmtNum_(c.ytd), c.jobs]);
  }).concat(r.costs.length ? [td(["<b>Total</b>", "", "<b>" + fmtNum_(r.totals.month) + "</b>", "<b>" + fmtNum_(r.totals.lastMonth) + "</b>", "<b>" + fmtNum_(r.totals.ytd) + "</b>", ""])] : []), "No costs recorded this year");

  h += '<p style="color:#5d6b7a;font-size:12px;margin-top:20px">Live sheet: <a href="' + ss.getUrl() + '">Vehicle Checks</a>. Automatic email from the LBH vehicle check system.</p></div>';
  return h;
}

/** Sends this week's report to every active supervisor (their own site, or all sites). */
function sendWeeklyReports() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = refreshFleet();
  const set = state.set, test = String(set["Email test mode"]).toUpperCase() !== "NO";
  const testTo = String(set["Test emails go to"] || Session.getEffectiveUser().getEmail());
  const sups = table_(ss, "Supervisors").rows.filter(function (s) { return s["Email"] && String(s["Active"]).toLowerCase() !== "no"; });
  let excel = null;
  const sent = [];
  sups.forEach(function (s) {
    const all = String(s["Receives"]).toLowerCase().indexOf("all") === 0;
    const r = buildReport_(state, all ? null : s["Site"]);
    const subject = "Fleet weekly report – " + r.site + " – w/c " + fmtDate_(r.week.from) + ": " +
      r.overdue + " overdue, " + r.missing.length + " checks missing, " + r.defects.length + " open defects";
    const opts = { htmlBody: reportHtml_(r, ss), name: "LBH Vehicle Checks" };
    if (all) {
      if (!excel) excel = excelBlob_(ss);
      opts.attachments = [excel];
    }
    const to = test ? testTo : String(s["Email"]);
    const subj = test ? "[TEST – for " + s["Name"] + " <" + s["Email"] + ">] " + subject : subject;
    MailApp.sendEmail(to, subj, "This email needs HTML to display. Open the live sheet: " + ss.getUrl(), opts);
    sent.push((test ? "TEST " : "") + s["Name"] + " → " + to);
  });
  Logger.log(sent.join("\n"));
  try { SpreadsheetApp.getUi().alert("Weekly reports sent" + (test ? " in TEST MODE (all to " + testTo + ")" : "") + ":\n\n" + sent.join("\n")); } catch (e) {}
  return sent;
}

function excelBlob_(ss) {
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  return UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() } })
    .getBlob().setName("Vehicle Checks " + Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd") + ".xlsx");
}

function startWeeklyReports() {
  stopWeeklyReports();
  ScriptApp.newTrigger("sendWeeklyReports").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  ScriptApp.newTrigger("refreshFleet").timeBased().everyDays(1).atHour(5).create();
  try { SpreadsheetApp.getUi().alert("Done. Weekly reports go out every Monday around 06:00, and the Dashboard refreshes every morning.\n\nWhile Settings ▸ Email test mode is YES, they only go to the test address."); } catch (e) {}
}
function stopWeeklyReports() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (["sendWeeklyReports", "refreshFleet"].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
}

// ---------------------------------------------------------------- controller endpoints
/** GET ?action=fleet&ckey=… – data for the controller page. */
function controllerData_(ss) {
  const state = fleetState_(ss);
  const fleet = state.list.map(function (v) {
    return { reg: v.reg, make: v.make, site: v.site, active: v.active, interval: v.interval, latestKm: v.latestKm || null,
      nextKm: v.nextKm || null, kmTo: v.kmTo, nextDate: v.nextDate ? fmtDate_(v.nextDate) : "", status: v.status, label: v.label,
      lastCheck: v.lastCheck ? fmtDate_(v.lastCheck) : "", lastService: v.lastSvcDate ? fmtDate_(v.lastSvcDate) : "",
      openDefects: v.defects.length };
  });
  const defects = [];
  state.list.forEach(function (v) {
    v.defects.forEach(function (d) {
      defects.push({ r: d._r, sid: String(d["Submission ID"]), reg: v.reg, date: fmtDate_(asDate_(d["Date"]) || asDate_(d["Received"])),
        item: d["Item"], problem: d["Problem"], status: d["Status"] || "Open", by: d["Reported by"] });
    });
  });
  const wl = table_(ss, "Work Log").rows, pu = table_(ss, "Parts Used").rows;
  const workshops = {}; wl.forEach(function (w) { if (w["Workshop"]) workshops[w["Workshop"]] = 1; });
  const parts = {}; pu.forEach(function (p) { if (p["Part"]) parts[p["Part"]] = { partNo: p["Part no."], unit: p["Unit cost"] }; });
  const recent = wl.slice(-25).reverse().map(function (w) {
    return { date: fmtDate_(asDate_(w["Date"])), reg: w["Matrícula"], type: w["Type"], workshop: w["Workshop"], total: num_(w["Total cost"]), work: String(w["Work done"] || "").split("\n")[0] };
  });
  return { ok: true, fleet: fleet, defects: defects, workshops: Object.keys(workshops).sort(), parts: parts, recent: recent,
    types: WORK_TYPES, currency: state.set["Currency"] };
}

/** POST {kind:"work"} from the controller page. */
function saveWork_(ss, body) {
  const wl = getFleetSheet_(ss, "Work Log", WORK_HEADERS);
  if (findRow_(wl, body.id)) return { ok: true, duplicate: true };
  const now = new Date(), row = body.row || {};
  const fr = table_(ss, "Fleet").rows.filter(function (f) { return normReg_(f["Matrícula"]) === normReg_(row["Matrícula"]); })[0];
  const v = fr ? { make: fr["Make"], site: String(fr["Site"] || "").trim() } : null;
  const parts = (body.parts || []).filter(function (p) { return p.part; });
  const partsCost = parts.reduce(function (t, p) { return t + (Number(p.qty) || 1) * (Number(p.unit) || 0); }, 0);
  const labour = Number(row["Labour cost"]) || 0, other = Number(row["Other cost"]) || 0;
  let total = partsCost + labour + other;
  if (!total && num_(row["Total cost"])) total = num_(row["Total cost"]);

  // invoice → Drive
  let invoice = "";
  if (body.invoice && body.invoice.data) {
    const root = folderByName_(INVOICE_FOLDER_NAME, "INVOICE_FOLDER_ID");
    const reg = prettyReg_(row["Matrícula"]);
    const it = root.getFoldersByName(reg), vf = it.hasNext() ? it.next() : root.createFolder(reg);
    const ext = (body.invoice.name || "").split(".").pop() || "pdf";
    const file = vf.createFile(Utilities.newBlob(Utilities.base64Decode(body.invoice.data), body.invoice.mime || "application/pdf",
      (row["Date"] || "") + " " + reg + (row["Invoice no."] ? " inv " + row["Invoice no."] : "") + "." + ext));
    invoice = '=HYPERLINK("' + file.getUrl() + '","View invoice")';
  }

  // close defects the controller ticked
  const closed = [];
  if (body.close && body.close.length) {
    const ds = ss.getSheetByName("Defects"), dh = headers_(ds);
    body.close.forEach(function (c) {
      const cur = readRow_(ds, Number(c.r));
      if (String(cur["Submission ID"]) !== String(c.sid) || String(cur["Item"]) !== String(c.item)) return; // row moved – skip safely
      writeRow_(ds, Number(c.r), { "Status": c.status || "Closed", "Action taken": c.action || ("Work " + body.id), "Closed on": c.status === "In progress" ? "" : now });
      closed.push(c.item + (c.status === "In progress" ? " (in progress)" : ""));
    });
  }

  appendRow_(wl, {
    "Received": now, "Work ID": body.id, "Date": row["Date"], "Matrícula": prettyReg_(row["Matrícula"]),
    "Vehicle": v ? v.make : (row["Vehicle"] || ""), "Site": v ? v.site : (row["Site"] || ""), "Type": row["Type"],
    "Km": row["Km"] ? Number(row["Km"]) : "", "Workshop": row["Workshop"], "Work done": row["Work done"],
    "Parts": parts.length ? parts.map(function (p) { return (p.qty || 1) + " × " + p.part + (p.partNo ? " (" + p.partNo + ")" : ""); }).join("\n") : (row["Parts"] || ""),
    "Parts cost": partsCost || "", "Labour cost": labour || "", "Other cost": other || "", "Total cost": total || "",
    "Invoice no.": row["Invoice no."], "Invoice": invoice, "Follow-up needed": (body.followups || []).join("\n"),
    "Defects closed": closed.join("\n"), "Entered by": row["Entered by"], "Source": row["Source"] || "Controller page", "Notes": row["Notes"] || ""
  });
  const wh = headers_(wl), lr = wl.getLastRow();
  ["Parts cost", "Labour cost", "Other cost", "Total cost", "Km"].forEach(function (k) { wl.getRange(lr, wh.indexOf(k) + 1).setNumberFormat("#,##0"); });
  wl.getRange(lr, 1, 1, wh.length).setVerticalAlignment("top");
  ["Work done", "Parts", "Follow-up needed", "Defects closed", "Notes"].forEach(function (k) { wl.getRange(lr, wh.indexOf(k) + 1).setWrap(true); });

  if (parts.length) {
    const ps = getFleetSheet_(ss, "Parts Used", PARTS_HEADERS);
    parts.forEach(function (p) {
      const q = Number(p.qty) || 1, u = Number(p.unit) || 0;
      appendRow_(ps, { "Work ID": body.id, "Date": row["Date"], "Matrícula": prettyReg_(row["Matrícula"]), "Site": v ? v.site : "",
        "Part": p.part, "Part no.": p.partNo || "", "Qty": q, "Unit cost": u || "", "Line cost": u ? q * u : "", "Workshop": row["Workshop"] });
    });
  }
  // follow-ups become open items in Defects so they stay on the list until done
  if (body.followups && body.followups.length) {
    const ds = getSheet_(ss, "Defects");
    body.followups.forEach(function (f, i) {
      appendRow_(ds, { "Received": now, "Submission ID": body.id, "Date": row["Date"], "Vehicle": v ? v.make : "", "Matrícula": prettyReg_(row["Matrícula"]),
        "Km": row["Km"] ? Number(row["Km"]) : "", "Item": "Follow-up " + (i + 1), "Problem": f, "Photo": "", "Reported by": (row["Entered by"] || "Controller") + " (controller)",
        "Status": "Open", "Action taken": "", "Closed on": "" });
    });
  }
  writeFleet_(ss, fleetState_(ss), true);
  return { ok: true, total: total };
}

function getFleetSheet_(ss, name, headers) {
  return ss.getSheetByName(name) || newSheet_(ss, name, headers, "#1e8e4e");
}
function folderByName_(name, prop) {
  const props = PropertiesService.getScriptProperties(), id = props.getProperty(prop);
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder(name); props.setProperty(prop, f.getId()); return f;
}
function controllerOk_(ss, key) {
  const k = String(settings_(ss)["Controller key"] || "");
  return k && String(key || "") === k;
}
