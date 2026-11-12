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
  "Service Records": { color: "#1e8e4e" }
};
const DEFECT_HEADERS = ["Received", "Submission ID", "Date", "Vehicle", "Matrícula", "Km", "Item", "Problem",
  "Photo", "Reported by", "Status", "Action taken", "Closed on"];

// ---------------------------------------------------------------- web endpoints
function doGet() {
  return json_({ ok: true, app: "vehicle-check", time: new Date().toISOString() });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: "bad json" }); }
  if (body.key !== APP_KEY) return json_({ ok: false, error: "bad key" });
  if (!body.id || !body.sheet || !body.row || !SHEETS[body.sheet]) return json_({ ok: false, error: "bad payload" });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getSheet_(ss, body.sheet);
    // Already received? (the phone retries if the signal drops mid-send)
    if (alreadyHave_(sh, body.id)) return json_({ ok: true, duplicate: true });

    const now = new Date();
    const row = Object.assign({ "Received": now }, body.row);

    // Photos → Drive
    const photos = body.photos || [];
    const byRef = {};
    let folder = null;
    if (photos.length) {
      folder = photoFolder_(body.row["Matrícula"], body.row["Date"], body.id);
      photos.forEach(function (p) {
        const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || "image/jpeg",
          (body.row["Date"] || "") + " " + (body.row["Matrícula"] || "") + " " + p.name + ".jpg");
        const file = folder.createFile(blob);
        (byRef[p.ref] = byRef[p.ref] || []).push(file.getUrl());
      });
      row["Photos"] = '=HYPERLINK("' + folder.getUrl() + '","' + photos.length + ' photo' + (photos.length > 1 ? "s" : "") + '")';
    } else {
      row["Photos"] = "";
    }
    appendRow_(sh, row);

    // Defects → their own sheet so the workshop can track them
    if (body.defects && body.defects.length) {
      const dsh = getSheet_(ss, "Defects");
      body.defects.forEach(function (d) {
        const links = byRef[d.ref] || [];
        const r = Object.assign({ "Received": now }, d.row);
        r["Photo"] = links.length ? '=HYPERLINK("' + (links.length === 1 ? links[0] : folder.getUrl()) + '","View ' +
          (links.length === 1 ? "photo" : links.length + " photos") + '")' : "";
        appendRow_(dsh, r);
      });
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
    const headers = name === "Defects" ? DEFECT_HEADERS : ["Received", "Submission ID"];
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

function alreadyHave_(sh, id) {
  const h = headers_(sh), col = h.indexOf("Submission ID") + 1;
  if (!col || sh.getLastRow() < 2) return false;
  return !!sh.getRange(2, col, sh.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
}

// Writes an object as a row, matching by column header. New fields get a new column automatically,
// so you can add checklist items in the app without touching this script.
function appendRow_(sh, obj) {
  let h = headers_(sh);
  const missing = Object.keys(obj).filter(function (k) { return h.indexOf(k) < 0; });
  if (missing.length) {
    sh.getRange(1, h.length + 1, 1, missing.length).setValues([missing]);
    h = h.concat(missing);
    styleHeader_(sh, SHEETS[sh.getName()].color);
  }
  const vals = h.map(function (k) { return k in obj ? toCell_(k, obj[k]) : ""; });
  sh.appendRow(vals);
  const r = sh.getLastRow();
  // keep text-looking values as text (e.g. matrícula), dates as dates
  h.forEach(function (k, i) {
    const v = vals[i];
    if (v instanceof Date) sh.getRange(r, i + 1).setNumberFormat(k === "Received" || k === "Submitted on phone" ? "dd/mm/yyyy hh:mm" : "dd/mm/yyyy");
  });
  // highlight problems
  const ri = h.indexOf("Result");
  if (ri >= 0) sh.getRange(r, ri + 1).setBackground(obj["Result"] === "ALL OK" ? "#e3f5ea" : "#fdeaea").setFontWeight("bold");
  const wi = h.indexOf("Warnings");
  if (wi >= 0 && obj["Warnings"]) sh.getRange(r, wi + 1).setBackground("#fff4df");
  h.forEach(function (k, i) { if (vals[i] === "X" || (vals[i] === "Yes" && /damage|Noises/i.test(k))) sh.getRange(r, i + 1).setBackground("#fdeaea").setFontColor("#c62828").setFontWeight("bold"); });
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
