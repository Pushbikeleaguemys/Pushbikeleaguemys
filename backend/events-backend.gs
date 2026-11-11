/**
 * CorePushBike League - Race Events backend (Google Apps Script web app)
 * Stores race events (name, date, location, link, poster) in a Google Sheet
 * and poster images in a Google Drive folder. Everyone can read; only the admin key can write.
 *
 * SETUP: 1) Set ADMIN_KEY below.  2) Run the function  authorize  once (Run button) and accept the permissions.
 *        3) Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone > Deploy.
 *        4) Copy the web app URL into EV_DB in the app's HTML.
 *        After any later code change: Deploy > Manage deployments > Edit > Version: New version > Deploy.
 */
var ADMIN_KEY = "CHANGE-THIS-TO-A-SECRET";
var SHEET_NAME = "events";
var FOLDER_NAME = "CPB Event Posters";

function authorize() { sheet_(); folder_(); return "ok"; }

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("SSID"), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) { ss = SpreadsheetApp.create("CPB Race Events Data"); props.setProperty("SSID", ss.getId()); }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(["id", "ts", "json"]); }
  return sh;
}

function folder_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty("FOLDER"), f = null;
  if (id) { try { f = DriveApp.getFolderById(id); } catch (e) { f = null; } }
  if (!f) { f = DriveApp.createFolder(FOLDER_NAME); props.setProperty("FOLDER", f.getId()); }
  return f;
}

function readAll_() {
  var sh = sheet_(), last = sh.getLastRow(), res = {};
  if (last < 2) return res;
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = 0; i < rows.length; i++) {
    try { var rec = JSON.parse(rows[i][2]); rec.ts = rec.ts || rows[i][1]; res[rows[i][0]] = rec; } catch (e) {}
  }
  return res;
}

function findRow_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function clean_(rec) {
  var r = {};
  r.name = String(rec.name || "").slice(0, 200);
  r.date = String(rec.date || "").slice(0, 10);
  r.location = String(rec.location || "").slice(0, 200);
  r.ts = Number(rec.ts) || Date.now();
  var link = String(rec.link || "").trim();
  r.link = /^https?:\/\/[^\s"'<>]+$/i.test(link) ? link.slice(0, 500) : "";
  var ph = [];
  if (rec.photos && rec.photos.length) {
    for (var i = 0; i < rec.photos.length && i < 3; i++) {
      var u = String(rec.photos[i]);
      if (/^https:\/\//.test(u) && u.length < 1000) ph.push(u);
    }
  }
  if (ph.length) r.photos = ph;
  if (!r.name || !/^\d{4}-\d{2}-\d{2}$/.test(r.date)) throw new Error("invalid event");
  return r;
}

function doGet(e) {
  var kind = e && e.parameter && e.parameter.kind;
  if (kind === "events") return out_(readAll_());
  return out_({ ok: true, service: "cpb-events" });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var b = JSON.parse(e.postData.contents);
    if (!b.key || b.key !== ADMIN_KEY) return out_({ error: "forbidden" });

    if (b.action === "upload") {
      var d = String(b.data || ""), c = d.indexOf(",");
      var head = d.substring(0, c);
      var m = /^data:(image\/(jpeg|png|webp|gif));base64$/.exec(head);
      if (c < 0 || !m) return out_({ error: "bad image" });
      var ext = m[2] === "jpeg" ? "jpg" : m[2];
      var blob = Utilities.newBlob(Utilities.base64Decode(d.substring(c + 1)), m[1], "poster-" + Date.now() + "." + ext);
      var file = folder_().createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return out_({ ok: true, url: "https://lh3.googleusercontent.com/d/" + file.getId() + "=w1000" });
    }

    if (b.kind !== "events") return out_({ error: "unknown kind" });
    var sh = sheet_(), row = b.id ? findRow_(sh, b.id) : 0;

    if (b.action === "put") {
      if (!b.id || !b.rec) return out_({ error: "missing data" });
      var rec = clean_(b.rec);
      if (row) sh.getRange(row, 1, 1, 3).setValues([[b.id, rec.ts, JSON.stringify(rec)]]);
      else sh.appendRow([b.id, rec.ts, JSON.stringify(rec)]);
      return out_({ ok: true });
    }
    if (b.action === "delete") {
      if (row) sh.deleteRow(row);
      return out_({ ok: true });
    }
    return out_({ error: "unknown action" });
  } catch (err) {
    return out_({ error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
