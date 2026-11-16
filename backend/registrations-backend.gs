/**
 * REGISTRATIONS ADD-ON for your existing Apps Script (marketplace-backend.gs)
 * ---------------------------------------------------------------------
 * 1. Paste this whole file into your Apps Script project as a NEW file (e.g. registrations.gs).
 * 2. In your EXISTING doPost(e), as the first lines after you parse the body, add:
 *
 *      var rr = regHandle_(body);          // 'body' = JSON.parse(e.postData.contents)
 *      if (rr) return ContentService.createTextOutput(JSON.stringify(rr))
 *                       .setMimeType(ContentService.MimeType.JSON);
 *
 * 3. Run setupRegistrations() once (Run button) and accept the permissions.
 *    It creates the "Sign up details" tab (master list, with the Approved / Pending / Rejected dropdown)
 *    AND one tab per team (CorePushBike, KL Rocket, ...). Every sign-up is copied automatically into its
 *    team's tab, and the copy updates by itself when you approve / reject / the email gets verified.
 *    Run syncAllTeams() any time to rebuild all team tabs from the master list.
 * 4. Deploy > Manage deployments > Edit (pencil) > New version > Deploy.
 *    (Keep the same web app URL.)
 * 5. Install the approval trigger: Triggers (clock icon) > Add Trigger >
 *    function: onRegEdit, event source: From spreadsheet, event type: On edit.
 *
 * APPROVING A RIDER: open the MASTER "Sign up details" tab (team tabs are read-only copies), check "Email Verified" is TRUE,
 * then set Status to Approved. The Rider ID (team code + number, e.g. KLR001) is
 * filled in automatically. CPB IDs are only ever issued to CorePushBike.
 */
var SHEET_ID = '1pF-kWBA4IkYBq76gwUF0t1wYSB2a5n_gRE5hlCJQ3w4';   // your "sign up details" Google Sheet
var FIREBASE_API_KEY = 'AIzaSyCbnx37DZ_qn24snAHTptXkvkzh5nHNqdU';
var REG_SHEET = 'Sign up details';   // tab name created inside that Sheet
var REG_HEAD = ['Timestamp','UID','Email','Email Verified','Rider Name','Nickname','Team','Team Code','Date of Birth','Bib','Parent Name','Phone','Status','Rider ID'];
var C = {TS:1,UID:2,EMAIL:3,VER:4,NAME:5,NICK:6,TEAM:7,CODE:8,DOB:9,BIB:10,PARENT:11,PHONE:12,STATUS:13,RID:14};
var TEAM_CODES = {CPB:'CorePushBike',KLR:'KL Rocket',ADD:'ADDA Cycle',JBR:'JB Rocket',RAF:'Rafa Kids',LIR:'Lion Rocket',M8:'M8',PRS:'Pro Selangor',FWN:'Fawn Speed',DKH:'Dark Horse',TOR:'Tornado Rider',SHK:'Sharkspeed'};
var RIDER_DOMAIN = 'rider.pushbikeleague.my';
// Admin login used by the app's Admin tab (same ID / password hash as the app). Needed to list + approve riders from the app.
var REG_ADMIN_ID = 'admin';
var REG_ADMIN_HASH = 'cc0b7906c1de4708f8888a45850be22b430dcd643a3d082bb7c9ce15161512be';
// New-registration email goes here. Leave '' to email the Google account that owns this script.
var ADMIN_EMAIL = '';

function regSs_(){ return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function regSheet_(){
  var ss = regSs_(), sh = ss.getSheetByName(REG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REG_SHEET);
    sh.getRange(1,1,1,REG_HEAD.length).setValues([REG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2,C.STATUS,1000,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'],true).build());
  }
  return sh;
}
function setupRegistrations(){
  regSheet_();
  Object.keys(TEAM_CODES).forEach(function(c){ teamSheet_(c); });
  syncAllTeams();
}

// ---- One tab per team: automatic copy of the master list ----
function teamSheet_(code){
  var ss = regSs_(), name = TEAM_CODES[code], sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1,1,1,REG_HEAD.length).setValues([REG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
// Copy (insert or update) one master row into its team's tab; remove it from any other team tab (team changed).
function mirrorRow_(master, row){
  if (row < 2) return;
  var v = master.getRange(row, 1, 1, REG_HEAD.length).getValues()[0];
  var uid = v[C.UID-1], code = String(v[C.CODE-1]||'').toUpperCase();
  if (!uid || !TEAM_CODES[code]) return;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    Object.keys(TEAM_CODES).forEach(function(c){
      if (c === code) return;
      var other = regSs_().getSheetByName(TEAM_CODES[c]);
      if (other) { var r0 = findRow_(other, uid); if (r0) other.deleteRow(r0); }
    });
    var t = teamSheet_(code), r = findRow_(t, uid);
    if (r) t.getRange(r, 1, 1, v.length).setValues([v]); else t.appendRow(v);
  } finally { lock.releaseLock(); }
}
// Rebuild every team tab from the master list (run manually if needed).
function syncAllTeams(){
  var m = regSheet_(), last = m.getLastRow();
  for (var r = 2; r <= last; r++) mirrorRow_(m, r);
}

// Verify a Firebase ID token with Google; returns {uid,email,verified} or null.
function verifyToken_(idToken){
  if (!idToken) return null;
  var res = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method:'post', contentType:'application/json', muteHttpExceptions:true,
    payload: JSON.stringify({idToken: idToken})
  });
  if (res.getResponseCode() !== 200) return null;
  var u = (JSON.parse(res.getContentText()).users || [])[0];
  return u ? {uid:u.localId, email:String(u.email||'').toLowerCase(), verified:!!u.emailVerified} : null;
}
function findRow_(sh, uid){
  var last = sh.getLastRow(); if (last < 2) return 0;
  var ids = sh.getRange(2, C.UID, last-1, 1).getValues();
  for (var i=0;i<ids.length;i++) if (ids[i][0] === uid) return i+2;
  return 0;
}
function clean_(v, n){ return String(v==null?'':v).replace(/^[=+\-@]/,"'$&").slice(0, n||200); }  // block formula injection

function regHandle_(body){
  if (body && (body.action === 'adminlist' || body.action === 'adminset')) return adminHandle_(body);
  if (!body || (body.action !== 'register' && body.action !== 'mystatus')) return null;
  var who = verifyToken_(body.idToken);
  if (!who) return {error:'bad token'};

  if (body.action === 'register') {
    var p = body.profile || {};
    if (!TEAM_CODES[p.teamCode]) return {error:'bad team'};
    if (new RegExp('@'+RIDER_DOMAIN.replace(/\./g,'\\.')+'$','i').test(who.email)) return {error:'not allowed'};
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = regSheet_(), row = findRow_(sh, who.uid), mrow = 0;
      var wasVerified = row ? (sh.getRange(row, C.VER).getValue() === true) : false;
      var vals = [new Date(), who.uid, who.email, who.verified, clean_(p.riderName), clean_(p.nickname), TEAM_CODES[p.teamCode], p.teamCode,
                  clean_(p.dob,20), clean_(p.bib,10), clean_(p.parentName), clean_(p.phone,30)];
      if (row) { sh.getRange(row, 1, 1, vals.length).setValues([vals]); mrow = row; }   // re-submit keeps Status / Rider ID
      else { sh.appendRow(vals.concat(['Pending','']));
             mrow = sh.getLastRow();
             sh.getRange(mrow, C.STATUS).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'],true).build()); }
    } finally { lock.releaseLock(); }
    try { mirrorRow_(sh, mrow); } catch (err) {}      // copy into the team's tab (never blocks the sign-up)
    if (who.verified && !wasVerified) { try { notifyAdmin_(vals); } catch (err) {} }   // email the admin: new verified registration
    return {ok:true, status:'pending'};
  }

  // mystatus
  var m = who.email.match(/^cpb(\d{3})@/i);
  if (m && who.email.slice(-RIDER_DOMAIN.length-1) === '@'+RIDER_DOMAIN && +m[1] >= 1 && +m[1] <= 35)
    return {status:'approved', riderId:'CPB'+m[1], team:'CPB'};
  var sh2 = regSheet_(), r = findRow_(sh2, who.uid);
  if (!r) return {status:'none'};
  var flipped = false;
  if (who.verified && sh2.getRange(r, C.VER).getValue() !== true) { sh2.getRange(r, C.VER).setValue(true); flipped = true; try { mirrorRow_(sh2, r); } catch (err) {} }
  var v = sh2.getRange(r, 1, 1, REG_HEAD.length).getValues()[0];
  if (flipped) { try { notifyAdmin_(v); } catch (err) {} }
  return {status:String(v[C.STATUS-1]||'Pending').toLowerCase(), riderId:String(v[C.RID-1]||''), team:String(v[C.CODE-1]||'')};
}

// Installable "On edit" trigger: when Status becomes Approved, issue the next Rider ID for that team,
// then refresh the team-tab copy of every edited row.
function onRegEdit(e){
  var sh = e.range.getSheet();
  if (sh.getName() !== REG_SHEET || e.range.getRow() < 2) return;
  var r1 = e.range.getRow(), r2 = r1 + e.range.getNumRows() - 1;
  if (e.range.getColumn() <= C.STATUS && e.range.getLastColumn() >= C.STATUS && String(e.value).toLowerCase() === 'approved') assignRid_(sh, r1);
  for (var r = r1; r <= r2; r++) { try { mirrorRow_(sh, r); } catch (err) {} }
}

// Run ONCE: installs the "on edit" trigger on your sign-up Sheet (works for standalone and Sheet-bound scripts).
function installRegTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if (t.getHandlerFunction() === 'onRegEdit') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onRegEdit').forSpreadsheet(SHEET_ID).onEdit().create();
}

// Issue the next Rider ID (team code + number) for an approved row, if it has none yet.
function assignRid_(sh, row){
  if (sh.getRange(row, C.RID).getValue()) return;
  var code = String(sh.getRange(row, C.CODE).getValue()).toUpperCase();
  if (!TEAM_CODES[code]) return;
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var last = sh.getLastRow(), ids = sh.getRange(2, C.RID, Math.max(last-1,1), 1).getValues(), max = 0;
    if (code === 'CPB') max = 35;                       // CPB001-CPB035 are the pre-issued CorePushBike riders
    ids.forEach(function(x){ var m = String(x[0]).match(new RegExp('^'+code+'(\\d+)$')); if (m) max = Math.max(max, +m[1]); });
    sh.getRange(row, C.RID).setValue(code + ('000' + (max+1)).slice(-3));
  } finally { lock.releaseLock(); }
}

// ---- Admin tab in the app: list registrations + approve / reject ----
function regHash_(s){
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return d.map(function(b){ b = b < 0 ? b + 256 : b; return ('0' + b.toString(16)).slice(-2); }).join('');
}
function adminAuth_(a){
  var p = PropertiesService.getScriptProperties(), f = {};
  try { f = JSON.parse(p.getProperty('fail_regadmin') || '{}'); } catch (e) {}
  var now = Date.now();
  if (f.n >= 8 && now - f.t < 15*60000) return false;                       // locked after 8 wrong tries for 15 min
  var ok = a && a.id === REG_ADMIN_ID && regHash_('cpb|' + a.id + '|' + a.pw) === REG_ADMIN_HASH;
  if (ok) { p.deleteProperty('fail_regadmin'); return true; }
  var n = (f.n >= 8 && now - f.t >= 15*60000) ? 1 : (f.n || 0) + 1;
  p.setProperty('fail_regadmin', JSON.stringify({n:n, t:now}));
  return false;
}
function rowObj_(v){
  return {uid:String(v[C.UID-1]), email:String(v[C.EMAIL-1]), verified:v[C.VER-1] === true,
          name:String(v[C.NAME-1]), nick:String(v[C.NICK-1]), team:String(v[C.TEAM-1]), code:String(v[C.CODE-1]),
          dob:String(v[C.DOB-1]), bib:String(v[C.BIB-1]), parent:String(v[C.PARENT-1]), phone:String(v[C.PHONE-1]),
          status:String(v[C.STATUS-1] || 'Pending'), rid:String(v[C.RID-1] || ''),
          ts: v[C.TS-1] instanceof Date ? v[C.TS-1].toISOString() : String(v[C.TS-1])};
}
function adminHandle_(body){
  if (!adminAuth_(body.auth)) return {error:'Admin login required'};
  var sh = regSheet_();
  if (body.action === 'adminlist') {
    var last = sh.getLastRow(), rows = last < 2 ? [] : sh.getRange(2, 1, last-1, REG_HEAD.length).getValues();
    return {ok:true, rows: rows.filter(function(v){ return v[C.UID-1]; }).map(rowObj_).reverse()};   // newest first
  }
  // adminset: {uid, status: Approved | Rejected | Pending}
  var st = String(body.status || '');
  if (['Approved','Rejected','Pending'].indexOf(st) < 0) return {error:'bad status'};
  var row = findRow_(sh, String(body.uid || ''));
  if (!row) return {error:'registration not found'};
  sh.getRange(row, C.STATUS).setValue(st);
  if (st === 'Approved') assignRid_(sh, row);
  try { mirrorRow_(sh, row); } catch (err) {}
  return {ok:true, row: rowObj_(sh.getRange(row, 1, 1, REG_HEAD.length).getValues()[0])};
}

// ---- Email to the admin when a new rider has registered and verified their email ----
function notifyAdmin_(v){
  var to = ADMIN_EMAIL || Session.getEffectiveUser().getEmail();
  if (!to) return;
  var o = rowObj_(v);
  MailApp.sendEmail(to, 'New rider registration: ' + o.name + ' (' + o.team + ')',
    'A new rider registered and verified their email. Open the Admin tab in the app to approve or reject.\n\n' +
    'Rider: ' + o.name + (o.nick ? ' (' + o.nick + ')' : '') + '\nTeam: ' + o.team + '\nDate of birth: ' + o.dob +
    '\nBib: ' + o.bib + '\nParent: ' + o.parent + '\nPhone: ' + o.phone + '\nEmail: ' + o.email + '\n\nSheet: ' + regSs_().getUrl());
}
// Run ONCE from the editor: grants the "send email" permission and sends you a test email.
function testAdminEmail(){
  MailApp.sendEmail(ADMIN_EMAIL || Session.getEffectiveUser().getEmail(), 'CorePushBike test email', 'Registration email alerts are working.');
}
