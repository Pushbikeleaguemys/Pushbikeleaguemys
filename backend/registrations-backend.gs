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
 *    It creates the "Sign up details" tab in your sign-up Sheet with an Approved / Pending / Rejected dropdown.
 * 4. Deploy > Manage deployments > Edit (pencil) > New version > Deploy.
 *    (Keep the same web app URL.)
 * 5. Install the approval trigger: Triggers (clock icon) > Add Trigger >
 *    function: onRegEdit, event source: From spreadsheet, event type: On edit.
 *
 * APPROVING A RIDER: open the "Sign up details" tab, check "Email Verified" is TRUE,
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

function ss_(){ return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function regSheet_(){
  var ss = ss_(), sh = ss.getSheetByName(REG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REG_SHEET);
    sh.getRange(1,1,1,REG_HEAD.length).setValues([REG_HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2,C.STATUS,1000,1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'],true).build());
  }
  return sh;
}
function setupRegistrations(){ regSheet_(); }

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
  if (!body || (body.action !== 'register' && body.action !== 'mystatus')) return null;
  var who = verifyToken_(body.idToken);
  if (!who) return {error:'bad token'};

  if (body.action === 'register') {
    var p = body.profile || {};
    if (!TEAM_CODES[p.teamCode]) return {error:'bad team'};
    if (new RegExp('@'+RIDER_DOMAIN.replace(/\./g,'\\.')+'$','i').test(who.email)) return {error:'not allowed'};
    var lock = LockService.getScriptLock(); lock.waitLock(20000);
    try {
      var sh = regSheet_(), row = findRow_(sh, who.uid);
      var vals = [new Date(), who.uid, who.email, who.verified, clean_(p.riderName), clean_(p.nickname), TEAM_CODES[p.teamCode], p.teamCode,
                  clean_(p.dob,20), clean_(p.bib,10), clean_(p.parentName), clean_(p.phone,30)];
      if (row) { sh.getRange(row, 1, 1, vals.length).setValues([vals]); }          // re-submit keeps Status / Rider ID
      else { sh.appendRow(vals.concat(['Pending','']));
             sh.getRange(sh.getLastRow(), C.STATUS).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Pending','Approved','Rejected'],true).build()); }
    } finally { lock.releaseLock(); }
    return {ok:true, status:'pending'};
  }

  // mystatus
  var m = who.email.match(/^cpb(\d{3})@/i);
  if (m && who.email.slice(-RIDER_DOMAIN.length-1) === '@'+RIDER_DOMAIN && +m[1] >= 1 && +m[1] <= 35)
    return {status:'approved', riderId:'CPB'+m[1], team:'CPB'};
  var sh2 = regSheet_(), r = findRow_(sh2, who.uid);
  if (!r) return {status:'none'};
  if (who.verified && sh2.getRange(r, C.VER).getValue() !== true) sh2.getRange(r, C.VER).setValue(true);
  var v = sh2.getRange(r, 1, 1, REG_HEAD.length).getValues()[0];
  return {status:String(v[C.STATUS-1]||'Pending').toLowerCase(), riderId:String(v[C.RID-1]||''), team:String(v[C.CODE-1]||'')};
}

// Installable "On edit" trigger: when Status becomes Approved, issue the next Rider ID for that team.
function onRegEdit(e){
  var sh = e.range.getSheet();
  if (sh.getName() !== REG_SHEET || e.range.getColumn() !== C.STATUS || e.range.getRow() < 2) return;
  if (String(e.value).toLowerCase() !== 'approved') return;
  var row = e.range.getRow();
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
