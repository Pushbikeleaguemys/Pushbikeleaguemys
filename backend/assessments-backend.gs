/**
 * RIDER ASSESSMENTS ADD-ON for your existing Apps Script (the one behind PHOTO_API)
 * ---------------------------------------------------------------------
 * Stores coach/admin rider assessments (8 skills scored 0-100) in a private tab of your Google Sheet,
 * so the coach and the admin see the same data on every device. Only the coach or admin password works.
 *
 * 1. Paste this whole file into your Apps Script project as a NEW file (e.g. assessments.gs).
 * 2. In your EXISTING doPost(e), right after you parse the body (the same place you added regHandle_), add:
 *
 *      var ar = asmHandle_(body);          // 'body' = JSON.parse(e.postData.contents)
 *      if (ar) return ContentService.createTextOutput(JSON.stringify(ar))
 *                       .setMimeType(ContentService.MimeType.JSON);
 *
 * 3. Deploy > Manage deployments > Edit (pencil) > New version > Deploy (keep the same web app URL).
 * 4. In the app file change   var ASM_SERVER=false;   to   var ASM_SERVER=true;   and publish.
 *
 * The tab "Assessments" is created automatically in the same Google Sheet as your sign-ups.
 */
var ASM_SHEET_ID = '1pF-kWBA4IkYBq76gwUF0t1wYSB2a5n_gRE5hlCJQ3w4';   // same Sheet as sign-ups; or leave '' to use the script's own spreadsheet
var ASM_TAB = 'Assessments';
var DET_TAB = 'RiderDetails';
var ASM_FB_KEY = 'AIzaSyCbnx37DZ_qn24snAHTptXkvkzh5nHNqdU';   // Firebase web API key (public, same as in the app)
var ASM_RIDER_DOMAIN = 'rider.pushbikeleague.my';
var ASM_HEAD = ['Rider ID','Date','Speed','Balance','Cornering','Braking','Endurance','Agility','Technique','Race Performance','Note','By','Updated'];
var ASM_KEYS = ['speed','balance','cornering','braking','endurance','agility','technique','race'];
// sha256 of "cpb|<id>|<password>" (same scheme the app uses); passwords themselves are never stored here
var ASM_AUTH = {
  'admin': 'cc0b7906c1de4708f8888a45850be22b430dcd643a3d082bb7c9ce15161512be',
  'coach': '6cdd610fe93f8a300211543214c07d11734d23761c9d78dda8e9262d14d37616'
};

function asmSha_(s){
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function(b){ return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function asmSheet_(){
  var ss = ASM_SHEET_ID ? SpreadsheetApp.openById(ASM_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ASM_TAB);
  if (!sh) { sh = ss.insertSheet(ASM_TAB); sh.appendRow(ASM_HEAD); sh.setFrozenRows(1); sh.getRange(1,1,1,ASM_HEAD.length).setFontWeight('bold'); }
  return sh;
}
// Verifies a Firebase login token with Google and returns the rider ID (e.g. CPB012) it belongs to, or ''.
function asmRiderFromToken_(tok){
  if (!tok) return '';
  try {
    var r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + ASM_FB_KEY,
      {method:'post', contentType:'application/json', payload:JSON.stringify({idToken:String(tok)}), muteHttpExceptions:true});
    if (r.getResponseCode() !== 200) return '';
    var u = (JSON.parse(r.getContentText()).users || [])[0];
    var m = u && u.email ? /^(cpb\d{3})@(.+)$/i.exec(u.email) : null;
    return (m && m[2].toLowerCase() === ASM_RIDER_DOMAIN) ? m[1].toUpperCase() : '';
  } catch (e) { return ''; }
}
function asmHandle_(body){
  if (!body || !body.action) return null;
  var act = String(body.action);
  if (act.indexOf('assess_') !== 0 && act.indexOf('det_') !== 0) return null;
  var mine = (act === 'assess_mine' || act === 'det_mine'), who = '', id = '';
  if (mine) {
    who = asmRiderFromToken_(body.token);            // a rider can only ever read their OWN record
    if (!who) return {ok:false, error:'Please log in again'};
  } else {
    id = String(body.id || '').toLowerCase();
    if (!ASM_AUTH[id] || asmSha_('cpb|' + id + '|' + String(body.pw || '')) !== ASM_AUTH[id]) return {ok:false, error:'Not allowed'};
  }
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    if (act === 'det_publish' || act === 'det_mine') {
      var ss2 = ASM_SHEET_ID ? SpreadsheetApp.openById(ASM_SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
      var ds = ss2.getSheetByName(DET_TAB) || ss2.insertSheet(DET_TAB);
      if (act === 'det_publish') {
        var data = body.data || {}, fields = body.fields || [], out = [['Rider ID','Details (JSON)','Fields (JSON)','Updated']];
        Object.keys(data).forEach(function(k){ if (/^CPB\d+$/i.test(k)) out.push([k.toUpperCase(), JSON.stringify(data[k]), JSON.stringify(fields), Date.now()]); });
        ds.clear(); ds.getRange(1, 1, out.length, 4).setNumberFormat('@').setValues(out); ds.setFrozenRows(1);
        return {ok:true, count:out.length - 1};
      }
      var dv = ds.getDataRange().getValues();
      for (var q = 1; q < dv.length; q++) if (String(dv[q][0]).toUpperCase() === who) return {ok:true, data:JSON.parse(dv[q][1]), fields:JSON.parse(dv[q][2] || '[]')};
      return {ok:true, data:null};
    }
    var sh = asmSheet_(), vals = sh.getDataRange().getValues();
    if (act === 'assess_list' || act === 'assess_mine') {
      var rows = [];
      for (var i = 1; i < vals.length; i++) {
        var r = vals[i]; if (!r[0]) continue;
        if (mine && String(r[0]).toUpperCase() !== who) continue;
        var s = {}; ASM_KEYS.forEach(function(k, j){ s[k] = Number(r[2 + j]) || 0; });
        var d = r[1] instanceof Date ? Utilities.formatDate(r[1], 'Asia/Kuala_Lumpur', 'yyyy-MM-dd') : String(r[1]);
        rows.push({rid:String(r[0]), date:d, s:s, note:String(r[10] || ''), by:String(r[11] || ''), ts:Number(r[12]) || 0});
      }
      return {ok:true, rows:rows};
    }
    function find(rid, date){
      for (var i = 1; i < vals.length; i++) {
        var d = vals[i][1] instanceof Date ? Utilities.formatDate(vals[i][1], 'Asia/Kuala_Lumpur', 'yyyy-MM-dd') : String(vals[i][1]);
        if (String(vals[i][0]) === rid && d === date) return i + 1;
      }
      return 0;
    }
    if (body.action === 'assess_save') {
      var rec = body.rec || {};
      var rid = String(rec.rid || ''), date = String(rec.date || '');
      if (!/^CPB\d+$/i.test(rid) || !/^\d{4}-\d{2}-\d{2}$/.test(date)) return {ok:false, error:'Bad rider or date'};
      var row = [rid, "'" + date];
      for (var k = 0; k < ASM_KEYS.length; k++) {
        var v = Math.round(Number(rec.s && rec.s[ASM_KEYS[k]]));
        if (isNaN(v) || v < 0 || v > 100) return {ok:false, error:'Scores must be 0-100'};
        row.push(v);
      }
      row.push(String(rec.note || '').slice(0, 400), id, Date.now());
      var at = find(rid, date);
      if (at) sh.getRange(at, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
      return {ok:true};
    }
    if (body.action === 'assess_delete') {
      var at2 = find(String(body.rid || ''), String(body.date || ''));
      if (at2) sh.deleteRow(at2);
      return {ok:true};
    }
    return {ok:false, error:'Unknown action'};
  } finally { lock.releaseLock(); }
}
