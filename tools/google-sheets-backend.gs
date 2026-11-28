/**
 * 多摩市 高齢者施設マップ　チーム共有用（Google スプレッドシート ＋ Apps Script）
 *
 * このスプレッドシートが「記録の保管庫」です。地図（GitHub Pages）は誰でも開けますが、
 * 記録の読み書きは「合言葉」を知っている人だけができます（URLが知られても合言葉がなければ不可）。
 *
 *  - シート「記録」：施設ごとに1行。同じ施設を同時に更新したら、更新時刻が新しいほうが残ります。
 *  - シート「履歴」：更新のたびに、記録者名つきで1件1行ずつ積み上げます（消えません）。
 *  - 右端の隠し列 ts / data は地図用です。手で触らないでください。
 *
 * 準備のしかたは tools/google-sheets-setup.md を見てください（1回だけ・5分ほど・無料）。
 */

var SHEET_REC = '記録';
var SHEET_LOG = '履歴';
var REC_HEAD = ['No', '施設名', '段階番号', '段階', '適性', '連絡した日', '訪問日', '先方の担当',
  '読み合わせ会', '電話で確認', 'メモ', '更新日時', '更新者', 'ts', 'data(地図用・編集しない)'];
var LOG_HEAD = ['日時', 'No', '施設名', '段階', 'メモ', '記録者', 'シートに届いた日時'];
var COL_TS   = REC_HEAD.indexOf('ts') + 1;                     // 14
var COL_DATA = REC_HEAD.indexOf('data(地図用・編集しない)') + 1; // 15

/* ---------- メニュー ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('高齢者施設マップ')
    .addItem('① 最初の準備（シートを作る）', 'setupSheets')
    .addItem('② 合言葉を決める・変える', 'setTeamKey')
    .addToUi();
}

function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  var rec = ss.getSheetByName(SHEET_REC) || ss.insertSheet(SHEET_REC, 0);
  if (rec.getLastRow() === 0) {
    rec.appendRow(REC_HEAD);
    rec.setFrozenRows(1); rec.setFrozenColumns(2);
    rec.getRange(1, 1, 1, REC_HEAD.length).setFontWeight('bold').setBackground('#dff3e6');
    rec.getRange('A:A').setNumberFormat('@');
    rec.getRange('F:G').setNumberFormat('@');   // 日付は地図と同じ「2026-03-11」の形のまま
    rec.setColumnWidth(2, 220); rec.setColumnWidth(11, 300);
    rec.hideColumns(COL_TS, 2);
  }
  var log = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG, 1);
  if (log.getLastRow() === 0) {
    log.appendRow(LOG_HEAD);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, LOG_HEAD.length).setFontWeight('bold').setBackground('#fdecd9');
    log.setColumnWidth(3, 220); log.setColumnWidth(5, 320);
  }
  var def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  safeAlert('シートを準備しました。次に「② 合言葉を決める・変える」をしてください。');
}

function setTeamKey() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt('合言葉を決める',
    'チームで使う合言葉を入れてください（8文字以上。地図を使う人に伝えます）。\n変えると、全員がもう一度入れ直すことになります。',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var k = String(res.getResponseText() || '').trim();
  if (k.length < 8) { ui.alert('8文字以上にしてください。'); return; }
  PropertiesService.getScriptProperties().setProperty('TEAM_KEY', k);
  ui.alert('合言葉を保存しました。');
}

function safeAlert(msg) { try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); } }

/* ---------- 地図からの通信 ---------- */
function doGet() {
  return json_({ ok: true, app: 'tama-senior-map', note: '地図から使う窓口です。' });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_request' }); }
  var key = PropertiesService.getScriptProperties().getProperty('TEAM_KEY');
  if (!key) return json_({ ok: false, error: 'no_key', message: 'スプレッドシートで合言葉がまだ決まっていません' });
  if (!req || req.key !== key) { Utilities.sleep(1200); return json_({ ok: false, error: 'key', message: '合言葉がちがいます' }); }
  try {
    if (req.action === 'ping') return json_({ ok: true, serverTime: Date.now() });
    if (req.action === 'pull') return json_(pull_());
    if (req.action === 'push') return json_(push_(req));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', message: String(err && err.message || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function recSheet_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REC);
  if (!sh) { try { setupSheets(); } catch (e) {} sh = SpreadsheetApp.getActive().getSheetByName(SHEET_REC); }
  return sh;
}

function readAll_(sh) {
  var n = sh.getLastRow() - 1, out = {};
  if (n <= 0) return out;
  var nos = sh.getRange(2, 1, n, 1).getValues();
  var data = sh.getRange(2, COL_TS, n, 2).getValues();   // ts, data
  for (var i = 0; i < n; i++) {
    var no = String(nos[i][0] || '');
    if (!no) continue;
    var rec = null;
    try { rec = JSON.parse(data[i][1] || 'null'); } catch (e) { rec = null; }
    if (!rec) continue;
    rec.ts = Number(data[i][0]) || rec.ts || 0;
    out[no] = { row: i + 2, rec: rec };
  }
  return out;
}

function pull_() {
  var all = readAll_(recSheet_()), recs = {};
  Object.keys(all).forEach(function (no) { recs[no] = all[no].rec; });
  return { ok: true, records: recs, serverTime: Date.now() };
}

function push_(req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = recSheet_();
    var logSh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
    var all = readAll_(sh);
    var accepted = [], rejected = {}, newLogs = [], appendRows = [];
    var nowStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    var who = String(req.who || '').slice(0, 40);
    (req.items || []).forEach(function (it) {
      if (!it || it.no == null || !it.rec) return;
      var no = String(it.no), rec = it.rec, info = it.info || {}, cur = all[no];
      var inTs = Number(rec.ts) || 0;
      if (cur && (Number(cur.rec.ts) || 0) >= inTs) { rejected[no] = cur.rec; return; }  // シートのほうが新しい
      var row = rowOf_(no, rec, info, who);
      if (cur) { sh.getRange(cur.row, 1, 1, row.length).setValues([row]); }
      else { appendRows.push(row); }
      all[no] = { row: cur ? cur.row : -1, rec: rec };
      accepted.push(no);
      newLogs.push([rec.updatedAt || nowStr, no, info.name || '', info.stage || '',
        (rec.memo || '').toString().slice(0, 500), rec.operator || who, nowStr]);
    });
    if (appendRows.length) sh.getRange(sh.getLastRow() + 1, 1, appendRows.length, REC_HEAD.length).setValues(appendRows);
    if (newLogs.length && logSh) logSh.getRange(logSh.getLastRow() + 1, 1, newLogs.length, LOG_HEAD.length).setValues(newLogs);
    return { ok: true, accepted: accepted, rejected: rejected, serverTime: Date.now() };
  } finally {
    lock.releaseLock();
  }
}

function rowOf_(no, r, info, who) {
  return [
    no, info.name || '', (info.step == null ? '' : info.step), info.stage || '', info.suitText || '',
    r.contactDate || '', r.visitDate || '', r.staff || '', r.yomiawase || '', info.checksText || '',
    r.memo || '', r.updatedAt || '', r.operator || who || '', Number(r.ts) || 0, JSON.stringify(r)
  ].map(function (v) { return (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : v; });
}
