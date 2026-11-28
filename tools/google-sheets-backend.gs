/**
 * 多摩市 高齢者施設マップ ─ 共有記録バックエンド（Google Apps Script Webアプリ）
 *
 * 使い方（1回だけ）:
 *  1) 新しいGoogleスプレッドシートを作成（https://sheets.new ）
 *  2) 拡張機能 > Apps Script を開き、このコードを全部貼り付けて保存
 *  3) デプロイ > 新しいデプロイ > 種類「ウェブアプリ」
 *       - 次のユーザーとして実行: 自分
 *       - アクセスできるユーザー: 全員（匿名を含む）
 *     デプロイ → 表示された「ウェブアプリのURL（/exec で終わる）」をコピー
 *  4) そのURLを担当者（Claude）に伝える → 地図に設定して公開
 *
 * シートは自動作成されます（records＝現在の状態 / log＝変更履歴）。
 * 誰が操作したかは、地図側で入力した「操作者名」を各行・各履歴に記録します。
 */
var RECORDS = 'records';
var LOG = 'log';
var REC_COLS = ['no','name','status','contactDate','visitDate','staff','yomiawase','checks','suit','memo','operator','updatedAt'];
var LOG_COLS = ['timestamp','operator','no','name','status','visitDate','memo'];

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name, headers){
  var s = ss_().getSheetByName(name);
  if(!s){ s = ss_().insertSheet(name); }
  if(s.getLastRow() === 0){ s.appendRow(headers); }
  return s;
}

function doGet(e){
  var cb = (e && e.parameter && e.parameter.callback) || '';
  var out;
  try { out = { ok:true, records: listRecords_() }; }
  catch(err){ out = { ok:false, error:String(err) }; }
  return reply_(out, cb);
}

function doPost(e){
  try {
    var body = JSON.parse(e.postData.contents);
    var r = body.record || {};
    var op = (body.operator || '').toString().slice(0, 40);
    upsert_(r, op);
    return reply_({ ok:true }, '');
  } catch(err){
    return reply_({ ok:false, error:String(err) }, '');
  }
}

function reply_(obj, cb){
  var json = JSON.stringify(obj);
  if(cb){
    return ContentService.createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function listRecords_(){
  var s = sheet_(RECORDS, REC_COLS);
  var vals = s.getDataRange().getValues();
  if(vals.length < 2) return [];
  var head = vals[0], out = [];
  for(var i=1;i<vals.length;i++){
    var row = vals[i];
    if(row[0] === '' || row[0] == null) continue;
    var o = {};
    for(var j=0;j<head.length;j++){ o[head[j]] = row[j]; }
    if(o.checks && typeof o.checks === 'string'){
      try { o.checks = JSON.parse(o.checks); } catch(e){ o.checks = {}; }
    }
    o.no = Number(o.no) || o.no;
    out.push(o);
  }
  return out;
}

function upsert_(r, op){
  var s = sheet_(RECORDS, REC_COLS);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  r.operator = op;
  r.updatedAt = now;
  var checksStr = (r.checks && typeof r.checks === 'object') ? JSON.stringify(r.checks) : (r.checks || '');
  var head = s.getRange(1,1,1,REC_COLS.length).getValues()[0];
  var rowArr = head.map(function(h){
    if(h === 'checks') return checksStr;
    return (r[h] != null) ? r[h] : '';
  });
  var vals = s.getDataRange().getValues();
  var found = -1;
  for(var i=1;i<vals.length;i++){ if(String(vals[i][0]) === String(r.no)){ found = i+1; break; } }
  if(found > 0){ s.getRange(found, 1, 1, head.length).setValues([rowArr]); }
  else { s.appendRow(rowArr); }

  var ls = sheet_(LOG, LOG_COLS);
  ls.appendRow([now, op, r.no, r.name || '', r.status || '', r.visitDate || '', r.memo || '']);
}
