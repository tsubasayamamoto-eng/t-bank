/**
 * BUTASAN データ保存用（Google スプレッドシートの「拡張機能 → Apps Script」に貼りつける）
 *
 * - シート「data」に 1行 = 1データ（path / json / updatedAt）で保存します。
 * - 合言葉は「プロジェクトの設定 → スクリプト プロパティ」に PASSCODE という名前で入れてください。
 *   合言葉が合わないリクエストには何も返しません。
 * - 削除したデータは json を空にして残し（updatedAt だけ更新）、ほかのスマホに「消えた」ことを伝えます。
 */
const SHEET_NAME = 'data';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['path', 'json', 'updatedAt']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function authed_(key) {
  const pass = PropertiesService.getScriptProperties().getProperty('PASSCODE');
  return !!pass && typeof key === 'string' && key === pass;
}

function rows_() {
  const values = sheet_().getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) rows.push({row: i + 1, path: String(values[i][0]), json: String(values[i][1] || ''), t: Number(values[i][2]) || 0});
  }
  return rows;
}

// 読む：?key=合言葉&since=前回のmax（0ならぜんぶ）
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!authed_(p.key)) return out_({ok: false, error: 'auth'});
  const since = Number(p.since) || 0;
  const docs = {};
  let max = 0;
  rows_().forEach(function (r) {
    if (r.t > max) max = r.t;
    if (r.t <= since) return;
    if (!r.json) { if (since) docs[r.path] = null; return; }
    try { docs[r.path] = JSON.parse(r.json); } catch (err) {}
  });
  return out_({ok: true, docs: docs, max: max});
}

// 書く：{key, ops:[{op:'set'|'delete', path, data}]}
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ok: false, error: 'bad'}); }
  if (!authed_(body.key)) return out_({ok: false, error: 'auth'});
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    const index = {};
    rows_().forEach(function (r) { index[r.path] = r.row; });
    let t = Date.now();
    (body.ops || []).slice(0, 200).forEach(function (op) {
      if (!op || typeof op.path !== 'string' || !op.path || op.path.length > 300) return;
      const json = op.op === 'delete' ? '' : JSON.stringify(op.data || {});
      if (json.length > 45000) return;
      t++;
      if (index[op.path]) sh.getRange(index[op.path], 2, 1, 2).setValues([[json, t]]);
      else { sh.appendRow([op.path, json, t]); index[op.path] = sh.getLastRow(); }
    });
    return out_({ok: true, max: t});
  } finally {
    lock.releaseLock();
  }
}
