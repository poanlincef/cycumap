/**
 * 中原營會地圖｜雲端存檔（Google Apps Script）
 *
 * 用法：
 * 1. 新增一個 Google 試算表 → 擴充功能 → Apps Script，把這整份貼上。
 * 2. （選填）把下面 TEAM_KEY 改成一組團隊密碼；留空代表不檢查。
 * 3. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *    執行身分：我；存取權：任何人 → 部署，複製網址（…/exec）。
 * 4. 把網址填進地圖頁面「雲端存檔 → 設定」，或寫在同資料夾的 config.js。
 *
 * 資料會存在試算表的「maps」分頁：每列一個圖層（或建築主檔），同名就覆蓋。
 */
const TEAM_KEY = 'siapublic';          // 例如 'cym2027'；留空＝不檢查
const SHEET = 'maps';
const HEAD = ['name', 'kind', 'json', 'updated', 'by'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function rows_() {
  const sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, HEAD.length).getValues().map((r, i) => ({
    row: i + 2, name: String(r[0]), kind: String(r[1]), json: String(r[2]),
    updated: r[3] instanceof Date ? r[3].toISOString() : String(r[3]), by: String(r[4])
  }));
}
/** 在編輯器裡選這個函式按「執行」，可先授權並建立 maps 分頁 */
function setup() { sheet_(); Logger.log('OK：maps 分頁已就緒'); }

function checkKey_(k) { return !TEAM_KEY || k === TEAM_KEY; }

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!checkKey_(p.key || '')) return out_({ ok: false, error: '團隊密碼不對' });
  const all = rows_();
  if (p.action === 'list') return out_({ ok: true, items: all.map(r => ({ name: r.name, kind: r.kind, updated: r.updated, by: r.by })) });
  if (p.action === 'get') {
    const r = all.find(x => x.name === p.name);
    return r ? out_({ ok: true, item: r }) : out_({ ok: false, error: '找不到 ' + p.name });
  }
  if (p.action === 'all') return out_({ ok: true, items: all });
  return out_({ ok: true, hello: '中原營會地圖雲端存檔' });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: '格式錯誤' }); }
  if (!checkKey_(b.key || '')) return out_({ ok: false, error: '團隊密碼不對' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(), all = rows_(), hit = all.find(x => x.name === b.name);
    if (b.action === 'save') {
      if (!b.name) return out_({ ok: false, error: '缺少名稱' });
      if (String(b.json).length > 49000) return out_({ ok: false, error: '這一層資料太大（超過試算表單格上限）' });
      const row = [b.name, b.kind || 'layer', b.json, new Date(), b.by || ''];
      if (hit) sh.getRange(hit.row, 1, 1, HEAD.length).setValues([row]); else sh.appendRow(row);
      return out_({ ok: true });
    }
    if (b.action === 'delete') {
      if (hit) sh.deleteRow(hit.row);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: '未知動作' });
  } finally { lock.releaseLock(); }
}
