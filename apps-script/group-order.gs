/**
 * FEEDO · 公司揪團後端（Google Apps Script + Google 試算表）
 * ------------------------------------------------------------
 * 設定方式請看同資料夾的 README.md。
 * 資料存在綁定的試算表，會自動建立兩個工作表：
 *   groups  — 每一團（團名、發起人、取茶方式、截止時間、是否截止）
 *   entries — 每個人的點單（名字、品項）
 * 網站以 POST（text/plain JSON）呼叫，action：get / create / save / remove / close
 */

const GROUP_HEAD = ['id', 'admin', 'title', 'host', 'mode', 'deadline', 'closed', 'createdAt'];
const ENTRY_HEAD = ['groupId', 'entryId', 'editToken', 'name', 'items', 'updatedAt', 'summary'];
const MAX_ENTRIES = 200;     // 每團最多幾個人
const MAX_ITEMS = 50;        // 每人最多幾種規格
const MAX_QTY = 999;         // 每種規格最多幾杯
const MAX_DAYS = 14;         // 截止時間最多設定幾天後
const REOPEN_MINUTES = 30;   // 已過截止時間又重新開放時，自動延長幾分鐘
const RETENTION_DAYS = 30;   // 揪團截止後幾天自動刪除（隱私權說明寫的保存期間，改這裡要一起改說明）

function doGet() {
  return json_({ ok: true, service: 'feedo-group-order' });
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ error: 'bad' });
  }

  // 讀取不需要排隊
  if (body.action === 'get') {
    try {
      return json_(getGroup_(body.id, body.admin));
    } catch (err) {
      return errorJson_(err);
    }
  }

  // 寫入一次只處理一個，避免多人同時送出時寫錯列
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    switch (body.action) {
      case 'create': return json_(createGroup_(body));
      case 'save':   return json_(saveEntry_(body));
      case 'remove': return json_(removeEntry_(body));
      case 'close':  return json_(setClosed_(body));
      default:       return json_({ error: 'bad' });
    }
  } catch (err) {
    return errorJson_(err);
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 動作 ---------- */

function createGroup_(b) {
  const title = text_(b.title, 30);
  const host = text_(b.host, 20);
  const mode = b.mode === 'delivery' ? 'delivery' : 'pickup';
  const deadline = new Date(b.deadline);
  const now = new Date();
  if (!title || !host || isNaN(deadline.getTime())) fail_('bad');
  if (deadline <= now || deadline - now > MAX_DAYS * 864e5) fail_('bad');

  const id = token_(10);
  const admin = token_(32);
  appendRow_(sheet_('groups', GROUP_HEAD),
    [id, admin, title, host, mode, deadline.toISOString(), 'false', now.toISOString()]);
  return { id: id, admin: admin };
}

function getGroup_(id, admin) {
  const g = findGroup_(id);
  const entries = entryRows_(g.id).map(function (r) {
    return { id: String(r.entryId), name: String(r.name), items: parseItems_(r.items), updatedAt: toIso_(r.updatedAt) };
  });
  return {
    group: {
      id: String(g.id), title: String(g.title), host: String(g.host), mode: String(g.mode),
      deadline: toIso_(g.deadline), closed: isClosed_(g)
    },
    entries: entries,
    isAdmin: isAdmin_(g, admin)
  };
}

function saveEntry_(b) {
  const g = findGroup_(b.id);
  const admin = isAdmin_(g, b.admin);
  if (isClosed_(g) && !admin) fail_('closed');

  const name = text_(b.name, 20);
  const items = cleanItems_(b.items);
  if (!name || !items.length) fail_('bad');

  const sh = sheet_('entries', ENTRY_HEAD);
  const rows = entryRows_(g.id);
  const values = [name, JSON.stringify(items), new Date().toISOString(), text_(b.summary, 500)];

  const mine = b.entryId ? rows.find(function (r) { return String(r.entryId) === String(b.entryId); }) : null;
  if (mine) {
    if (String(mine.editToken) !== String(b.editToken)) fail_('auth');
    writeRow_(sh, mine._row, 4, values);
    return Object.assign({ entryId: String(mine.entryId), editToken: String(mine.editToken) }, getGroup_(g.id, b.admin));
  }

  // 沒有舊點單（或已被發起人刪除）就新增一筆
  if (rows.length >= MAX_ENTRIES) fail_('full');
  const entryId = token_(10);
  const editToken = token_(32);
  appendRow_(sh, [String(g.id), entryId, editToken].concat(values));
  return Object.assign({ entryId: entryId, editToken: editToken }, getGroup_(g.id, b.admin));
}

function removeEntry_(b) {
  const g = findGroup_(b.id);
  const admin = isAdmin_(g, b.admin);
  if (isClosed_(g) && !admin) fail_('closed');

  const row = entryRows_(g.id).find(function (r) { return String(r.entryId) === String(b.entryId); });
  if (row) {
    if (!admin && String(row.editToken) !== String(b.editToken)) fail_('auth');
    sheet_('entries', ENTRY_HEAD).deleteRow(row._row);
  }
  return getGroup_(g.id, b.admin);
}

function setClosed_(b) {
  const g = findGroup_(b.id);
  if (!isAdmin_(g, b.admin)) fail_('auth');

  const closing = b.closed !== false;
  let deadline = new Date(toIso_(g.deadline));
  if (!closing && deadline <= new Date()) deadline = new Date(Date.now() + REOPEN_MINUTES * 60000);
  writeRow_(sheet_('groups', GROUP_HEAD), g._row, 6, [deadline.toISOString(), String(closing)]);
  return getGroup_(g.id, b.admin);
}

/* ---------- 自動刪除 ---------- */

// 由每天的時間觸發器執行：刪除截止超過 RETENTION_DAYS 天的揪團，以及那些團的所有點單
function purgeExpired() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cut = Date.now() - RETENTION_DAYS * 864e5;
    const gs = sheet_('groups', GROUP_HEAD);
    const expired = rows_(gs).filter(function (g) {
      const d = new Date(toIso_(g.deadline)).getTime();
      return !isNaN(d) && d < cut;
    });
    if (!expired.length) return 0;

    const ids = {};
    expired.forEach(function (g) { ids[String(g.id)] = true; });
    const es = sheet_('entries', ENTRY_HEAD);
    deleteRows_(es, rows_(es).filter(function (r) { return ids[String(r.groupId)]; }));
    deleteRows_(gs, expired);
    return expired.length;
  } finally {
    lock.releaseLock();
  }
}

// 在 Apps Script 編輯器手動執行一次：建立每天台灣時間凌晨 3 點的自動刪除
// 重複執行只會留下一個觸發器
function installPurgeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'purgeExpired') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('purgeExpired').timeBased().everyDays(1).atHour(3).inTimezone('Asia/Taipei').create();
}

/* ---------- 試算表工具 ---------- */

// 由下往上刪，前面的列號才不會跑掉
function deleteRows_(sh, rows) {
  rows.map(function (r) { return r._row; })
    .sort(function (a, b) { return b - a; })
    .forEach(function (n) { sh.deleteRow(n); });
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, width).getValues()[0];
  return sh.getRange(2, 1, last - 1, width).getValues().map(function (r, k) {
    const o = { _row: k + 2 };
    head.forEach(function (h, c) { o[h] = r[c]; });
    return o;
  });
}

function findGroup_(id) {
  if (!id) fail_('notfound');
  const g = rows_(sheet_('groups', GROUP_HEAD)).find(function (r) { return String(r.id) === String(id); });
  if (!g) fail_('notfound');
  return g;
}

function entryRows_(groupId) {
  return rows_(sheet_('entries', ENTRY_HEAD)).filter(function (r) { return String(r.groupId) === String(groupId); });
}

function appendRow_(sh, values) {
  writeRow_(sh, sh.getLastRow() + 1, 1, values);
}

// 一律以純文字寫入：避免編號被當成數字、使用者輸入被當成公式
function writeRow_(sh, row, col, values) {
  sh.getRange(row, col, 1, values.length)
    .setNumberFormat('@')
    .setValues([values.map(function (v) {
      const s = String(v);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    })]);
}

/* ---------- 資料檢查 ---------- */

function cleanItems_(items) {
  if (!Array.isArray(items)) return [];
  return items.slice(0, MAX_ITEMS).map(function (it) {
    it = it || {};
    return {
      i: Math.floor(Number(it.i)),
      sweet: text_(it.sweet, 10),
      ice: text_(it.ice, 10),
      foam: it.foam === true,
      qty: Math.min(MAX_QTY, Math.max(1, Math.floor(Number(it.qty)) || 1)),
      note: text_(it.note, 60)
    };
  }).filter(function (it) { return it.i >= 0 && it.i < 200; });
}

function parseItems_(s) {
  try {
    const a = JSON.parse(String(s));
    return Array.isArray(a) ? a : [];
  } catch (e) {
    return [];
  }
}

function isClosed_(g) {
  return String(g.closed) === 'true' || new Date() > new Date(toIso_(g.deadline));
}

function isAdmin_(g, admin) {
  return !!admin && String(admin) === String(g.admin);
}

function text_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

function toIso_(v) {
  return v instanceof Date ? v.toISOString() : String(v);
}

function token_(len) {
  let s = '';
  while (s.length < len) s += Utilities.getUuid().replace(/-/g, '');
  return s.slice(0, len);
}

function fail_(code) {
  const e = new Error(code);
  e.code = code;
  throw e;
}

function errorJson_(err) {
  return json_({ error: (err && err.code) || 'server', message: String((err && err.message) || err) });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
