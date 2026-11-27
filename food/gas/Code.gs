/**
 * 食材きげん帳 - 家族共有用サーバー (Google Apps Script)
 *
 * 使い方は food/README.md を参照してください。
 * データはこのスクリプトを紐づけたスプレッドシートの「items」シートに保存されます。
 */

// 家族だけが使えるようにする合言葉（任意）。空文字のままなら合言葉なしで使えます。
// 設定した場合は、アプリの設定画面にも同じ合言葉を入力してください。
const PASSCODE = '';

const SHEET_NAME = 'items';
const HEADERS = [
  'id', 'name', 'expiry', 'kind', 'location', 'qty', 'unit', 'memo',
  'addedBy', 'updatedBy', 'createdAt', 'updatedAt',
];
const LOCATIONS = ['冷蔵', '冷凍', '野菜室', '常温'];
const KINDS = ['消費期限', '賞味期限'];

function doGet(e) {
  return handle_(e && e.parameter ? e.parameter : {});
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { /* ignore */ }
  return handle_(p);
}

function handle_(p) {
  try {
    if (PASSCODE && String(p.key || '') !== PASSCODE) {
      return out_({ ok: false, error: '合言葉が違います' });
    }
    const sh = getSheet_();
    const action = p.action || 'list';

    if (action === 'list') {
      return out_({ ok: true, items: readAll_(sh) });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (action === 'upsert') {
        upsert_(sh, sanitize_(p.item));
      } else if (action === 'delete') {
        remove_(sh, String(p.id || ''));
      } else {
        return out_({ ok: false, error: '不明な操作です: ' + action });
      }
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return out_({ ok: true, items: readAll_(sh) });
  } catch (err) {
    return out_({ ok: false, error: String(err && err.message || err) });
  }
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // 日付などが自動変換されないよう、全体を書式なしテキストに
    sh.getRange(1, 1, sh.getMaxRows(), HEADERS.length).setNumberFormat('@');
  }
  return sh;
}

function readAll_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();
  return values
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const o = {};
      HEADERS.forEach(function (h, i) {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        o[h] = v;
      });
      o.id = String(o.id);
      o.name = String(o.name);
      o.expiry = String(o.expiry);
      o.qty = Number(o.qty) || 1;
      o.createdAt = Number(o.createdAt) || 0;
      o.updatedAt = Number(o.updatedAt) || 0;
      return o;
    });
}

function sanitize_(item) {
  if (!item || typeof item !== 'object') throw new Error('データがありません');
  const s = function (v, max) { return String(v == null ? '' : v).slice(0, max); };
  const id = s(item.id, 64);
  const name = s(item.name, 60).trim();
  const expiry = s(item.expiry, 10);
  if (!id) throw new Error('id がありません');
  if (!name) throw new Error('食材名がありません');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(expiry)) throw new Error('期限日の形式が不正です');
  return {
    id: id,
    name: name,
    expiry: expiry,
    kind: KINDS.indexOf(item.kind) >= 0 ? item.kind : '消費期限',
    location: LOCATIONS.indexOf(item.location) >= 0 ? item.location : '冷蔵',
    qty: Math.max(1, Math.min(999, parseInt(item.qty, 10) || 1)),
    unit: s(item.unit, 10),
    memo: s(item.memo, 200),
    addedBy: s(item.addedBy, 20),
    updatedBy: s(item.updatedBy, 20),
    createdAt: Number(item.createdAt) || Date.now(),
    updatedAt: Number(item.updatedAt) || Date.now(),
  };
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) return i + 2;
  }
  return -1;
}

function upsert_(sh, item) {
  const row = HEADERS.map(function (h) {
    const v = item[h];
    // 先頭が = + - @ の文字列はスプレッドシートの数式として解釈されないようにする
    return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
  });
  const r = findRow_(sh, item.id);
  if (r > 0) {
    sh.getRange(r, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
  } else {
    const next = sh.getLastRow() + 1;
    sh.getRange(next, 1, 1, HEADERS.length).setNumberFormat('@').setValues([row]);
  }
}

function remove_(sh, id) {
  if (!id) return;
  const r = findRow_(sh, id);
  if (r > 0) sh.deleteRow(r);
}
