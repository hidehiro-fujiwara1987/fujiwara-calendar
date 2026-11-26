/**
 * 藤原整体院 予約状況カレンダー（Google Apps Script）
 *
 * このスクリプトはスプレッドシートに紐づけて使います（拡張機能 → Apps Script）。
 *  - GET  /exec              … 公開ページ用の予約状況 JSON（誰でも読める・読み取り専用）
 *  - POST /exec              … 公開ページからの匿名アクセス記録（日別の数を足すだけ）
 *  - GET  /exec?page=admin   … 管理画面（スクリプト所有者の Google アカウントのみ）
 *
 * 保存するのは「○以外の日の状態」と「日別のアクセス数」だけ。個人情報は保存しません。
 */

var TZ = 'Asia/Tokyo';
var PUBLIC_DAYS_AHEAD = 30;   // 公開: 今日〜30日後
var ADMIN_DAYS_AHEAD = 60;    // 管理: 今日〜60日後
var STATUS_SHEET = '予約状況';
var ACCESS_SHEET = 'アクセス';
var CACHE_KEY = 'status-v1';
var CACHE_SEC = 600;

// 公開ページの URL（GitHub Pages など）。管理画面の「公開ページを見る」に使います。
var PUBLIC_PAGE_URL = '';

// 状態キー ⇔ シート上の表記。○は「行なし」で表す。
var STATUS_MARK = { few: '△', full: '×', am: '午前', pm: '午後', off: '休' };

/* ========== 入口 ========== */

function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'admin') {
    if (!isOwner_()) {
      return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">このページは表示できません。</p>')
        .setTitle('藤原整体院');
    }
    var t = HtmlService.createTemplateFromFile('admin');
    t.initialJson = JSON.stringify(adminData_()).replace(/</g, '\\u003c');
    return t.evaluate()
      .setTitle('予約状況の編集｜藤原整体院')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
  return json_(publicData_());
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body && body.type === 'pv') recordHit_(body);
  } catch (err) {
    // 集計の失敗は無視（公開ページには影響させない）
  }
  return json_({ ok: true });
}

/* ========== 管理画面から呼ぶ関数（所有者のみ） ========== */

/**
 * @param {Object} changes { 'YYYY-MM-DD': 'ok'|'few'|'full'|'am'|'pm'|'off' }
 */
function saveStatuses(changes) {
  assertOwner_();
  if (!changes || typeof changes !== 'object') throw new Error('送信内容が正しくありません。');
  var keys = Object.keys(changes);
  if (keys.length === 0 || keys.length > 400) throw new Error('送信内容が正しくありません。');
  keys.forEach(function (k) {
    var v = changes[k];
    if (!isKey_(k) || !(v === 'ok' || STATUS_MARK.hasOwnProperty(v))) throw new Error('送信内容が正しくありません（' + k + '）。');
  });

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('混み合っています。少し待ってから再送してください。');
  try {
    var today = todayKey_();
    var all = readStatuses_();
    keys.forEach(function (k) {
      if (changes[k] === 'ok') delete all[k];
      else all[k] = changes[k];
    });
    Object.keys(all).forEach(function (k) { if (k < today) delete all[k]; }); // 過去日は削除
    writeStatuses_(all);
    var updatedAt = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('updatedAt', updatedAt);
    CacheService.getScriptCache().remove(CACHE_KEY);
    SpreadsheetApp.flush();
    return { ok: true, updatedAt: updatedAt, today: today, days: all };
  } finally {
    lock.releaseLock();
  }
}

function getAdminData() {
  assertOwner_();
  return adminData_();
}

function getAnalytics() {
  assertOwner_();
  var today = todayKey_();
  var month = today.slice(0, 7);
  var rows = readAccess_();
  var byDay = {};
  rows.forEach(function (r) { byDay[r.date] = r; });

  var series = [];
  for (var i = 29; i >= 0; i--) {
    var k = addDays_(today, -i);
    var r = byDay[k] || {};
    series.push({ date: k, pv: r.pv || 0, uv: r.uv || 0 });
  }

  var m = { pv: 0, uv: 0, linePv: 0, lineUv: 0 };
  var prevMonth = addDays_(month + '-01', -1).slice(0, 7);
  var pm = { pv: 0, uv: 0 };
  rows.forEach(function (r) {
    var rm = r.date.slice(0, 7);
    if (rm === month) { m.pv += r.pv; m.uv += r.uvMonth; m.linePv += r.linePv; m.lineUv += r.lineUv; }
    if (rm === prevMonth) { pm.pv += r.pv; pm.uv += r.uvMonth; }
  });
  var t = byDay[today] || {};
  return {
    today: today,
    todayPv: t.pv || 0, todayUv: t.uv || 0, todayLinePv: t.linePv || 0,
    monthPv: m.pv, monthUv: m.uv, monthLinePv: m.linePv, monthLineUv: m.lineUv,
    prevMonthPv: pm.pv, prevMonthUv: pm.uv,
    series: series
  };
}

/* ========== 初期設定（エディタから1回だけ実行） ========== */

function setup() {
  statusSheet_();
  accessSheet_();
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'dailyCleanup') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('dailyCleanup').timeBased().atHour(1).everyDays(1).inTimezone(TZ).create();
  Logger.log('準備ができました。次に「デプロイ」から公開してください。');
}

// 毎日深夜に過去日の予約状況を消す（表示には影響しない掃除）
function dailyCleanup() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    var today = todayKey_();
    var all = readStatuses_();
    var changed = false;
    Object.keys(all).forEach(function (k) { if (k < today) { delete all[k]; changed = true; } });
    if (changed) writeStatuses_(all);
    CacheService.getScriptCache().remove(CACHE_KEY);
  } finally {
    lock.releaseLock();
  }
}

/* ========== 内部処理 ========== */

function publicData_() {
  var today = todayKey_();
  var end = addDays_(today, PUBLIC_DAYS_AHEAD);
  var cache = CacheService.getScriptCache();
  var payload = null;
  var hit = cache.get(CACHE_KEY);
  if (hit) { try { payload = JSON.parse(hit); } catch (e) { payload = null; } }
  if (!payload) {
    payload = { updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt'), all: readStatuses_() };
    cache.put(CACHE_KEY, JSON.stringify(payload), CACHE_SEC);
  }
  var days = {};
  Object.keys(payload.all).forEach(function (k) { if (k >= today && k <= end) days[k] = payload.all[k]; });
  return { ok: true, today: today, daysAhead: PUBLIC_DAYS_AHEAD, updatedAt: payload.updatedAt || null, days: days };
}

function adminData_() {
  var today = todayKey_();
  var end = addDays_(today, ADMIN_DAYS_AHEAD);
  var all = readStatuses_();
  var days = {};
  Object.keys(all).forEach(function (k) { if (k >= today && k <= end) days[k] = all[k]; });
  return {
    today: today, daysAhead: ADMIN_DAYS_AHEAD, days: days,
    updatedAt: PropertiesService.getScriptProperties().getProperty('updatedAt'),
    publicUrl: PUBLIC_PAGE_URL
  };
}

function recordHit_(b) {
  var newDay = b.newDay === 1 ? 1 : 0;
  var newMonth = b.newMonth === 1 ? 1 : 0;
  var line = b.line === 1 ? 1 : 0;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sh = accessSheet_();
    var today = todayKey_();
    var last = sh.getLastRow();
    var row = -1;
    if (last >= 2) {
      // 今日の行は通常いちばん下
      var v = sh.getRange(last, 1).getValue();
      if (normKey_(v) === today) row = last;
    }
    if (row < 0) {
      sh.appendRow([today, 0, 0, 0, 0, 0]);
      row = sh.getLastRow();
    }
    var r = sh.getRange(row, 2, 1, 5);
    var c = r.getValues()[0].map(function (x) { return Number(x) || 0; });
    // 閲覧数, 訪問者(推定), 今月はじめての訪問者(推定), LINE経由閲覧数, LINE経由訪問者(推定)
    r.setValues([[c[0] + 1, c[1] + newDay, c[2] + newMonth, c[3] + line, c[4] + (line && newDay ? 1 : 0)]]);
  } finally {
    lock.releaseLock();
  }
}

function readAccess_() {
  var sh = accessSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues().map(function (r) {
    return { date: normKey_(r[0]), pv: +r[1] || 0, uv: +r[2] || 0, uvMonth: +r[3] || 0, linePv: +r[4] || 0, lineUv: +r[5] || 0 };
  }).filter(function (r) { return isKey_(r.date); });
}

function readStatuses_() {
  var sh = statusSheet_();
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  var markToKey = {};
  Object.keys(STATUS_MARK).forEach(function (k) { markToKey[STATUS_MARK[k]] = k; markToKey[k] = k; });
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
    var k = normKey_(r[0]);
    var s = markToKey[String(r[1]).trim()];
    if (isKey_(k) && s) out[k] = s;
  });
  return out;
}

function writeStatuses_(all) {
  var sh = statusSheet_();
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 2).clearContent();
  var keys = Object.keys(all).sort();
  if (keys.length) {
    sh.getRange(2, 1, keys.length, 2).setValues(keys.map(function (k) { return [k, STATUS_MARK[all[k]]]; }));
  }
}

function statusSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STATUS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STATUS_SHEET);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, 3).setValues([['日付', '状態', '※ ○の日は行がありません。編集は管理画面から行ってください。']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function accessSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ACCESS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACCESS_SHEET);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange(1, 1, 1, 6).setValues([['日付', '閲覧数', '訪問者(推定)', '今月初の訪問者(推定)', 'LINE経由閲覧数', 'LINE経由訪問者(推定)']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function isOwner_() {
  var active = Session.getActiveUser().getEmail();
  var owner = Session.getEffectiveUser().getEmail();
  return !!active && active === owner;
}

function assertOwner_() {
  if (!isOwner_()) throw new Error('この操作を行う権限がありません。');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function todayKey_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
}

function normKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).trim();
}

function isKey_(k) {
  if (typeof k !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(k)) return false;
  var a = k.split('-').map(Number);
  var d = new Date(Date.UTC(a[0], a[1] - 1, a[2]));
  return d.getUTCFullYear() === a[0] && d.getUTCMonth() === a[1] - 1 && d.getUTCDate() === a[2];
}

function addDays_(key, n) {
  var a = key.split('-').map(Number);
  return new Date(Date.UTC(a[0], a[1] - 1, a[2]) + n * 86400000).toISOString().slice(0, 10);
}
