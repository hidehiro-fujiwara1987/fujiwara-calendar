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
var PUBLIC_PAGE_URL = 'https://hidehiro-fujiwara1987.github.io/fujiwara-calendar/';

// アクセスシートの列（すべて日別の「回数」だけ。個人を識別する情報は持たない）
var ACCESS_COLS = ['日付', '閲覧数', '訪問者(推定)', '今月初の訪問者(推定)', 'LINE経由閲覧数', 'LINE経由訪問者(推定)',
  '初めての訪問者(推定)', 'LINEボタン', '電話ボタン', 'iPhone', 'Android', 'その他端末',
  '流入:LINEメニュー', '流入:LINE配信', '流入:LINEプロフィール', '流入:LINEその他', '流入:LINE以外'];
for (var h_ = 0; h_ < 24; h_++) ACCESS_COLS.push(h_ + '時');
var AC = { pv: 2, uv: 3, uvMonth: 4, linePv: 5, lineUv: 6, newEver: 7, clickLine: 8, clickTel: 9,
  ios: 10, android: 11, otherDev: 12, srcMenu: 13, srcNews: 14, srcProfile: 15, srcLine: 16, srcOther: 17, hour0: 18 };
var SRC_COL = { menu: AC.srcMenu, news: AC.srcNews, profile: AC.srcProfile, line: AC.srcLine, other: AC.srcOther };
var DEV_COL = { ios: AC.ios, android: AC.android, other: AC.otherDev };

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
    if (body && body.type === 'pv') recordPv_(body);
    else if (body && body.type === 'click') recordClick_(body);
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
  var from = addDays_(today, -29);
  var L = { pv: 0, uv: 0, newEver: 0, hours: [], weekdays: [0, 0, 0, 0, 0, 0, 0],
    src: { menu: 0, news: 0, profile: 0, line: 0, other: 0 }, dev: { ios: 0, android: 0, other: 0 } };
  for (var h = 0; h < 24; h++) L.hours.push(0);
  var mc = { line: 0, tel: 0 };
  rows.forEach(function (r) {
    if (r.date.slice(0, 7) === month) { mc.line += r.clickLine; mc.tel += r.clickTel; }
    if (r.date < from || r.date > today) return;
    L.pv += r.pv; L.uv += r.uv; L.newEver += r.newEver;
    r.hours.forEach(function (v, i) { L.hours[i] += v; });
    var a = r.date.split('-').map(Number);
    L.weekdays[new Date(Date.UTC(a[0], a[1] - 1, a[2])).getUTCDay()] += r.pv;
    L.src.menu += r.srcMenu; L.src.news += r.srcNews; L.src.profile += r.srcProfile; L.src.line += r.srcLine; L.src.other += r.srcOther;
    L.dev.ios += r.ios; L.dev.android += r.android; L.dev.other += r.otherDev;
  });
  var t = byDay[today] || {};
  return {
    monthClickLine: mc.line, monthClickTel: mc.tel, last30: L,
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

function recordPv_(b) {
  var newDay = b.newDay === 1, newMonth = b.newMonth === 1, newEver = b.newEver === 1;
  // 古い版のページは src を送らず line だけ送ってくる
  var src = SRC_COL.hasOwnProperty(b.src) ? b.src : (b.line === 1 ? 'line' : 'other');
  var isLine = src !== 'other';
  var dev = DEV_COL.hasOwnProperty(b.device) ? b.device : 'other';
  var inc = {};
  inc[AC.pv] = 1;
  if (newDay) inc[AC.uv] = 1;
  if (newMonth) inc[AC.uvMonth] = 1;
  if (isLine) inc[AC.linePv] = 1;
  if (isLine && newDay) inc[AC.lineUv] = 1;
  if (newEver) inc[AC.newEver] = 1;
  if (newDay) inc[DEV_COL[dev]] = 1;            // 端末は1日1回だけ数える（訪問者ベース）
  inc[SRC_COL[src]] = 1;
  inc[AC.hour0 + Number(Utilities.formatDate(new Date(), TZ, 'H'))] = 1;
  increment_(inc);
}

function recordClick_(b) {
  var col = b.target === 'line' ? AC.clickLine : b.target === 'tel' ? AC.clickTel : 0;
  if (!col) return;
  var inc = {};
  inc[col] = 1;
  increment_(inc);
}

// 今日の行の指定列に数を足す
function increment_(inc) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var sh = accessSheet_();
    var today = todayKey_();
    var n = ACCESS_COLS.length;
    var last = sh.getLastRow();
    var row = -1;
    if (last >= 2 && normKey_(sh.getRange(last, 1).getValue()) === today) row = last; // 今日の行は通常いちばん下
    if (row < 0) {
      var blank = [today];
      for (var i = 1; i < n; i++) blank.push(0);
      sh.appendRow(blank);
      row = sh.getLastRow();
    }
    var r = sh.getRange(row, 2, 1, n - 1);
    var v = r.getValues()[0].map(function (x) { return Number(x) || 0; });
    Object.keys(inc).forEach(function (c) { v[Number(c) - 2] += inc[c]; });
    r.setValues([v]);
  } finally {
    lock.releaseLock();
  }
}

function readAccess_() {
  var sh = accessSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ACCESS_COLS.length).getValues().map(function (r) {
    var o = { date: normKey_(r[0]), hours: [] };
    Object.keys(AC).forEach(function (k) { if (k !== 'hour0') o[k] = +r[AC[k] - 1] || 0; });
    for (var h = 0; h < 24; h++) o.hours.push(+r[AC.hour0 - 1 + h] || 0);
    return o;
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
    sh.setFrozenRows(1);
  }
  // 列が増えた版に合わせて見出しをそろえる（既存の数字はそのまま）
  if (sh.getLastColumn() < ACCESS_COLS.length) sh.getRange(1, 1, 1, ACCESS_COLS.length).setValues([ACCESS_COLS]);
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
