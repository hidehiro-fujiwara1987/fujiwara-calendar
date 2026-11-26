// Apps Script のロジックを、Google のサービスを模したモックで確認する
// 実行: node tests/gas.test.js
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function makeEnv({ now, activeEmail = '', ownerEmail = 'owner@example.com' }) {
  const sheets = {};
  function sheet(name) {
    const rows = [];
    return {
      name, rows,
      getRange(a, b, c, d) {
        if (typeof a === 'string') return { setNumberFormat() {} };
        const r0 = a, c0 = b, nr = c || 1, nc = d || 1;
        return {
          getValue() { return (rows[r0 - 1] || [])[c0 - 1]; },
          getValues() { const o = []; for (let i = 0; i < nr; i++) { const row = rows[r0 - 1 + i] || []; o.push(Array.from({ length: nc }, (_, j) => row[c0 - 1 + j] ?? '')); } return o; },
          setValues(v) { v.forEach((row, i) => { rows[r0 - 1 + i] = rows[r0 - 1 + i] || []; row.forEach((x, j) => { rows[r0 - 1 + i][c0 - 1 + j] = x; }); }); },
          clearContent() { for (let i = 0; i < nr; i++) rows[r0 - 1 + i] = []; },
        };
      },
      getLastColumn() { return Math.max(0, ...rows.map(r => (r || []).length)); },
      getLastRow() { let n = rows.length; while (n > 0 && (!rows[n - 1] || rows[n - 1].every(x => x === '' || x == null))) n--; return n; },
      appendRow(r) { rows.length = this.getLastRow(); rows.push(r.slice()); },
      setFrozenRows() {},
    };
  }
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = sheet(n)) };
  const cache = {}; const props = {};
  const ctx = {
    Date: class extends Date { constructor(...a) { a.length ? super(...a) : super(env.now); } static now() { return +new Date(env.now); } },
    JSON, Object, String, Number, Math, Error,
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() {} },
    CacheService: { getScriptCache: () => ({ get: k => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: k => { delete cache[k]; } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => env.activeEmail }), getEffectiveUser: () => ({ getEmail: () => ownerEmail }) },
    Utilities: { formatDate: (d, tz, f) => f === 'H'
      ? String(Number(new Intl.DateTimeFormat('en-US', { timeZone: tz, hour: '2-digit', hourCycle: 'h23' }).format(d)))
      : new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit' }).format(d) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType: () => ({ body: s }) }) },
    Logger: { log() {} },
  };
  const env = { now, activeEmail, ctx, sheets, cache };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../gas/Code.gs'), 'utf8'), ctx);
  return env;
}

let n = 0;
function t(name, fn) { fn(); n++; console.log('ok -', name); }
const get = env => JSON.parse(env.ctx.doGet({ parameter: {} }).body);
const post = (env, b) => env.ctx.doPost({ postData: { contents: JSON.stringify(b) } });

t('何も入力していなければ全日○（daysが空）', () => {
  const env = makeEnv({ now: '2026-11-26T01:00:00Z' });
  const d = get(env);
  assert.strictEqual(d.today, '2026-11-26');
  assert.deepStrictEqual(d.days, {});
});

t('第三者は保存できない', () => {
  const env = makeEnv({ now: '2026-11-26T01:00:00Z', activeEmail: '' });
  assert.throws(() => env.ctx.saveStatuses({ '2026-11-27': 'full' }), /権限/);
  const env2 = makeEnv({ now: '2026-11-26T01:00:00Z', activeEmail: 'someone@example.com' });
  assert.throws(() => env2.ctx.saveStatuses({ '2026-11-27': 'full' }), /権限/);
  assert.throws(() => env2.ctx.getAnalytics(), /権限/);
});

t('管理画面URLは第三者には表示されない', () => {
  const env = makeEnv({ now: '2026-11-26T01:00:00Z' });
  let html = '';
  env.ctx.HtmlService = { createHtmlOutput: s => { html = s; return { setTitle: () => ({}) }; } };
  env.ctx.doGet({ parameter: { page: 'admin' } });
  assert.match(html, /表示できません/);
});

t('保存した状態が公開データに反映、○に戻すと消える', () => {
  const env = makeEnv({ now: '2026-11-26T01:00:00Z', activeEmail: 'owner@example.com' });
  get(env); // キャッシュを作る
  env.ctx.saveStatuses({ '2026-11-27': 'full', '2026-11-28': 'am', '2026-12-26': 'off', '2026-12-27': 'few' });
  let d = get(env);
  assert.deepStrictEqual(d.days, { '2026-11-27': 'full', '2026-11-28': 'am', '2026-12-26': 'off' }, '30日先より後は公開しない');
  assert.ok(d.updatedAt);
  env.ctx.saveStatuses({ '2026-11-27': 'ok' });
  d = get(env);
  assert.ok(!('2026-11-27' in d.days));
  assert.strictEqual(env.sheets['予約状況'].rows[1][1], '午前', 'シートには記号で保存');
});

t('不正な値は拒否', () => {
  const env = makeEnv({ now: '2026-11-26T01:00:00Z', activeEmail: 'owner@example.com' });
  assert.throws(() => env.ctx.saveStatuses({ '2026-02-30': 'full' }));
  assert.throws(() => env.ctx.saveStatuses({ '2026-11-27': 'xxx' }));
});

t('日付が進むと過去日は公開されない（年またぎ）', () => {
  const env = makeEnv({ now: '2026-12-31T01:00:00Z', activeEmail: 'owner@example.com' });
  env.ctx.saveStatuses({ '2026-12-31': 'few', '2027-01-01': 'off' });
  env.now = '2026-12-31T16:00:00Z'; // = 2027-01-01 01:00 JST
  const d = get(env);
  assert.strictEqual(d.today, '2027-01-01');
  assert.deepStrictEqual(d.days, { '2027-01-01': 'off' });
  env.ctx.dailyCleanup();
  assert.strictEqual(env.sheets['予約状況'].getLastRow(), 2);
});

t('アクセス集計（日別・月別・LINE・流入元・端末・時間帯・ボタン）', () => {
  const env = makeEnv({ now: '2026-11-25T03:00:00Z', activeEmail: 'owner@example.com' });
  // 旧版ページ（src なし、line だけ）
  post(env, { type: 'pv', newDay: 1, newMonth: 1, line: 1 });                                            // 11/25 12時 A
  env.now = '2026-11-26T12:30:00Z';                                                                       // 11/26 21:30 JST
  post(env, { type: 'pv', newDay: 1, newMonth: 0, newEver: 0, src: 'menu', device: 'ios' });             // A
  post(env, { type: 'pv', newDay: 0, newMonth: 0, newEver: 0, src: 'menu', device: 'ios' });             // A 再訪
  post(env, { type: 'pv', newDay: 1, newMonth: 1, newEver: 1, src: 'other', device: 'android' });        // B はじめて
  post(env, { type: 'click', target: 'line' });
  post(env, { type: 'click', target: 'tel' });
  post(env, { type: 'click', target: 'line' });
  post(env, { type: 'click', target: 'xxx' });                                                            // 無視
  post(env, { type: 'evil', newDay: 1 });                                                                 // 無視
  post(env, { type: 'pv', newDay: 1, src: '<script>', device: 'toaster' });                               // 不明値は other 扱い
  const a = env.ctx.getAnalytics();
  assert.strictEqual(a.todayPv, 4);
  assert.strictEqual(a.todayUv, 3);
  assert.strictEqual(a.monthPv, 5);
  assert.strictEqual(a.monthUv, 2);
  assert.strictEqual(a.monthLinePv, 3);
  assert.strictEqual(a.monthLineUv, 2);
  assert.strictEqual(a.monthClickLine, 2);
  assert.strictEqual(a.monthClickTel, 1);
  assert.strictEqual(a.series.length, 30);
  assert.strictEqual(a.series[29].date, '2026-11-26');
  assert.strictEqual(a.series[28].pv, 1);
  const L = a.last30;
  assert.strictEqual(L.pv, 5);
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.src)), { menu: 2, news: 0, profile: 0, line: 1, other: 2 });
  assert.deepStrictEqual(JSON.parse(JSON.stringify(L.dev)), { ios: 1, android: 1, other: 2 });
  assert.strictEqual(L.newEver, 1);
  assert.strictEqual(L.hours[21], 4);
  assert.strictEqual(L.hours[12], 1);
  assert.strictEqual(L.weekdays[4], 4); // 木曜
  assert.strictEqual(L.weekdays[3], 1); // 水曜
  assert.strictEqual(env.sheets['アクセス'].rows[0].length, 41, '見出しが41列');
});

t('旧形式（6列）の行が残っていても読める', () => {
  const env = makeEnv({ now: '2026-11-26T03:00:00Z', activeEmail: 'owner@example.com' });
  const sh = env.ctx.SpreadsheetApp.getActiveSpreadsheet().insertSheet('アクセス');
  sh.rows.push(['日付', '閲覧数', '訪問者(推定)', '今月初の訪問者(推定)', 'LINE経由閲覧数', 'LINE経由訪問者(推定)']);
  sh.rows.push(['2026-11-26', 5, 3, 3, 2, 1]);
  post(env, { type: 'pv', newDay: 1, newMonth: 1, newEver: 1, src: 'news', device: 'ios' });
  const a = env.ctx.getAnalytics();
  assert.strictEqual(a.todayPv, 6);
  assert.strictEqual(a.todayUv, 4);
  assert.strictEqual(a.last30.src.news, 1);
  assert.strictEqual(sh.rows[0].length, 41);
  assert.strictEqual(sh.rows[0][40], '23時');
});

console.log(`\n${n} tests passed`);
