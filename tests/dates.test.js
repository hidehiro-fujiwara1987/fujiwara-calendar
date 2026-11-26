// 実行: node tests/dates.test.js
const assert = require('assert');
const D = require('../public/dates.js');

let n = 0;
function t(name, fn) { fn(); n++; console.log('ok -', name); }

const range = (start) => { const ds = D.buildDays(start, 31); return [ds[0].key, ds[30].key, ds.length]; };

t('今日から30日先まで（31日分）', () => {
  assert.deepStrictEqual(range('2026-11-26'), ['2026-11-26', '2026-12-26', 31]);
});
t('翌日になると1日ずつ進む', () => {
  assert.deepStrictEqual(range('2026-11-27'), ['2026-11-27', '2026-12-27', 31]);
});
t('年をまたぐ', () => {
  assert.deepStrictEqual(range('2026-12-15'), ['2026-12-15', '2027-01-14', 31]);
});
t('閏年の2月29日を含む', () => {
  const keys = D.buildDays('2028-02-20', 31).map(d => d.key);
  assert.ok(keys.includes('2028-02-29'));
  assert.ok(keys.includes('2028-03-01'));
});
t('平年は2月29日がない', () => {
  const keys = D.buildDays('2027-02-20', 31).map(d => d.key);
  assert.ok(!keys.includes('2027-02-29'));
});
t('曜日が正しい', () => {
  assert.strictEqual(D.buildDays('2026-11-26', 1)[0].wd, '木');
  assert.strictEqual(D.buildDays('2027-01-01', 1)[0].wd, '金');
});
t('日本時間で今日を判定（UTCではまだ前日の時刻）', () => {
  // 2026-11-26 00:30 JST = 2026-11-25 15:30 UTC
  assert.strictEqual(D.todayKey(new Date('2026-11-25T15:30:00Z')), '2026-11-26');
  // 2026-11-26 23:59 JST
  assert.strictEqual(D.todayKey(new Date('2026-11-26T14:59:00Z')), '2026-11-26');
  // 大晦日→元日
  assert.strictEqual(D.todayKey(new Date('2026-12-31T15:00:00Z')), '2027-01-01');
});
t('日付キーの検証', () => {
  assert.ok(D.isKey('2028-02-29'));
  assert.ok(!D.isKey('2027-02-29'));
  assert.ok(!D.isKey('2026-13-01'));
  assert.ok(!D.isKey('abc'));
});
t('addDays', () => {
  assert.strictEqual(D.addDays('2026-12-31', 1), '2027-01-01');
  assert.strictEqual(D.addDays('2026-03-01', -1), '2026-02-28');
});

console.log(`\n${n} tests passed`);
