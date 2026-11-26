/*
 * 日付まわりの共通処理（日本時間 Asia/Tokyo 固定）
 * 日付は保存せず、毎回「今日」から計算する。月末・年末・閏年は Date.UTC に任せる。
 */
(function (root) {
  'use strict';

  var TZ = 'Asia/Tokyo';
  var DAY_MS = 86400000;
  var WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

  // 日本時間の今日を "YYYY-MM-DD" で返す（端末のタイムゾーンに左右されない）
  function todayKey(now) {
    var parts = new Intl.DateTimeFormat('en-US', {
      timeZone: TZ, year: 'numeric', month: '2-digit', day: '2-digit'
    }).formatToParts(now || new Date());
    var o = {};
    parts.forEach(function (p) { o[p.type] = p.value; });
    return o.year + '-' + o.month + '-' + o.day;
  }

  function isKey(k) {
    if (typeof k !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(k)) return false;
    var a = k.split('-').map(Number);
    var d = new Date(Date.UTC(a[0], a[1] - 1, a[2]));
    return d.getUTCFullYear() === a[0] && d.getUTCMonth() === a[1] - 1 && d.getUTCDate() === a[2];
  }

  function addDays(key, n) {
    var a = key.split('-').map(Number);
    return new Date(Date.UTC(a[0], a[1] - 1, a[2]) + n * DAY_MS).toISOString().slice(0, 10);
  }

  // startKey から count 日分の日付情報を作る
  function buildDays(startKey, count) {
    var a = startKey.split('-').map(Number);
    var base = Date.UTC(a[0], a[1] - 1, a[2]);
    var out = [];
    for (var i = 0; i < count; i++) {
      var d = new Date(base + i * DAY_MS);
      out.push({
        key: d.toISOString().slice(0, 10),
        y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
        w: d.getUTCDay(), wd: WEEKDAYS[d.getUTCDay()], i: i
      });
    }
    return out;
  }

  var api = { TZ: TZ, WEEKDAYS: WEEKDAYS, todayKey: todayKey, isKey: isKey, addDays: addDays, buildDays: buildDays };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.FujiwaraDates = api;
})(this);
