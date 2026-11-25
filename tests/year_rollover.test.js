// index.html の「年またぎ」の月の扱いを、ページ本体から関数を取り出して確認する。
// 実行: node --test tests/year_rollover.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8').replace(/\r\n/g, '\n');

function extract(name) {
  const start = html.indexOf('    function ' + name + '(');
  assert.ok(start >= 0, name + ' not found');
  const end = html.indexOf('\n    }\n', start);
  return html.slice(start, end + 7);
}
const NAMES = ['scheduleMonthOffset', 'scheduleYearForMonth', 'daysInMonth', 'firstWeekday', 'currentScheduleMonth', 'isArchivedScheduleGroup',
  'getScheduleWindowMonths', 'scheduleGroupBelongsToWindow', 'getVisibleScheduleMonths', 'getStableScheduleMonth', 'getNearestVisibleScheduleMonth', 'getScheduleKey', 'getScheduleCategorySuffix', 'getEventId'];

// 擬似的な「今日」を差し替えて動かす
function env(today, groups, extra = {}) {
  const RealDate = Date;
  class FakeDate extends RealDate {
    constructor(...args) { if (args.length === 0) super(today); else super(...args); }
    static now() { return new RealDate(today).getTime(); }
  }
  const ctx = vm.createContext({ Date: FakeDate, Number, Object, Array, Math, Set, String, Boolean, Integer: undefined, console });
  vm.runInContext('var schedules = ' + JSON.stringify(groups) + '; var activeScheduleMonth = ' + (extra.active || 11) + ';'
    + 'function getMemberScheduleCategory(i){return i.category||"wonder"} function normalizeScheduleTime(m){return m||""} function normalizeScheduleTitleKey(t){return t}', ctx);
  vm.runInContext(NAMES.map(extract).join('\n'), ctx);
  return ctx;
}
const grp = (month, year, n = 3) => ({ month, items: Array.from({ length: n }, (_, i) => ({ day: i + 1, year, id: `${year}|${month}|${i + 1}|wonder|x|` })) });
const key = (m, c = 'Wonder') => `month${m}${c}`;
const DATA = { [key(9)]: grp(9, 2026), [key(10)]: grp(10, 2026), [key(11)]: grp(11, 2026), [key(12)]: grp(12, 2026), [key(1)]: grp(1, 2027) };
const run = (ctx, code) => JSON.parse(JSON.stringify(vm.runInContext(code, ctx)));

test('11月(今日=2026/11/26): 表示する月は 10,11,12,1(時系列)。1月は2027年で、過去扱いにならない', () => {
  const ctx = env('2026-11-26T10:00:00+09:00', DATA);
  assert.deepEqual(run(ctx, 'getVisibleScheduleMonths()'), [10, 11, 12, 1]);
  assert.equal(run(ctx, 'scheduleYearForMonth(1)'), 2027);
  assert.equal(run(ctx, 'scheduleYearForMonth(12)'), 2026);
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month1Wonder)'), false);
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month10Wonder)'), true);    // 前月は従来どおり「過去」(運営参加チェックの月タブには残る)
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month11Wonder)'), false);
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month9Wonder)'), true);
});

test('カレンダーの曜日と日数は、1月なら2027年で計算する(2027/1/1=金曜、31日)', () => {
  const ctx = env('2026-11-26T10:00:00+09:00', DATA);
  assert.equal(run(ctx, 'firstWeekday(1)'), 5);          // 月=1 … 金=5
  assert.equal(run(ctx, 'daysInMonth(1)'), 31);
  assert.equal(run(ctx, 'firstWeekday(12)'), 2);         // 2026/12/1=火曜
  assert.equal(run(ctx, 'daysInMonth(2)'), 28);          // データがない2月 → 2027年(平年)
});

test('12月に切り替わっても、前月11月〜2月まで続く。表示は時系列', () => {
  const data = { ...DATA, [key(2)]: grp(2, 2027) };
  const ctx = env('2026-12-10T10:00:00+09:00', data, { active: 12 });
  assert.deepEqual(run(ctx, 'getVisibleScheduleMonths()'), [11, 12, 1, 2]);
  assert.equal(run(ctx, 'getStableScheduleMonth(1)'), 1);
  assert.equal(run(ctx, 'getNearestVisibleScheduleMonth(1)'), 1);
});

test('年が明けた1月(今日=2027/1/5): 前月は2026年12月、1月は過去ではない。去年の古い1月データは出さない', () => {
  const data = { [key(12)]: grp(12, 2026), [key(1)]: grp(1, 2027), [key(2)]: grp(2, 2027), [key(7)]: grp(7, 2026) };
  const ctx = env('2027-01-05T10:00:00+09:00', data, { active: 1 });
  assert.deepEqual(run(ctx, 'getVisibleScheduleMonths()'), [12, 1, 2]);
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month12Wonder)'), true);   // 12月(2026)は過去
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month1Wonder)'), false);
  assert.equal(run(ctx, 'isArchivedScheduleGroup(schedules.month7Wonder)'), true);
  // 前の年のデータが残った月は、新しい年の範囲に混ぜない(例: 2026年の2月データが残っている場合)
  const stale = { ...data, [key(2)]: grp(2, 2026) };
  const ctx2 = env('2027-01-05T10:00:00+09:00', stale, { active: 1 });
  assert.deepEqual(run(ctx2, 'getVisibleScheduleMonths()'), [12, 1]);
});

test('範囲の外の月を指定したら、近い側に寄せる。予定IDには予定データの年が入る', () => {
  const ctx = env('2026-11-26T10:00:00+09:00', DATA);
  assert.equal(run(ctx, 'getStableScheduleMonth(3)'), 1);       // 先の月 → 最後
  assert.equal(run(ctx, 'getStableScheduleMonth(8)'), 10);      // 過去の月 → 最初
  assert.equal(run(ctx, 'getEventId({ month: 1 }, { day: 7, title: "x", meta: "", year: 2027 })'), '2027|1|7|wonder|x');
  assert.equal(run(ctx, 'getEventId({ month: 1 }, { day: 7, title: "x", meta: "" })'), '2027|1|7|wonder|x');   // 年が無いデータも、1月は2027年
});
