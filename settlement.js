/* 精算確認(運営参加チェックの「精算確認」タブ)
 * イベントの一覧にチェックを付けると「精算未完了(赤)」→「精算完了済み(緑)」になり、
 * 誰がいつチェックしたかを表示する。運営参加チェックとは別のデータ(サーバーのシート「精算確認」)。
 * チェックは押した時点で保存する(保存ボタンなし)。他の人の変更は1分ごと・タブを開くたびに取り込む。
 */
(function () {
  'use strict';
  var CATEGORY = 'settlement';
  var S = { done: {}, loaded: false, loading: false, error: '', pending: {}, errAt: 0, timer: 0 };

  function esc(v) { return typeof escapeHtml === 'function' ? escapeHtml(v) : String(v == null ? '' : v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function token() { try { return localStorage.getItem('wonderMemberToken') || ''; } catch (e) { return ''; } }
  function member() { try { return typeof currentMember !== 'undefined' ? currentMember : null; } catch (e) { return null; } }
  function active() { try { return activeMemberCategory === CATEGORY; } catch (e) { return false; } }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  function api(action, params, retried) {
    var t = token();
    if (!t) return Promise.reject(new Error('ログインしてください。'));
    return callPortalAction(action, Object.assign({ token: t }, params || {})).then(function (p) {
      if (!p || p.ok === false) throw new Error((p && (p.message || p.error)) || '処理に失敗しました。');
      return p;
    }).catch(function (err) {
      if (!retried && /ログイン|期限|認証/.test(String(err && err.message)) && typeof restoreTokenFromSavedLogin === 'function') {
        return restoreTokenFromSavedLogin({ silent: true }).then(function (nt) { if (!nt) throw err; return api(action, params, true); });
      }
      throw err;
    });
  }

  function stamp(iso) {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '';
    return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  function load(force) {
    if (!member() || S.loading) return Promise.resolve();
    if (!force && Date.now() - S.errAt < 30000) return Promise.resolve();
    S.loading = true;
    return api('settlementList').then(function (p) {
      var merged = Object.assign({}, p.done || {});
      Object.keys(S.pending).forEach(function (id) { if (S.pending[id] === 1) merged[id] = S.done[id]; else delete merged[id]; });   // 保存中の自分の操作は消さない
      S.done = merged; S.loaded = true; S.error = ''; S.errAt = 0;
    }, function (err) { S.error = (err && err.message) || '読み込めませんでした。'; S.errAt = Date.now(); })
      .then(function () { S.loading = false; if (active()) render(); });
  }

  function entries() {
    var month = typeof activeMemberMonth !== 'undefined' ? activeMemberMonth : 0;
    var list = getScheduleEntries({ includeArchived: true }).filter(function (e) {
      return e.group.month === month && getMemberScheduleCategory(e.item) !== 'meeting';
    });
    return list.map(function (e, i) { return { group: e.group, item: e.item, order: i }; })
      .sort(function (a, b) { return (Number(a.item.day) - Number(b.item.day)) || (a.order - b.order); });
  }

  function render() {
    var list = document.getElementById('memberScheduleList');
    if (!list || !active()) return;
    var savebar = document.querySelector('.member-savebar'); if (savebar) savebar.style.display = 'none';
    var note = document.getElementById('meetingNote'); if (note) note.classList.remove('is-visible');
    if (!member()) {
      list.innerHTML = '<div class="event-row"><div class="datebox">確認</div><div><div class="event-title">ポータルログイン後に精算確認を使えます</div></div></div>';
      return;
    }
    var rows = entries();
    var doneCount = rows.filter(function (e) { return S.done[getEventId(e.group, e.item)]; }).length;
    var head = '<div class="settle-summary"><span class="settle-open">精算未完了 ' + (rows.length - doneCount) + '件</span><span class="settle-done">完了済み ' + doneCount + '件</span>' +
      '<span class="settle-note">チェックすると精算完了済みになり、チェックした人の名前が表示されます(運営参加チェックとは別です)。</span></div>';
    if (S.error) head += '<div class="settle-error">最新の状態を読み込めませんでした(' + esc(S.error) + ')。前回の表示を使っています。</div>';
    list.innerHTML = head + (rows.length ? rows.map(function (e) {
      var id = getEventId(e.group, e.item), d = S.done[id], busy = S.pending[id] !== undefined;
      var cat = getMemberCategoryLabel(getMemberScheduleCategory(e.item));
      var meta = (formatScheduleDisplayMeta(e.item, e.group) || '詳細未入力') + ' ・ ' + cat;
      var state = d
        ? '<div class="settle-state is-done">精算完了済み<span class="settle-by">' + esc(d.name || '') + (d.at ? ' / ' + esc(stamp(d.at)) : '') + '</span></div>'
        : '<div class="settle-state is-open">精算未完了</div>';
      return '<label class="settle-row ' + (d ? 'is-done' : 'is-open') + '">' +
        '<input type="checkbox" data-settle-id="' + esc(id) + '" data-title="' + esc(e.item.title) + '" data-date="' + e.group.month + '/' + e.item.day + '"' + (d ? ' checked' : '') + (busy ? ' disabled' : '') + '>' +
        '<div class="datebox">' + e.group.month + '/' + e.item.day + '</div>' +
        '<div><div class="event-title">' + esc(e.item.title) + '</div><div class="event-meta">' + esc(meta) + '</div>' + state + '</div></label>';
    }).join('') : '<div class="event-row"><div class="datebox">未反映</div><div><div class="event-title">' + (typeof activeMemberMonth !== 'undefined' ? activeMemberMonth : '') + '月の予定はまだありません</div></div></div>');
  }

  function toggle(input) {
    var id = input.dataset.settleId, on = input.checked, me = member();
    if (!id || !me || S.pending[id] !== undefined) return;
    var before = S.done[id];
    S.pending[id] = on ? 1 : 0;
    if (on) S.done[id] = { name: me.name, email: me.email, at: new Date().toISOString() }; else delete S.done[id];
    render();
    api('settlementSet', { eventId: id, title: input.dataset.title || '', date: input.dataset.date || '', done: on ? '1' : '0' }).then(function (r) {
      if (on && r.entry) S.done[id] = r.entry;
    }, function (err) {
      if (before) S.done[id] = before; else delete S.done[id];
      alert((err && err.message) || '保存できませんでした。もう一度お試しください。');
    }).then(function () { delete S.pending[id]; render(); });
  }

  function boot() {
    var list = document.getElementById('memberScheduleList');
    if (!list) return;
    list.addEventListener('change', function (ev) { var t = ev.target; if (t && t.matches && t.matches('input[data-settle-id]')) toggle(t); });
    // 運営参加チェックの描画から呼ばれる(精算確認タブのときだけ)。他のタブへ戻ったら保存ボタン欄を元に戻す。
    window.renderSettlementList = function () { render(); load(false); };
    document.addEventListener('click', function (ev) {
      var b = ev.target.closest && ev.target.closest('.tab-button[data-member-category]');
      if (!b) return;
      var sv = document.querySelector('.member-savebar');
      if (sv) sv.style.display = b.dataset.memberCategory === CATEGORY ? 'none' : '';
    });
    S.timer = setInterval(function () { if (active() && !document.hidden && member()) load(true); }, 60000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden && active() && member()) load(true); });
    if (active()) { render(); load(true); }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
