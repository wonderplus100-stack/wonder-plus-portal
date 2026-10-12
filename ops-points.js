/*
 * 運営ポイント(スタッフ能力評価・運営可否判定) - front end.
 *
 * Loaded after the main portal script. It never changes how 運営参加チェック saves:
 *  - decorates the existing check rows with the verdict (read-only, from action=opsOverview)
 *  - adds an admin-only management card (role=admin; the server enforces this again)
 * Every number shown is calculated on the server; nothing here is typed in by hand.
 */
(function () {
  'use strict';
  if (window.__opsPointsLoaded) return;
  window.__opsPointsLoaded = true;

  var CACHE_KEY = 'wonderPlus.ops.overview.v1';
  var META = {
    abundant: { icon: '◎', cls: 'is-good' }, stable: { icon: '○', cls: 'is-good' }, ok: { icon: '○', cls: 'is-good' },
    adjust: { icon: '▲', cls: 'is-yellow' }, ng: { icon: '✕', cls: 'is-red' }, hold: { icon: '？', cls: 'is-blue' }
  };
  var STATUS_ORDER = ['ng', 'hold', 'adjust', 'ok', 'stable', 'abundant'];
  var ROLE_LABEL = { main: '主担当', support: '補助', trainee: '研修(人数外)', ng: '担当不可', unrated: '未評価' };
  var CAP_LABEL = { none: '制限なし', second_one: '2人目以降は1pt', second_zero: '2人目以降は0pt' };

  /* ------------------------------------------------------------------ helpers */
  function esc(v) {
    if (typeof escapeHtml === 'function') return escapeHtml(v);
    return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
  }
  function me() { try { return typeof currentMember !== 'undefined' ? currentMember : null; } catch (e) { return null; } }
  function isAdmin() { var m = me(); return !!m && m.role === 'admin'; }
  function token() { try { return localStorage.getItem('wonderMemberToken') || ''; } catch (e) { return ''; } }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function pct(n) { return n === null || n === undefined ? '—' : (Math.round(n * 10) / 10) + '%'; }
  function clone(v) { return JSON.parse(JSON.stringify(v)); }

  function api(action, params, retried) {
    var t = token();
    if (!t) return Promise.reject(new Error('ログインしてください。'));
    return callPortalAction(action, Object.assign({ token: t }, params || {})).then(function (p) {
      if (!p || p.ok === false) throw new Error((p && (p.message || p.error)) || '処理に失敗しました。');
      return p;
    }).catch(function (err) {
      var authProblem = /ログイン|期限|認証/.test(String(err && err.message));
      if (authProblem && !retried && typeof restoreTokenFromSavedLogin === 'function') {
        return restoreTokenFromSavedLogin({ silent: true }).then(function (nt) { if (!nt) throw err; return api(action, params, true); });
      }
      throw err;
    });
  }

  /* ------------------------------------------------------------------ member view */
  var S = { overview: null, loading: false, error: '', email: '' };

  function readCache() {
    try {
      var c = JSON.parse(localStorage.getItem(CACHE_KEY) || 'null');
      var m = me();
      if (c && m && c.email === String(m.email || '').toLowerCase()) S.overview = c.data;
    } catch (e) { /* ignore */ }
  }
  function writeCache(data) {
    try { var m = me(); localStorage.setItem(CACHE_KEY, JSON.stringify({ email: String(m && m.email || '').toLowerCase(), data: data })); } catch (e) { /* ignore */ }
  }

  function loadOverview() {
    if (!me() || !token()) { S.overview = null; decorate(); return; }
    if (S.loading) { S.again = true; return; }
    S.loading = true;
    api('opsOverview').then(function (p) {
      S.overview = p; S.error = ''; writeCache(p);
    }).catch(function (e) {
      S.error = String(e && e.message || e);
    }).then(function () {
      S.loading = false; decorate(); renderMe();
      if (S.again) { S.again = false; loadOverview(); }
    });
  }

  function badgeHtml(v) {
    var m = META[v.status] || META.hold;
    var short = v.staffShort > 0 ? ' <span class="ops-short">不足:' + v.staffShort + '名</span>' : '';
    var minText = v.min === null || v.min === undefined ? '未設定' : v.min + '名';
    var lines = '<span>参加 ' + v.count + '名／最低' + minText + short + '</span>';
    if (v.rate !== null && v.rate !== undefined) lines += '<span>達成率 ' + pct(v.rate) + '</span>';
    lines += '<span>判定：<span class="ops-chip ' + m.cls + '">' + m.icon + ' ' + esc(v.label) + '</span></span>';
    var reasons = (v.reasons || []).map(function (r) { return '<li>' + esc(r) + '</li>'; }).join('');
    var calc = (v.required !== null && v.required !== undefined)
      ? '<li>必要ポイント ' + v.required + 'pt / 実ポイント合計 ' + v.raw + 'pt / 制限適用後 ' + v.effective + 'pt' + (v.pointShort > 0 ? '(あと' + v.pointShort + 'pt)' : '') + '</li>' : '';
    var hist = v.fromHistory ? '<li>終了したイベントのため、判定時点の記録を表示しています。</li>' : '';
    return '<div class="ops-lines">' + lines + '</div>'
      + '<div class="ops-reason">' + esc((v.reasons || [])[0] || '') + '<button type="button" class="ops-more">詳細</button></div>'
      + '<div class="ops-detail" hidden><ul>' + calc + reasons + hist + '</ul></div>';
  }

  var observer = null;
  function decorate() {
    var list = document.getElementById('memberScheduleList');
    if (!list) return;
    if (observer) observer.disconnect();
    var events = (S.overview && S.overview.events) || {};
    list.querySelectorAll('.check-row').forEach(function (row) {
      var input = row.querySelector('input[data-event-id]');
      var body = row.children[2];
      if (!input || !body) return;
      var v = events[input.dataset.eventId];
      var existing = row.querySelector('.ops-badge');
      row.classList.remove('ops-row-yellow', 'ops-row-red', 'ops-row-blue');
      if (!v) { if (existing) existing.remove(); return; }
      var sig = JSON.stringify(v);
      if (v.color && v.color !== 'none') row.classList.add('ops-row-' + v.color);
      if (existing && existing.dataset.sig === sig) return;
      var open = existing && existing.querySelector('.ops-detail') && !existing.querySelector('.ops-detail').hidden;
      if (existing) existing.remove();
      var el = document.createElement('div');
      el.className = 'ops-badge'; el.dataset.sig = sig; el.innerHTML = badgeHtml(v);
      if (open) el.querySelector('.ops-detail').hidden = false;
      body.appendChild(el);
    });
    renderMe();
    if (observer && list) observer.observe(list, { childList: true });
  }

  function renderMe() {
    var list = document.getElementById('memberScheduleList');
    if (!list) return;
    var el = document.getElementById('opsMeLine');
    var o = S.overview;
    if (!me() || !o || !o.me) { if (el) el.remove(); return; }
    if (!el) { el = document.createElement('div'); el.id = 'opsMeLine'; el.className = 'ops-me'; list.parentNode.insertBefore(el, list); }
    var text;
    if (!o.me.linked) text = '運営ポイント: あなたのアカウントはまだスタッフ評価と紐付いていません(管理者が確認すると判定に反映されます)。';
    else if (o.me.provisional) text = '運営ポイント: ' + esc(o.me.name) + 'さんの評価は未評価' + o.me.unratedCount + '項目あり、暫定扱いです。';
    else text = '運営ポイント: ' + esc(o.me.name) + 'さんは <strong>' + esc(o.me.rank) + 'ランク / ' + o.me.points + 'pt</strong>(評価合計 ' + o.me.total + ')';
    if (S.error) text += ' ※判定の更新に失敗しました(' + esc(S.error) + ')';
    el.innerHTML = text;
  }

  document.addEventListener('click', function (e) {
    var btn = e.target.closest && e.target.closest('.ops-more');
    if (!btn) return;
    e.preventDefault(); e.stopPropagation();
    var d = btn.closest('.ops-badge').querySelector('.ops-detail');
    if (d) { d.hidden = !d.hidden; btn.textContent = d.hidden ? '詳細' : '閉じる'; }
  }, true);

  /* ------------------------------------------------------------------ admin */
  var A = { data: null, loading: false, tab: 'dash', msg: null, busy: {}, open: {}, detail: {}, f: { month: '', status: '', upcoming: true }, compare: null, audit: null, auto: {}, edit: null, root: null, ready: false };

  function setMsg(type, text) {
    A.msg = { type: type, text: text };
    if (type === 'ok') { var m = A.msg; setTimeout(function () { if (A.msg === m) { A.msg = null; render(); } }, 6000); }
  }

  /** Runs one request at a time per key: disables the button, shows progress, reports success/failure. */
  function run(key, label, fn, okText) {
    if (A.busy[key]) return Promise.resolve();
    A.busy[key] = label || true; setMsg('warn', (label || '処理中') + '…'); render();
    return fn().then(function (res) {
      delete A.busy[key];
      setMsg('ok', typeof okText === 'function' ? okText(res) : (okText || '保存しました。'));
      render(); return res;
    }).catch(function (err) {
      delete A.busy[key]; setMsg('error', String(err && err.message || err)); render();
    });
  }
  function busyAttr(key) { return A.busy[key] ? ' disabled' : ''; }
  function spin(key, text) { return (A.busy[key] ? '<span class="ops-spin"></span>' : '') + text; }

  function ensureRoot() {
    if (A.root && document.body.contains(A.root)) return A.root;
    var card = document.getElementById('adminCard');
    if (!card) return null;
    var root = document.createElement('div');
    root.id = 'opsAdminCard'; root.className = 'member-card member-session ops-admin';
    card.parentNode.insertBefore(root, card.nextSibling);
    root.addEventListener('click', onClick);
    root.addEventListener('change', onChange);
    A.root = root;
    return root;
  }

  function syncAdmin() {
    var root = ensureRoot();
    if (!root) return;
    if (!isAdmin()) { root.style.display = 'none'; return; }
    root.style.display = '';
    render();
  }

  function loadAdmin() {
    if (A.loading) return;
    A.loading = true; A.error = ''; render();
    api('opsAdminOverview').then(function (d) { A.data = d; A.error = ''; })
      .catch(function (e) { A.error = String(e && e.message || e); })
      .then(function () { A.loading = false; render(); });
  }

  var TABS = [['dash', '運営状況'], ['staff', 'スタッフ能力'], ['rules', 'イベント基準'], ['settings', '設定'], ['tools', 'ツール・履歴']];

  function render() {
    var root = A.root; if (!root || !isAdmin()) return;
    var head = '<div class="member-actions"><div><strong>運営ポイント管理</strong><div class="ops-sub">スタッフの能力評価と、イベントごとの運営可否を確認・調整します(管理者のみ)。</div></div>'
      + '<button class="member-button" type="button" data-act="reload"' + (A.loading ? ' disabled' : '') + '>' + (A.loading ? '<span class="ops-spin"></span>読み込み中' : (A.data ? '最新に更新' : '開いて読み込む')) + '</button></div>';
    var msg = A.msg ? '<div class="ops-note is-' + (A.msg.type === 'warn' ? 'warn' : A.msg.type) + '" role="status">' + esc(A.msg.text) + '</div>' : '';
    if (!A.data) {
      root.innerHTML = head + msg + (A.error ? '<div class="ops-note is-error">' + esc(A.error) + '</div>' : '<div class="ops-sub">「開いて読み込む」を押すと表示します(数秒かかります)。</div>');
      return;
    }
    var d = A.data, banners = '';
    if (!d.sourceOk) banners += '<div class="ops-note is-error">運営ポイントのExcel(スプレッドシート)を正しく読めていません: ' + esc((d.sourceErrors || []).join(' / ')) + '</div>';
    if (!d.sheetsReady) {
      banners += '<div class="ops-note is-warn">保存用のシートがまだありません。管理シートに「運営_」で始まる6つのシートを<strong>新規追加</strong>します(既存のシート・データは変更しません)。 <button class="member-button" type="button" data-act="setup"' + busyAttr('setup') + '>' + spin('setup', '保存用シートを準備する') + '</button></div>';
    }
    if (d.settingsWarnings && d.settingsWarnings.length) banners += '<div class="ops-note is-warn">保存済みの設定に不正な値があったため初期値を使っています: ' + esc(d.settingsWarnings.join(' / ')) + '</div>';
    var pending = (d.staff || []).filter(function (s) { return ['confirmed', 'auto'].indexOf(s.mapping.status) < 0; }).length;
    if (pending) banners += '<div class="ops-note is-warn">アカウントとの対応付けが未確認のスタッフが ' + pending + '名 います(「スタッフ能力」で確認)。確認が済むまで、そのスタッフが参加するイベントは判定保留になります。</div>';
    var tabs = '<div class="ops-tabs">' + TABS.map(function (t) { return '<button type="button" class="tab-button ' + (A.tab === t[0] ? 'is-active' : '') + '" data-act="tab" data-tab="' + t[0] + '">' + t[1] + '</button>'; }).join('') + '</div>';
    var body = { dash: renderDash, staff: renderStaff, rules: renderRules, settings: renderSettings, tools: renderTools }[A.tab](d);
    root.innerHTML = head + msg + banners + tabs + body;
  }

  /* ---- 運営状況 ---- */
  function monthOf(e) { return Math.floor((e.ymd % 10000) / 100); }
  function todayYmd() { var n = new Date(); return n.getFullYear() * 10000 + (n.getMonth() + 1) * 100 + n.getDate(); }

  function chip(status, label) {
    var m = META[status] || META.hold;
    return '<span class="ops-chip ' + m.cls + '">' + m.icon + ' ' + esc(label) + '</span>';
  }

  function renderDash(d) {
    var events = d.events.filter(function (e) {
      if (A.f.upcoming && e.ymd < todayYmd()) return false;
      if (A.f.month && String(monthOf(e)) !== A.f.month) return false;
      if (A.f.status === 'todo') return ['adjust', 'ng', 'hold'].indexOf(e.status) >= 0;
      if (A.f.status && A.f.status !== 'todo') return e.status === A.f.status;
      return true;
    }).sort(function (a, b) { return a.ymd - b.ymd; });
    var counts = {};
    d.events.filter(function (e) { return !A.f.upcoming || e.ymd >= todayYmd(); }).forEach(function (e) { counts[e.status] = (counts[e.status] || 0) + 1; });
    var months = {};
    d.events.forEach(function (e) { months[monthOf(e)] = true; });
    var summary = '<div class="ops-summary">' + STATUS_ORDER.map(function (s) { return counts[s] ? chip(s, d.verdicts[s].label + ' ' + counts[s]) : ''; }).join('') + '</div>';
    var tools = '<div class="ops-toolbar">'
      + '<select data-f="month" aria-label="月"><option value="">すべての月</option>' + Object.keys(months).sort(function (a, b) { return a - b; }).map(function (m) { return '<option value="' + m + '"' + (A.f.month === m ? ' selected' : '') + '>' + m + '月</option>'; }).join('') + '</select>'
      + '<select data-f="status" aria-label="判定"><option value="">すべての判定</option><option value="todo"' + (A.f.status === 'todo' ? ' selected' : '') + '>要対応(要調整・運営不可・保留)</option>'
      + STATUS_ORDER.map(function (s) { return '<option value="' + s + '"' + (A.f.status === s ? ' selected' : '') + '>' + esc(d.verdicts[s].label) + '</option>'; }).join('') + '</select>'
      + '<label class="ops-sub"><input type="checkbox" data-f="upcoming"' + (A.f.upcoming ? ' checked' : '') + '> 今後のみ</label></div>';
    var list = events.length ? events.map(function (e) { return eventHtml(e, d); }).join('') : '<div class="ops-note">該当するイベントはありません。</div>';
    var unmatched = (d.unmatched || []).length
      ? '<div class="ops-note is-warn">種別を判定できないイベントが ' + d.unmatched.length + '件 あります。「イベント基準」で種別を設定すると判定されます(自動では他のイベントに紐付けていません)。</div>' : '';
    return '<h4>イベントごとの運営状況</h4>' + summary + unmatched + tools + '<div class="ops-event-list">' + list + '</div>';
  }

  function eventHtml(e, d) {
    var color = META[e.status] && e.status === 'adjust' ? 'yellow' : e.status === 'ng' ? 'red' : e.status === 'hold' ? 'blue' : '';
    var open = !!A.open[e.eventId];
    var short = e.staffShort > 0 ? ' ・ <span class="ops-short">不足:' + e.staffShort + '名</span>' : '';
    var meta = '参加 ' + e.count + '名／最低' + (e.min === null ? '未設定' : e.min + '名') + short + (e.rate !== null ? ' ・ 達成率 ' + pct(e.rate) : '') + (e.typeName ? ' ・ 基準: ' + esc(e.typeName) : ' ・ 基準: 未判定');
    return '<div class="ops-event ' + (color ? 'is-' + color : '') + '">'
      + '<button type="button" class="ops-event-head" data-act="event" data-id="' + esc(e.eventId) + '" aria-expanded="' + open + '">'
      + '<span class="ops-date">' + esc(e.date) + '</span><span><strong>' + esc(e.title) + '</strong> ' + chip(e.status, e.label) + '</span>'
      + '<span class="ops-meta">' + meta + '</span>'
      + '<span class="ops-meta">' + esc((e.reasons || [])[0] || '') + '</span></button>'
      + (open ? '<div class="ops-event-body">' + detailHtml(e.eventId) + '</div>' : '') + '</div>';
  }

  function relClock(timeRange, rel) {
    var m = /(\d{1,2}):(\d{2})/.exec(timeRange || '');
    if (!m || rel === null || rel === undefined) return '';
    var t = Number(m[1]) * 60 + Number(m[2]) + rel;
    t = ((t % 1440) + 1440) % 1440;
    return pad(Math.floor(t / 60)) + ':' + pad(t % 60);
  }

  function detailHtml(id) {
    var dt = A.detail[id];
    if (!dt) return '<div class="ops-sub"><span class="ops-spin"></span>読み込み中…</div>';
    if (dt.error) return '<div class="ops-note is-error">' + esc(dt.error) + '</div>';
    var j = dt.judgement, tm = {};
    (dt.tasks || []).forEach(function (t) { tm[t.taskId] = t; });
    var out = '<div class="ops-grid cols-2"><div>';
    out += '<h4>判定の理由</h4><ul>' + j.reasons.map(function (r) { return '<li>' + esc(r) + '</li>'; }).join('') + '</ul>';
    out += '<h4>ポイント計算(' + esc(CAP_LABEL[j.points.capMode] || '') + ')</h4>';
    if (j.points.lines.length) {
      out += '<div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>参加者</th><th class="num">ランク</th><th class="num">実pt</th><th class="num">計上pt</th><th>備考</th></tr></thead><tbody>'
        + j.points.lines.map(function (l) { return '<tr><td>' + esc(l.name) + '</td><td class="num">' + esc(l.rank) + '</td><td class="num">' + l.points + '</td><td class="num">' + l.counted + '</td><td>' + esc(l.note) + '</td></tr>'; }).join('')
        + '<tr><td><strong>合計</strong></td><td></td><td class="num"><strong>' + j.points.raw + '</strong></td><td class="num"><strong>' + j.points.effective + '</strong></td><td>必要 ' + (j.points.required === null ? '未設定' : j.points.required + 'pt') + ' / 達成率 ' + pct(j.points.rate) + '</td></tr></tbody></table></div>';
    } else out += '<div class="ops-sub">評価が確定している参加者がいません。</div>';
    if (j.unconfirmed && j.unconfirmed.length) out += '<div class="ops-note is-warn">評価が確定していない参加者: ' + j.unconfirmed.map(function (u) { return esc(u.name) + '(' + esc(u.reason) + ')'; }).join('、') + '</div>';
    out += '</div><div>';
    out += '<h4>人数</h4><div>参加 ' + j.headcount.count + '名 / 最低 ' + (j.headcount.min === null ? '未設定' : j.headcount.min + '名') + (j.headcount.shortfall > 0 ? ' / <span class="ops-short">不足 ' + j.headcount.shortfall + '名</span>' : '') + '</div>';
    out += '<h4>業務ごとの充足</h4>';
    if (!j.tasks.length) out += '<div class="ops-sub">必須業務の設定がありません。</div>';
    else {
      var savedKey = {};
      (dt.manual || []).forEach(function (m) { savedKey[m.taskId + '|' + m.staffKey] = true; });
      out += '<div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>業務(時間帯)</th><th class="num">必要</th><th class="num">割当</th><th class="num">不足</th><th class="num">適格</th><th>担当</th></tr></thead><tbody>'
        + j.tasks.map(function (t) {
          var meta = tm[t.taskId] || {};
          var win = meta.window ? relClock(j.event.timeRange, meta.window[0]) + '〜' + relClock(j.event.timeRange, meta.window[1]) : '';
          var pills = t.assigned.map(function (a) {
            var saved = savedKey[t.taskId + '|' + a.staffKey];
            var cls = a.role === 'main' || a.role === 'support' ? '' : (a.role === 'trainee' ? ' is-warn' : ' is-bad');
            return '<span class="ops-pill' + cls + '">' + esc(a.name) + ' <small>' + esc(ROLE_LABEL[a.role] || '') + '</small>'
              + (saved ? '<button type="button" title="外す" data-act="unassign" data-event="' + esc(id) + '" data-task="' + t.taskId + '" data-staff="' + esc(a.staffKey) + '"' + busyAttr('as:' + id) + '>×</button>' : '<small>(案)</small>') + '</span>';
          }).join(' ');
          var options = (dt.people || []).filter(function (p) {
            return p.staffKey && p.complete && !t.assigned.some(function (a) { return a.staffKey === p.staffKey; });
          }).map(function (p) {
            var r = p.ratings ? p.ratings[meta.ability] : null;
            return '<option value="' + esc(p.staffKey) + '"' + (r === 5 || r === null ? ' disabled' : '') + '>' + esc(p.name) + '(評価' + (r === null ? '—' : r) + ')</option>';
          }).join('');
          var add = options ? '<span class="ops-assign"><select data-pick="' + t.taskId + '" aria-label="' + esc(t.name) + 'の担当を選ぶ"><option value="">追加…</option>' + options + '</select><button class="member-button" type="button" data-act="assign" data-event="' + esc(id) + '" data-task="' + t.taskId + '"' + busyAttr('as:' + id) + '>割当</button></span>' : '';
          return '<tr><td>' + esc(t.name) + '<div class="ops-sub">' + esc(win) + '</div></td><td class="num">' + t.required + '</td><td class="num">' + t.counted + '</td><td class="num">' + (t.shortfall ? '<span class="ops-short">' + t.shortfall + '</span>' : '0') + '</td><td class="num">' + t.eligibleCount + '</td><td>' + pills + ' ' + add + '</td></tr>';
        }).join('') + '</tbody></table></div>';
      if (j.conflicts && j.conflicts.length) out += '<div class="ops-note is-error">同じ時間帯の重複割当があります。担当を見直してください。</div>';
      if (j.assignmentSource === 'auto') {
        out += '<div class="ops-note is-warn">業務割当が未登録のため、<strong>自動割当案</strong>で判定しています(確定ではありません)。 <button class="member-button" type="button" data-act="apply-auto" data-event="' + esc(id) + '"' + busyAttr('as:' + id) + '>' + spin('as:' + id, 'この案を確定する') + '</button></div>';
      } else if (j.tasks.some(function (t) { return t.shortfall > 0; })) {
        out += '<div class="ops-toolbar"><button class="member-button" type="button" data-act="auto" data-event="' + esc(id) + '"' + busyAttr('auto:' + id) + '>' + spin('auto:' + id, '不足分の割当案を出す') + '</button></div>';
        var sug = A.auto[id];
        if (sug) out += '<div class="ops-note">' + (sug.proposed.length ? '案: ' + sug.proposed.map(function (p) { var tk = tm[p.taskId] || {}; return esc((tk.name || p.taskId) + '=' + (sug.names[p.staffKey] || p.staffKey)); }).join('、') + (sug.complete ? '' : '(すべては埋まりません)') + ' <button class="member-button" type="button" data-act="apply-sug" data-event="' + esc(id) + '"' + busyAttr('as:' + id) + '>この案を確定する</button>' : '適格者と時間の条件を満たす追加の割当案はありません。') + '</div>';
      }
    }
    out += '</div></div>';
    return out;
  }

  /* ---- スタッフ能力 ---- */
  function renderStaff(d) {
    var skills = d.skills;
    var head = '<tr><th class="ops-sticky">名前</th>' + skills.map(function (s) { return '<th class="num">' + esc(s.name) + '</th>'; }).join('')
      + '<th class="num">合計</th><th class="num">ランク</th><th class="num">pt</th><th class="num">未評価</th><th>最終更新</th><th>評価者</th><th>ポータルアカウント</th></tr>';
    var users = d.users || [];
    var rows = d.staff.map(function (s) {
      var rated = skills.map(function (k) { var v = s.ratings[k.id]; return '<td class="num ops-r' + (v === null ? 0 : v) + '">' + (v === null ? '—' : v) + '</td>'; }).join('');
      var rank = s.complete ? esc(s.rank) : '<span title="未評価があるため確定していません">暫定' + (s.provisionalRange ? esc(s.provisionalRange.best === s.provisionalRange.worst ? s.provisionalRange.best : s.provisionalRange.best + '〜' + s.provisionalRange.worst) : '') + '</span>';
      var map = s.mapping, acct;
      if (map.status === 'confirmed' || map.status === 'auto') {
        acct = '<span class="ops-pill">✓ ' + esc(map.email) + (map.status === 'auto' ? ' <small>名前が一致</small>' : '') + '<button type="button" title="対応付けを解除" data-act="unlink" data-staff="' + esc(s.staffKey) + '"' + busyAttr('map') + '>×</button></span>';
      } else {
        var sug = s.suggestions.map(function (u) { return '<option value="' + esc(u.email) + '">候補: ' + esc(u.name) + ' (' + esc(u.email) + ')</option>'; }).join('');
        var others = users.filter(function (u) { return !s.suggestions.some(function (x) { return x.email === u.email; }); }).map(function (u) { return '<option value="' + esc(u.email) + '">' + esc(u.name) + ' (' + esc(u.email) + ')</option>'; }).join('');
        acct = '<span class="ops-assign"><select data-map="' + esc(s.staffKey) + '" aria-label="' + esc(s.name) + 'のアカウント"><option value="">未確認(選択)</option>' + sug + (others ? '<optgroup label="その他のアカウント">' + others + '</optgroup>' : '') + '</select>'
          + '<button class="member-button" type="button" data-act="confirm-map" data-staff="' + esc(s.staffKey) + '"' + busyAttr('map') + '>確定</button></span>';
      }
      return '<tr><td class="ops-sticky"><strong>' + esc(s.name) + '</strong></td>' + rated + '<td class="num">' + s.total + (s.complete ? '' : '*') + '</td><td class="num">' + rank + '</td><td class="num">' + (s.points === null ? '—' : s.points) + '</td><td class="num">' + (s.unratedCount || '') + '</td><td>' + esc(s.updatedAt ? new Date(s.updatedAt).toLocaleDateString('ja-JP') : '未記録') + '</td><td>' + esc(s.evaluator || '未記録') + '</td><td>' + acct + '</td></tr>';
    }).join('');
    var single = d.staff.filter(function (s) { return ['confirmed', 'auto'].indexOf(s.mapping.status) < 0 && s.suggestions.length === 1; });
    return '<h4>スタッフ能力(Excelの値を表示)</h4>'
      + '<div class="ops-sub">評価はExcel「運営ポイント」シートで編集します。編集するとここにも反映されます(最大5分)。合計・ランク・ポイントはポータル側で自動計算し、手入力値は使いません。* は未評価あり(暫定)。</div>'
      + '<div class="ops-toolbar"><a class="member-button" href="' + esc(d.sourceUrl) + '" target="_blank" rel="noopener">Excelを開いて編集</a>'
      + (single.length ? '<button class="member-button" type="button" data-act="confirm-all"' + busyAttr('map') + '>' + spin('map', '候補が1件のスタッフ(' + single.length + '名)を一括確定') + '</button>' : '') + '</div>'
      + '<div class="ops-table-wrap"><table class="ops-table"><thead>' + head + '</thead><tbody>' + rows + '</tbody></table></div>'
      + '<div class="ops-sub">氏名だけでは同一人物と決められないため、アカウントとの対応付けは管理者の確定が必要です。</div>';
  }

  /* ---- イベント基準 ---- */
  function renderRules(d) {
    var tasks = d.tasks, ov = {};
    (d.overrides || []).forEach(function (o) { if (o.scope === 'type' && !o.disabled) ov[o.key] = o; });
    var rows = d.eventRules.map(function (r) {
      var eff = ov[r.typeId] || {};
      var req = eff.requiredPoints !== undefined ? eff.requiredPoints : r.requiredPoints, min = eff.minStaff !== undefined ? eff.minStaff : r.minStaff;
      var counts = Object.assign({}, r.taskCounts, eff.taskCounts || {});
      var summary = tasks.filter(function (t) { return counts[t.id] > 0; }).map(function (t) { return esc(t.name) + counts[t.id]; }).join('・');
      var unset = req === null || min === null;
      return '<tr><td><strong>' + esc(r.name) + '</strong>' + (eff.requiredPoints !== undefined || eff.minStaff !== undefined || eff.taskCounts ? ' <span class="ops-pill is-warn">例外あり</span>' : '') + '</td><td class="num">' + (unset ? '<span class="ops-chip is-blue">基準未設定</span>' : req) + '</td><td class="num">' + (min === null ? '—' : min) + '</td><td>' + (summary || '—') + '</td><td><button class="member-button" type="button" data-act="edit-rule" data-type="' + esc(r.typeId) + '">' + (unset ? '基準を設定' : '例外を設定') + '</button></td></tr>';
    }).join('');
    var editor = A.edit ? ruleEditor(d) : '';
    var unmatched = (d.unmatched || []).map(function (u) {
      var opts = d.eventRules.filter(function (r) { return r.configured; }).map(function (r) { return '<option value="' + esc(r.typeId) + '">' + esc(r.name) + '</option>'; }).join('');
      var cand = u.status === 'ambiguous' ? '(複数に該当: ' + esc(u.candidates.join('・')) + ')' : '';
      return '<tr><td>' + esc(u.date) + '</td><td>' + esc(u.title) + ' ' + cand + '</td><td><span class="ops-assign"><select data-alias="' + esc(u.eventId) + '" aria-label="種別"><option value="">種別を選ぶ</option>' + opts + '</select><button class="member-button" type="button" data-act="alias" data-event="' + esc(u.eventId) + '" data-title="' + esc(u.title) + '"' + busyAttr('alias') + '>この名前を紐付け</button></span></td></tr>';
    }).join('');
    var custom = (d.aliasCustom || []).map(function (r) {
      return '<tr><td>' + esc(d.typeNames[r.eventTypeId] || r.eventTypeId) + '</td><td class="ops-code">' + esc(r.pattern) + '</td><td>' + (r.disabled ? '無効' : '有効') + '</td><td><button class="member-button" type="button" data-act="alias-toggle" data-id="' + esc(r.id) + '"' + busyAttr('alias') + '>' + (r.disabled ? '有効にする' : '無効にする') + '</button></td></tr>';
    }).join('');
    return '<h4>イベント種別ごとの基準(Excel「イベント別」)</h4>'
      + '<div class="ops-sub">基準の元データはExcelです。ここでの「例外」はExcelを書き換えず、ポータル側にだけ上書きとして保存します(履歴あり)。Excelに無い地域・イベントは推測せず「基準未設定」です。</div>'
      + '<div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>イベント</th><th class="num">必要pt</th><th class="num">最低人数</th><th>業務必要人数</th><th></th></tr></thead><tbody>' + rows + '</tbody></table></div>'
      + editor
      + '<h4>種別を判定できないイベント</h4>'
      + (unmatched ? '<div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>日付</th><th>イベント名</th><th>種別の紐付け</th></tr></thead><tbody>' + unmatched + '</tbody></table></div>' : '<div class="ops-sub">今後のイベントはすべて種別が判定できています(基準が未設定のものは上の表を参照)。</div>')
      + (custom ? '<h4>追加した別名ルール</h4><div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>種別</th><th>イベント名の条件</th><th>状態</th><th></th></tr></thead><tbody>' + custom + '</tbody></table></div>' : '');
  }

  function ruleEditor(d) {
    var r = d.eventRules.filter(function (x) { return x.typeId === A.edit; })[0];
    if (!r) return '';
    var ov = (d.overrides || []).filter(function (o) { return o.scope === 'type' && o.key === r.typeId && !o.disabled; })[0] || {};
    var v = function (k, base) { return ov[k] !== undefined ? ov[k] : (base === null || base === undefined ? '' : base); };
    var counts = Object.assign({}, r.taskCounts, ov.taskCounts || {});
    return '<div class="ops-note"><div class="ops-form" data-form="rule" data-type="' + esc(r.typeId) + '"><strong>' + esc(r.name) + ' の基準</strong>'
      + '<div class="ops-inline"><label>必要ポイント<input type="number" min="0" step="1" name="requiredPoints" value="' + esc(v('requiredPoints', r.requiredPoints)) + '"></label>'
      + '<label>最低人数<input type="number" min="0" step="1" name="minStaff" value="' + esc(v('minStaff', r.minStaff)) + '"></label></div>'
      + '<div class="ops-inline">' + d.tasks.map(function (t) { return '<label>' + esc(t.name) + '<input type="number" min="0" step="1" name="task:' + t.id + '" value="' + (counts[t.id] === undefined ? 0 : counts[t.id]) + '"></label>'; }).join('') + '</div>'
      + '<label>理由メモ(履歴に残ります)<input type="text" name="note" value="' + esc(ov.note || '') + '"></label>'
      + '<div class="ops-toolbar"><button class="member-button primary" type="button" data-act="save-rule"' + busyAttr('rule') + '>' + spin('rule', '保存') + '</button><button class="member-button" type="button" data-act="cancel-rule">閉じる</button></div></div></div>';
  }

  /* ---- 設定 ---- */
  function renderSettings(d) {
    var s = d.settings;
    var sel = function (name, value, opts) { return '<select name="' + name + '">' + opts.map(function (o) { return '<option value="' + o[0] + '"' + (String(value) === String(o[0]) ? ' selected' : '') + '>' + esc(o[1]) + '</option>'; }).join('') + '</select>'; };
    var num = function (name, value, min, max) { return '<input type="number" name="' + name + '" value="' + esc(value) + '" min="' + (min === undefined ? 0 : min) + '"' + (max ? ' max="' + max + '"' : '') + ' step="1">'; };
    var ranks = s.rankRules.map(function (r, i) {
      return '<tr><td><input type="number" name="rank:' + i + ':min" value="' + r.min + '" style="width:64px"></td><td><input type="number" name="rank:' + i + ':max" value="' + r.max + '" style="width:64px"></td><td><input type="text" name="rank:' + i + ':rank" value="' + esc(r.rank) + '" style="width:56px"></td><td><input type="number" name="rank:' + i + ':points" value="' + r.points + '" style="width:64px"></td></tr>';
    }).join('');
    var anchors = [['open', '開場'], ['start', '開始'], ['end', '終了']];
    var wins = d.tasks.map(function (t) {
      var w = s.taskWindows[t.id] || { from: { anchor: 'start', offset: 0 }, to: { anchor: 'end', offset: 0 } };
      return '<tr><td>' + esc(t.name) + '</td><td>' + sel('win:' + t.id + ':fa', w.from.anchor, anchors) + ' ' + num('win:' + t.id + ':fo', w.from.offset, -1000) + '分</td><td>' + sel('win:' + t.id + ':ta', w.to.anchor, anchors) + ' ' + num('win:' + t.id + ':to', w.to.offset, -1000) + '分</td></tr>';
    }).join('');
    var pairs = (s.concurrentPairs || []).map(function (p, i) {
      var nm = function (id) { var t = d.tasks.filter(function (x) { return x.id === id; })[0]; return t ? t.name : id; };
      return '<span class="ops-pill">' + esc(nm(p[0])) + '＋' + esc(nm(p[1])) + '<button type="button" title="削除" data-act="pair-del" data-i="' + i + '">×</button></span>';
    }).join(' ') || '<span class="ops-sub">なし(同じ時間帯の業務は別スタッフが必要)</span>';
    var taskOpts = d.tasks.map(function (t) { return '<option value="' + t.id + '">' + esc(t.name) + '</option>'; }).join('');
    var abil = ['miraiba_guide', 'miraiba_ops'].map(function (id) {
      var t = d.tasks.filter(function (x) { return x.id === id; })[0];
      var cur = s.abilityMap[id] || (t && t.ability);
      return '<label>' + esc(t.name) + 'が参照する能力' + sel('ab:' + id, cur, d.skills.map(function (k) { return [k.id, k.name]; })) + '</label>';
    }).join('');
    return '<h4>運営判定の設定</h4><div class="ops-sub">仕様が確定していない項目は仮の初期値で動いています。変更すると全イベントの判定が再計算されます(変更履歴は「ツール・履歴」)。</div>'
      + '<div class="ops-form" data-form="settings">'
      + '<h4>ランクとポイント(合計が小さいほど高評価)</h4><div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>合計 下限</th><th>上限</th><th>ランク</th><th>運営pt</th></tr></thead><tbody>' + ranks + '</tbody></table></div>'
      + '<div class="ops-sub">既定は S 12〜15=5pt / A 16〜19=4pt / B 20〜24=3pt / C 25〜29=2pt / D 30〜60=1pt。Excelの凡例は「30〜36=D」ですが、5段階評価との整合のため37〜60もDとしています。</div>'
      + '<h4>人数・ポイントの数え方</h4><div class="ops-inline">'
      + '<label>5ptスタッフの制限(暫定仕様)' + sel('fivePointCap', s.fivePointCap, [['second_one', '2人目以降は1ptで計算'], ['none', '制限なし'], ['second_zero', '2人目以降は0ptで計算']]) + '</label>'
      + '<label>評価未完了の参加者がいる場合' + sel('unratedPolicy', s.unratedPolicy, [['hold', 'イベントを判定保留にする'], ['exclude', 'その人を除いて判定する']]) + '</label>'
      + '<label>業務割当が未登録のとき' + sel('assignmentMode', s.assignmentMode, [['auto_fallback', '自動割当案で判定する'], ['manual', '未割当として扱う(厳格)']]) + '</label></div>'
      + '<h4>判定のしきい値(達成率 %)</h4><div class="ops-inline"><label>余裕あり(以上)' + num('th:abundant', s.thresholds.abundant, 1) + '</label><label>安定運営(以上)' + num('th:stable', s.thresholds.stable, 1) + '</label><label>運営可能(以上)' + num('th:ok', s.thresholds.ok, 1) + '</label></div>'
      + '<h4>業務の適格基準(評価の値)</h4><div class="ops-inline"><label>主担当 〜' + num('el:mainMax', s.eligibility.mainMax, 1, 5) + '</label><label>補助 〜' + num('el:supportMax', s.eligibility.supportMax, 1, 5) + '</label><label>研修扱い 〜' + num('el:trainingMax', s.eligibility.trainingMax, 1, 5) + '</label><label>主担当が必要な人数' + num('mainSlots', s.mainSlots, 0, 10) + '</label></div>'
      + '<div class="ops-sub">既定: 評価1〜2=主担当可 / 3=補助のみ / 4=研修扱い(人数に数えない) / 5=担当不可。</div>'
      + '<h4>ミライバ業務と能力の対応</h4><div class="ops-inline">' + abil + '</div><div class="ops-sub">初期状態では「ミライバ案内」「ミライバ運営」とも「ミライバ」能力を参照します。同時に充足できるとは見なしません。</div>'
      + '<h4>業務の時間帯(仮設定)</h4><div class="ops-sub">同じスタッフを同じ時間帯の複数業務に割り当てられないための基準です。開場＝開始の' + s.openLeadMinutes + '分前。</div>'
      + '<div class="ops-table-wrap"><table class="ops-table"><thead><tr><th>業務</th><th>開始</th><th>終了</th></tr></thead><tbody>' + wins + '</tbody></table></div>'
      + '<div class="ops-inline"><label>開場は開始の何分前' + num('openLead', s.openLeadMinutes, 0, 600) + '</label><label>時間が不明なときの所要(分)' + num('defaultDuration', s.defaultDurationMinutes, 1, 600) + '</label></div>'
      + '<h4>同時兼務を許可する業務の組み合わせ</h4><div>' + pairs + '</div>'
      + '<div class="ops-inline"><label>業務A<select name="pairA">' + taskOpts + '</select></label><label>業務B<select name="pairB">' + taskOpts + '</select></label><button class="member-button" type="button" data-act="pair-add">組み合わせを追加</button></div>'
      + '<h4>その他</h4><div class="ops-inline"><label>開催日になっても必須条件が未達なら' + sel('closedAsNg', s.closedAsNg ? '1' : '0', [['1', '運営不可にする'], ['0', '要調整のまま']]) + '</label><label>判定を始める時期(開催の何日前から。0=常に)' + num('judgeFrom', s.judgeFromDaysBefore, 0, 1000) + '</label></div>'
      + '<div class="ops-toolbar"><button class="member-button primary" type="button" data-act="save-settings"' + busyAttr('settings') + '>' + spin('settings', '設定を保存') + '</button><button class="member-button" type="button" data-act="reset-settings"' + busyAttr('settings') + '>初期値に戻す</button></div></div>';
  }

  function readSettingsForm() {
    var root = A.root.querySelector('[data-form="settings"]');
    var val = function (n) { var el = root.querySelector('[name="' + n + '"]'); return el ? el.value : ''; };
    var d = A.data, s = {};
    s.rankRules = d.settings.rankRules.map(function (r, i) { return { min: Number(val('rank:' + i + ':min')), max: Number(val('rank:' + i + ':max')), rank: val('rank:' + i + ':rank').trim(), points: Number(val('rank:' + i + ':points')) }; });
    s.fivePointCap = val('fivePointCap'); s.unratedPolicy = val('unratedPolicy'); s.assignmentMode = val('assignmentMode');
    s.thresholds = { abundant: Number(val('th:abundant')), stable: Number(val('th:stable')), ok: Number(val('th:ok')) };
    s.eligibility = { mainMax: Number(val('el:mainMax')), supportMax: Number(val('el:supportMax')), trainingMax: Number(val('el:trainingMax')) };
    s.mainSlots = Number(val('mainSlots'));
    s.abilityMap = {}; ['miraiba_guide', 'miraiba_ops'].forEach(function (id) { s.abilityMap[id] = val('ab:' + id); });
    s.taskWindows = {};
    d.tasks.forEach(function (t) { s.taskWindows[t.id] = { from: { anchor: val('win:' + t.id + ':fa'), offset: Number(val('win:' + t.id + ':fo')) }, to: { anchor: val('win:' + t.id + ':ta'), offset: Number(val('win:' + t.id + ':to')) } }; });
    s.openLeadMinutes = Number(val('openLead')); s.defaultDurationMinutes = Number(val('defaultDuration'));
    s.concurrentPairs = (d.settings.concurrentPairs || []).slice();
    s.closedAsNg = val('closedAsNg') === '1'; s.judgeFromDaysBefore = Number(val('judgeFrom'));
    return s;
  }

  /* ---- ツール・履歴 ---- */
  function renderTools(d) {
    var cmp = A.compare ? '<div class="ops-note ' + (A.compare.differences.length || A.compare.taskDifferences.length ? 'is-warn' : 'is-ok') + '">スタッフ ' + A.compare.staffChecked + '/' + A.compare.staffTotal + '名を比較: 差異 ' + A.compare.differences.length + '件'
      + (A.compare.differences.length ? '<ul>' + A.compare.differences.map(function (x) { return '<li>' + esc(x.name) + ': ' + esc(x.message) + '</li>'; }).join('') + '</ul>' : '')
      + (A.compare.taskDifferences.length ? '<ul>' + A.compare.taskDifferences.map(function (x) { return '<li>' + esc(x.name) + ': ' + esc(x.message) + '</li>'; }).join('') + '</ul>' : '') + '</div>' : '';
    var log = A.audit ? '<div class="ops-table-wrap"><table class="ops-table ops-log"><thead><tr><th>日時</th><th>操作者</th><th>内容</th><th>変更前→後</th></tr></thead><tbody>' + (A.audit.entries.length ? A.audit.entries.map(function (e) {
      return '<tr><td>' + esc(new Date(e.at).toLocaleString('ja-JP')) + '</td><td>' + esc(e.actor) + '</td><td>' + esc(e.action) + '<div class="ops-sub">' + esc(e.target) + '</div></td><td class="ops-code">' + esc(e.before) + ' → ' + esc(e.after) + '</td></tr>';
    }).join('') : '<tr><td colspan="4">履歴はまだありません。</td></tr>') + '</tbody></table></div>' : '';
    return '<h4>Excelとの照合</h4><div class="ops-sub">Excelに入っている合計・ランク・運営ポイントの計算値と、ポータルの自動計算を比較します。</div>'
      + '<div class="ops-toolbar"><button class="member-button" type="button" data-act="compare"' + busyAttr('compare') + '>' + spin('compare', 'Excelと比較する') + '</button>'
      + '<button class="member-button" type="button" data-act="sync"' + busyAttr('sync') + '>' + spin('sync', 'Excelの最新評価を取り込む') + '</button></div>' + cmp
      + '<div class="ops-sub">「最新評価を取り込む」はExcelの評価が変わったスタッフの更新日時を記録します。</div>'
      + '<h4>設定・対応付け・業務割当の変更履歴</h4><div class="ops-toolbar"><button class="member-button" type="button" data-act="audit"' + busyAttr('audit') + '>' + spin('audit', '履歴を読み込む') + '</button></div>' + log;
  }

  /* ---- events ---- */
  function onChange(e) {
    var f = e.target.getAttribute && e.target.getAttribute('data-f');
    if (f) {
      A.f[f] = f === 'upcoming' ? e.target.checked : e.target.value;
      render();
    }
  }

  function refreshDetail(id, payload) {
    A.detail[id] = { judgement: payload.judgement, people: payload.people, tasks: payload.tasks, manual: payload.manual };
    var e = (A.data.events || []).filter(function (x) { return x.eventId === id; })[0], j = payload.judgement;
    if (e) { e.status = j.status; e.label = j.label; e.color = j.color; e.count = j.headcount.count; e.min = j.headcount.min; e.staffShort = j.headcount.shortfall; e.rate = j.points.rate; e.reasons = (j.publicReasons || j.reasons).slice(0, 3); }
    delete A.auto[id];
    loadOverview(); // keep the check-row badges in step
  }

  function assignCall(id, action, params, okText) {
    return run('as:' + id, '保存中', function () {
      return api(action, Object.assign({ eventId: id }, params)).then(function (p) { refreshDetail(id, p); return p; });
    }, okText || '割当を保存し、判定を更新しました。');
  }

  function onClick(e) {
    var el = e.target.closest('[data-act]');
    if (!el || el.disabled) return;
    var act = el.getAttribute('data-act'), id = el.getAttribute('data-event') || el.getAttribute('data-id');
    var d = A.data;
    if (act === 'reload') { A.detail = {}; return loadAdmin(); }
    if (act === 'tab') { A.tab = el.getAttribute('data-tab'); return render(); }
    if (act === 'setup') return run('setup', '準備中', function () { return api('opsSetup'); }, function (r) { loadAdmin(); return '保存用シートを準備しました(' + (r.created || []).length + '件追加)。'; });
    if (act === 'event') {
      A.open[id] = !A.open[id]; render();
      if (A.open[id] && !A.detail[id]) {
        api('opsEventDetail', { eventId: id }).then(function (p) { A.detail[id] = p; }).catch(function (err) { A.detail[id] = { error: String(err && err.message || err) }; }).then(render);
      }
      return;
    }
    if (act === 'assign') {
      var pick = A.root.querySelector('[data-pick="' + el.getAttribute('data-task') + '"]');
      var staffKey = pick && pick.value;
      if (!staffKey) { setMsg('error', '担当にするスタッフを選んでください。'); return render(); }
      return assignCall(id, 'opsAssign', { taskId: el.getAttribute('data-task'), staffKey: staffKey, on: '1' });
    }
    if (act === 'unassign') return assignCall(id, 'opsAssign', { taskId: el.getAttribute('data-task'), staffKey: el.getAttribute('data-staff'), on: '0' }, '担当を外し、判定を更新しました。');
    if (act === 'auto') return run('auto:' + id, '割当案を作成中', function () { return api('opsAutoAssign', { eventId: id }).then(function (p) { A.auto[id] = p; return p; }); }, '割当案を作成しました(まだ確定していません)。');
    if (act === 'apply-sug' || act === 'apply-auto') {
      var items = act === 'apply-sug' ? (A.auto[id] && A.auto[id].proposed) : (A.detail[id] && A.detail[id].judgement.proposed);
      if (!items || !items.length) { setMsg('error', '確定できる割当案がありません。'); return render(); }
      if (!window.confirm('この割当案を確定します(' + items.length + '件)。よろしいですか?')) return;
      return assignCall(id, 'opsAssignBatch', { items: JSON.stringify(items.map(function (x) { return { taskId: x.taskId, staffKey: x.staffKey }; })) }, '割当案を確定しました。');
    }
    if (act === 'confirm-map' || act === 'unlink') {
      var sk = el.getAttribute('data-staff'), email = '';
      if (act === 'confirm-map') {
        var sel = A.root.querySelector('[data-map="' + sk + '"]'); email = sel && sel.value;
        if (!email) { setMsg('error', '紐付けるアカウントを選んでください。'); return render(); }
      } else if (!window.confirm('このスタッフとアカウントの対応付けを解除します。よろしいですか?')) return;
      return run('map', '保存中', function () { return api('opsSaveStaffMap', { items: JSON.stringify([{ staffKey: sk, email: email, status: act === 'unlink' ? 'unlinked' : 'confirmed' }]) }).then(reloadAfter); }, act === 'unlink' ? '対応付けを解除しました。' : '対応付けを確定しました。');
    }
    if (act === 'confirm-all') {
      var singles = d.staff.filter(function (s) { return ['confirmed', 'auto'].indexOf(s.mapping.status) < 0 && s.suggestions.length === 1; });
      if (!window.confirm('次の対応付けを確定します。\n\n' + singles.map(function (s) { return s.name + ' → ' + s.suggestions[0].name + ' (' + s.suggestions[0].email + ')'; }).join('\n') + '\n\n氏名が似ているだけの別人ではないか確認のうえ、よろしければOKを押してください。')) return;
      return run('map', '保存中', function () { return api('opsSaveStaffMap', { items: JSON.stringify(singles.map(function (s) { return { staffKey: s.staffKey, email: s.suggestions[0].email, status: 'confirmed' }; })) }).then(reloadAfter); }, '対応付けを確定しました。');
    }
    if (act === 'edit-rule') { A.edit = el.getAttribute('data-type'); return render(); }
    if (act === 'cancel-rule') { A.edit = null; return render(); }
    if (act === 'save-rule') {
      var form = A.root.querySelector('[data-form="rule"]'), get = function (n) { var i = form.querySelector('[name="' + n + '"]'); return i ? i.value : ''; };
      var counts = {}; d.tasks.forEach(function (t) { counts[t.id] = Number(get('task:' + t.id)); });
      var override = { scope: 'type', key: form.getAttribute('data-type'), requiredPoints: get('requiredPoints'), minStaff: get('minStaff'), taskCounts: counts, note: get('note') };
      if (override.requiredPoints === '' || override.minStaff === '') { setMsg('error', '必要ポイントと最低人数を入力してください。'); return render(); }
      return run('rule', '保存中', function () { return api('opsSaveOverride', { override: JSON.stringify(override) }).then(function (r) { A.edit = null; return reloadAfter(r); }); }, '基準を保存し、判定を更新しました。');
    }
    if (act === 'alias') {
      var sel2 = A.root.querySelector('[data-alias="' + id + '"]');
      if (!sel2 || !sel2.value) { setMsg('error', '紐付ける種別を選んでください。'); return render(); }
      var title = el.getAttribute('data-title'), norm = title.normalize('NFKC').replace(/[\s　]+/g, '').toLowerCase();
      var ev = (d.events || []).filter(function (x) { return x.eventId === id; })[0];
      var rule = { id: 'custom_' + norm.replace(/[^a-z0-9぀-ヿ一-鿿]/g, '').slice(0, 24), eventTypeId: sel2.value, pattern: '^' + norm.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '$', categories: ev && ev.category ? [ev.category] : [], venues: [], priority: 90, note: title };
      if (!window.confirm('「' + title + '」を「' + (d.typeNames[sel2.value] || sel2.value) + '」の基準で判定します。同じ名前のイベントすべてに適用されます。よろしいですか?')) return;
      var next = (d.aliasCustom || []).filter(function (r) { return r.id !== rule.id; }).concat([rule]);
      return run('alias', '保存中', function () { return api('opsSaveAliasRules', { rules: JSON.stringify(next) }).then(reloadAfter); }, '別名を保存し、判定を更新しました。');
    }
    if (act === 'alias-toggle') {
      var nextRules = (d.aliasCustom || []).map(function (r) { return r.id === id ? Object.assign({}, r, { disabled: !r.disabled }) : r; });
      return run('alias', '保存中', function () { return api('opsSaveAliasRules', { rules: JSON.stringify(nextRules) }).then(reloadAfter); }, '別名ルールを更新しました。');
    }
    if (act === 'pair-add') {
      var a = A.root.querySelector('[name="pairA"]').value, b = A.root.querySelector('[name="pairB"]').value;
      if (a === b) { setMsg('error', '異なる2つの業務を選んでください。'); return render(); }
      var cur = readSettingsForm(); cur.concurrentPairs.push([a, b]);
      return saveSettings(cur, '兼務できる組み合わせを追加しました。');
    }
    if (act === 'pair-del') { var cur2 = readSettingsForm(); cur2.concurrentPairs.splice(Number(el.getAttribute('data-i')), 1); return saveSettings(cur2, '組み合わせを削除しました。'); }
    if (act === 'save-settings') return saveSettings(readSettingsForm(), '設定を保存し、全イベントの判定を再計算しました。');
    if (act === 'reset-settings') {
      if (!window.confirm('設定をすべて初期値に戻します。よろしいですか?')) return;
      return saveSettings(clone(d.defaultSettings), '設定を初期値に戻しました。');
    }
    if (act === 'compare') return run('compare', '照合中', function () { return api('opsCompare').then(function (r) { A.compare = r; return r; }); }, function (r) { return r.differences.length ? '差異が' + r.differences.length + '件あります。' : 'Excelの計算値とすべて一致しました。'; });
    if (act === 'sync') return run('sync', '取り込み中', function () { return api('opsSync').then(reloadAfter); }, function (r) { return r && r.changed !== undefined ? '評価が変わったスタッフ ' + r.changed + '名を記録しました。' : '取り込みました。'; });
    if (act === 'audit') return run('audit', '読み込み中', function () { return api('opsAudit').then(function (r) { A.audit = r; return r; }); }, '履歴を読み込みました。');
  }

  /** Only what differs from the defaults is sent/stored (the server merges it back over the defaults). */
  function diffFromDefaults(settings) {
    var defs = A.data.defaultSettings, out = {};
    Object.keys(settings).forEach(function (k) {
      if (k === 'taskWindows') {
        var w = {};
        Object.keys(settings.taskWindows).forEach(function (t) { if (JSON.stringify(settings.taskWindows[t]) !== JSON.stringify((defs.taskWindows || {})[t])) w[t] = settings.taskWindows[t]; });
        if (Object.keys(w).length) out.taskWindows = w;
      } else if (k === 'abilityMap') {
        var m = {};
        Object.keys(settings.abilityMap).forEach(function (t) {
          var base = (A.data.tasks.filter(function (x) { return x.id === t; })[0] || {}).ability;
          if (settings.abilityMap[t] !== base) m[t] = settings.abilityMap[t];
        });
        if (Object.keys(m).length) out.abilityMap = m;
      } else if (JSON.stringify(settings[k]) !== JSON.stringify(defs[k])) out[k] = settings[k];
    });
    return out;
  }

  function saveSettings(settings, okText) {
    return run('settings', '保存中', function () {
      return api('opsSaveSettings', { settings: JSON.stringify(diffFromDefaults(settings)) }).then(reloadAfter);
    }, okText);
  }

  /** After any write: reload the admin data (and drop stale event details) so every number is recalculated. */
  function reloadAfter(result) {
    A.detail = {}; A.auto = {};
    return api('opsAdminOverview').then(function (d) {
      A.data = d;
      Object.keys(A.open).forEach(function (id) {
        if (!A.open[id]) return;
        api('opsEventDetail', { eventId: id }).then(function (p) { A.detail[id] = p; render(); }).catch(function () { /* detail reloads on next open */ });
      });
      loadOverview();
      return result;
    });
  }

  /* ------------------------------------------------------------------ boot */
  function boot() {
    readCache();
    var list = document.getElementById('memberScheduleList');
    if (list) {
      observer = new MutationObserver(function () { decorate(); });
      observer.observe(list, { childList: true });
    }
    // the portal re-renders the list from the saved checks; refresh verdicts right after a save
    try {
      if (typeof callPortalAction === 'function') {
        var orig = callPortalAction;
        callPortalAction = function (action) {
          var p = orig.apply(this, arguments);
          if (/^toggleAssignment/.test(String(action))) { p.then(function () { setTimeout(loadOverview, 400); }, function () { /* ignore */ }); }
          return p;
        };
      }
    } catch (err) { /* fall back to the periodic refresh below */ }
    var lastUser = '';
    setInterval(function () {
      var m = me(), email = m ? String(m.email || '').toLowerCase() : '';
      if (email !== lastUser) {
        lastUser = email; S.overview = null; A.data = null; A.detail = {};
        if (email) { readCache(); decorate(); loadOverview(); }
        else decorate();
        syncAdmin();
      }
    }, 1500);
    setInterval(function () { if (me() && !document.hidden) loadOverview(); }, 120000);
    document.addEventListener('visibilitychange', function () { if (!document.hidden && me()) loadOverview(); });
    syncAdmin();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
