// Связь с сервером правок (apps-script/Code.gs). Без адреса сервера на этом компьютере
// или по ссылке с ?demo=1 работает демо-режим: правки и история живут только в этом браузере.
(function () {
  'use strict';

  var CFG = window.SCHEDULE_CONFIG || {};
  var SESSION_KEY = 'raspisanie.admin.v1';
  var DEMO_EDITS = 'raspisanie.demo.edits.v1';
  var DEMO_HISTORY = 'raspisanie.demo.history.v1';
  var url = String(CFG.adminApi || '').trim();
  var demo = !url && (/^(localhost|127\.0\.0\.1)$/.test(location.hostname) || /[?&]demo=1(&|$)/.test(location.search));
  var mode = url ? 'live' : demo ? 'demo' : 'off';

  function readJSON(key, fallback) {
    try { var v = JSON.parse(localStorage.getItem(key) || 'null'); return v == null ? fallback : v; } catch (e) { return fallback; }
  }
  function writeJSON(key, v) {
    try { localStorage.setItem(key, JSON.stringify(v)); return true; } catch (e) { return false; }
  }

  // ---------- session ----------
  var session = readJSON(SESSION_KEY, null);
  if (!session || !session.token || !session.user || session.expires < Date.now() || session.mode !== mode) session = null;
  function setSession(s) {
    session = s;
    try { if (s) localStorage.setItem(SESSION_KEY, JSON.stringify(s)); else localStorage.removeItem(SESSION_KEY); } catch (e) { /* storage unavailable */ }
  }

  // ---------- live server ----------
  // text/plain keeps the request "simple", so the browser sends it to Apps Script without a preflight
  function post(body) {
    return fetch(url, { method: 'POST', body: JSON.stringify(body), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow' })
      .then(function (r) { if (!r.ok) throw new Error('Сервер ответил ' + r.status); return r.json(); });
  }
  function get(params) {
    var q = Object.keys(params).map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]); }).join('&');
    return fetch(url + (url.indexOf('?') >= 0 ? '&' : '?') + q, { cache: 'no-store', redirect: 'follow' })
      .then(function (r) { if (!r.ok) throw new Error('Сервер ответил ' + r.status); return r.json(); });
  }
  function checked(res) {
    if (res && res.ok) return res;
    if (res && /Сессия/.test(res.error || '')) setSession(null);
    var err = new Error(res && res.error ? res.error : 'Сервер не ответил');
    err.data = res;
    throw err;
  }

  // ---------- demo server: the same rules as Code.gs, kept in this browser ----------
  var DEMO_USER = { login: 'demo', name: 'Демо-администратор', role: 'главный' };
  function demoId(p) { return p + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4); }
  function slotKey(e) { return [e.group, e.date, e.start, e.sub || 0].join('|'); }
  function snapshot(e) { return e ? JSON.stringify({ kind: e.kind, disc: e.disc, teacher: e.teacher, room: e.room, sub: e.sub, note: e.note, status: e.status }) : ''; }
  function demoLog(action, e, before, after, comment) {
    var h = readJSON(DEMO_HISTORY, []);
    h.push({ id: demoId('h'), time: new Date().toISOString(), login: DEMO_USER.login, name: DEMO_USER.name, action: action, editId: e.id, group: e.group, date: e.date, start: e.start, before: before, after: after, comment: comment || '' });
    writeJSON(DEMO_HISTORY, h.slice(-1000));
  }
  function demoClean(e) {
    var out = {
      id: e.id || '', group: String(e.group || ''), date: String(e.date || ''), start: String(e.start || ''), end: String(e.end || ''),
      kind: e.kind, disc: String(e.disc || '').trim(), teacher: String(e.teacher || '').trim(), room: String(e.room || '').trim(),
      sub: +e.sub || 0, note: String(e.note || '').trim(), was: typeof e.was === 'string' ? e.was : JSON.stringify(e.was || ''), link: e.link || ''
    };
    if (!out.group) throw new Error('Не указана группа');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(out.date)) throw new Error('Дата должна быть в виде ГГГГ-ММ-ДД');
    if (!/^\d{1,2}:\d{2}$/.test(out.start)) throw new Error('Время начала должно быть в виде Ч:ММ');
    if (['replace', 'cancel', 'add'].indexOf(out.kind) < 0) throw new Error('Неизвестный вид правки');
    if (out.kind !== 'cancel' && !out.disc) throw new Error('Укажите дисциплину');
    if (out.kind === 'cancel') { out.disc = ''; out.teacher = ''; out.room = ''; }
    return out;
  }
  var demoServer = {
    edits: function () { return { ok: true, edits: readJSON(DEMO_EDITS, []).filter(function (e) { return e.status === 'active'; }), at: Date.now() }; },
    login: function (b) {
      if (String(b.login).trim().toLowerCase() !== 'demo' || b.password !== 'demo') return { ok: false, error: 'В демо-режиме логин demo и пароль demo' };
      return { ok: true, token: demoId('t'), user: DEMO_USER, expires: Date.now() + 6 * 3600e3 };
    },
    save: function (b) {
      var all = readJSON(DEMO_EDITS, []), saved = [], items = (b.edits || [b.edit]).map(demoClean);
      var link = items.length > 1 ? demoId('l') : items[0].link;
      items.forEach(function (it) {
        if (items.length > 1) it.link = link;
        var now = new Date().toISOString();
        var cur = all.filter(function (e) { return (it.id && e.id === it.id) || (e.status === 'active' && slotKey(e) === slotKey(it)); })[0];
        if (cur) {
          var before = cur.status === 'active' ? snapshot(cur) : '';
          ['group', 'date', 'start', 'end', 'kind', 'disc', 'teacher', 'room', 'sub', 'note', 'link'].forEach(function (k) { cur[k] = it[k]; });
          if (it.was) cur.was = it.was;
          cur.status = 'active'; cur.updated = now; cur.author = DEMO_USER.login;
          demoLog(before ? 'update' : 'create', cur, before, snapshot(cur), b.comment);
          saved.push(cur);
        } else {
          it.id = demoId('e'); it.created = now; it.updated = now; it.author = DEMO_USER.login; it.status = 'active';
          all.push(it);
          demoLog('create', it, '', snapshot(it), b.comment);
          saved.push(it);
        }
      });
      writeJSON(DEMO_EDITS, all);
      return { ok: true, edits: saved };
    },
    remove: function (b) {
      var all = readJSON(DEMO_EDITS, []), ids = b.ids || [], links = {}, removed = [];
      all.forEach(function (e) { if (ids.indexOf(e.id) >= 0 && e.link) links[e.link] = true; });
      all.forEach(function (e) {
        if (e.status !== 'active' || (ids.indexOf(e.id) < 0 && !(e.link && links[e.link]))) return;
        var before = snapshot(e);
        e.status = 'deleted'; e.updated = new Date().toISOString();
        demoLog('delete', e, before, '', b.comment);
        removed.push(e.id);
      });
      writeJSON(DEMO_EDITS, all);
      return { ok: true, removed: removed };
    },
    history: function (b) {
      var items = readJSON(DEMO_HISTORY, []).filter(function (h) { return !b.group || h.group === b.group; }).slice(-(b.limit || 300)).reverse();
      return { ok: true, items: items, full: true };
    },
    password: function () { return { ok: false, error: 'В демо-режиме пароль не меняется' }; }
  };
  function demoCall(action, body) {
    return new Promise(function (resolve) {
      setTimeout(function () {
        try { resolve(demoServer[action](body || {})); } catch (e) { resolve({ ok: false, error: e.message }); }
      }, 150);
    });
  }

  function call(action, body, isGet) {
    if (mode === 'off') return Promise.reject(new Error('Сервер правок не подключён'));
    if (mode === 'demo') return demoCall(action, body).then(checked);
    var payload = Object.assign({ action: action }, body || {});
    return (isGet ? get(payload) : post(payload)).then(checked);
  }
  function withToken(body) { return Object.assign({ token: session ? session.token : '' }, body || {}); }

  window.RaspApi = {
    mode: mode,
    user: function () { return session ? session.user : null; },
    loadEdits: function () { return call('edits', {}, true).then(function (r) { return r.edits || []; }); },
    login: function (login, password) {
      return call('login', { login: login, password: password }).then(function (r) {
        setSession({ token: r.token, user: r.user, expires: r.expires, mode: mode });
        return r.user;
      });
    },
    logout: function () {
      var t = session ? session.token : '';
      setSession(null);
      return mode === 'live' && t ? post({ action: 'logout', token: t }).catch(function () { return null; }) : Promise.resolve();
    },
    save: function (edits, comment) { return call('save', withToken({ edits: edits, comment: comment || '' })).then(function (r) { return r.edits; }); },
    remove: function (ids, comment) { return call('remove', withToken({ ids: ids, comment: comment || '' })).then(function (r) { return r.removed; }); },
    history: function (group) {
      // a signed-in admin sees who made each change, so the token goes with the request
      return call('history', withToken({ group: group || '', limit: 500 }), mode === 'live' && !session).then(function (r) { return r; });
    },
    password: function (old, next) { return call('password', withToken({ old: old, next: next })); }
  };
})();
