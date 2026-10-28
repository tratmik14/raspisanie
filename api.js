// Связь с сервером правок (apps-script/Code.gs). Без адреса сервера на этом компьютере
// или по ссылке с ?demo=1 работает демо-режим: тот же Code.gs запускается в браузере (gas-sim.js),
// и правки, история и журнал живут только в этом браузере.
(function () {
  'use strict';

  var CFG = window.SCHEDULE_CONFIG || {};
  var SESSION_KEY = 'raspisanie.admin.v1';
  var url = String(CFG.adminApi || '').trim();
  var demo = !url && !!window.RaspSim && (/^(localhost|127\.0\.0\.1)$/.test(location.hostname) || /[?&]demo=1(&|$)/.test(location.search));
  var mode = url ? 'live' : demo ? 'demo' : 'off';

  function readJSON(key, fallback) {
    try { var v = JSON.parse(localStorage.getItem(key) || 'null'); return v == null ? fallback : v; } catch (e) { return fallback; }
  }

  // ---------- session ----------
  var session = readJSON(SESSION_KEY, null);
  if (!session || !session.token || !session.user || session.expires < Date.now() || session.mode !== mode) session = null;
  function setSession(s) {
    session = s;
    try { if (s) localStorage.setItem(SESSION_KEY, JSON.stringify(s)); else localStorage.removeItem(SESSION_KEY); } catch (e) { /* storage unavailable */ }
  }

  // ---------- transport ----------
  // text/plain keeps the request "simple", so the browser sends it to Apps Script without a preflight
  function post(body) {
    if (mode === 'demo') return window.RaspSim.call(body);
    return fetch(url, { method: 'POST', body: JSON.stringify(body), headers: { 'Content-Type': 'text/plain;charset=utf-8' }, redirect: 'follow' })
      .then(function (r) { if (!r.ok) throw new Error('Сервер ответил ' + r.status); return r.json(); });
  }
  function get(params) {
    if (mode === 'demo') return window.RaspSim.call(params);
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
  // public reads go as GET (cacheable, no token); anything with a session goes as POST
  function call(action, body) {
    if (mode === 'off') return Promise.reject(new Error('Сервер правок не подключён'));
    var payload = Object.assign({ action: action }, body || {});
    if (session && session.token && payload.token == null) payload.token = session.token;
    return (payload.token ? post(payload) : get(payload)).then(checked);
  }
  function admin(action, body) { return call(action, Object.assign({ token: session ? session.token : '' }, body || {})); }

  window.RaspApi = {
    mode: mode,
    user: function () { return session ? session.user : null; },
    // everything laid over the timetable: edits, schedule corrections, imported exams
    loadState: function () { return call('state'); },
    login: function (login, password) {
      return post({ action: 'login', login: login, password: password }).then(checked).then(function (r) {
        setSession({ token: r.token, user: r.user, expires: r.expires, mode: mode });
        return r.user;
      });
    },
    logout: function () {
      var t = session ? session.token : '';
      setSession(null);
      return t ? post({ action: 'logout', token: t }).catch(function () { return null; }) : Promise.resolve();
    },
    save: function (edits, opts) {
      opts = opts || {};
      return admin('save', { edits: edits, batch: opts.batch || '', comment: opts.comment || '' }).then(function (r) { return r.edits; });
    },
    remove: function (ids, opts) {
      opts = opts || {};
      return admin('remove', { ids: ids || [], batch: opts.batch || '', comment: opts.comment || '' }).then(function (r) { return r.removed; });
    },
    revert: function (historyId) { return admin('revert', { id: historyId }); },
    history: function (group) { return call('history', { group: group || '', limit: 500 }); },
    grafikSave: function (items, comment) { return admin('grafikSave', { items: items, comment: comment || '' }).then(function (r) { return r.items; }); },
    snapshots: function () { return admin('snapshots').then(function (r) { return r.snapshots || {}; }); },
    importTable: function (payload) { return admin('importTable', payload); },
    tableJournal: function (group) { return call('tableJournal', { group: group || '', limit: 1500 }); },
    examsImport: function (records, source) { return admin('examsImport', { records: records, source: source || '' }); },
    fetchDoc: function (docUrl) { return admin('fetchDoc', { url: docUrl }).then(function (r) { return r.html; }); },
    password: function (old, next) { return admin('password', { old: old, next: next }); }
  };
})();
