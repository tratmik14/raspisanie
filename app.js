(function () {
  'use strict';

  var CFG = window.SCHEDULE_CONFIG || {};
  var DAY = 864e5;
  var PREFS_KEY = 'raspisanie.prefs.v1';
  var CACHE_KEY = 'raspisanie.sheet.v1:';
  var CONCURRENCY = 6;

  var MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  var WEEKDAYS = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  var WEEKDAYS_MON = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  var WEEKDAY_NAMES = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];
  var TEACHER_RE = /^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.?)?$/;
  var GROUP_RE = /^[А-ЯЁA-Z][А-ЯЁA-Zа-яё0-9]*-\d{2}[А-ЯЁа-яё]?$/;

  function $(id) { return document.getElementById(id); }

  // ---------- small helpers ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function todayISO() { var t = new Date(); return ymd(t.getFullYear(), t.getMonth() + 1, t.getDate()); }
  function isoParts(iso) { var p = iso.split('-'); return { y: +p[0], m: +p[1], d: +p[2] }; }
  function isoToDate(iso) { var p = isoParts(iso); return new Date(p.y, p.m - 1, p.d); }
  function utcOfIso(iso) { var p = isoParts(iso); return Date.UTC(p.y, p.m - 1, p.d); }
  function isoOfUtc(ms) { var d = new Date(ms); return ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
  function mondayOfUtc(ms) { return ms - ((new Date(ms).getUTCDay() + 6) % 7) * DAY; }
  function daysBetween(a, b) { return Math.round((utcOfIso(b) - utcOfIso(a)) / DAY); }
  function plural(n, one, few, many) {
    var a = Math.abs(n) % 100, b = a % 10;
    if (a > 10 && a < 20) return many;
    if (b === 1) return one;
    if (b > 1 && b < 5) return few;
    return many;
  }
  function count(n, one, few, many) { return n.toLocaleString('ru-RU') + ' ' + plural(n, one, few, many); }
  function fmtDayMonth(iso) { var d = isoToDate(iso); return d.getDate() + ' ' + MONTHS_GEN[d.getMonth()]; }
  function fmtRange(a, b) {
    var x = isoToDate(a), y = isoToDate(b);
    if (x.getMonth() === y.getMonth()) return x.getDate() + '–' + y.getDate() + ' ' + MONTHS_GEN[y.getMonth()];
    return x.getDate() + ' ' + MONTHS_GEN[x.getMonth()] + ' – ' + y.getDate() + ' ' + MONTHS_GEN[y.getMonth()];
  }
  function fmtStamp(ms) {
    var d = new Date(ms), t = pad(d.getHours()) + ':' + pad(d.getMinutes());
    var iso = ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
    return iso === todayISO() ? 'сегодня в ' + t : fmtDayMonth(iso) + ' в ' + t;
  }
  function relLabel(diff) {
    if (diff === 0) return 'сегодня';
    if (diff === 1) return 'завтра';
    if (diff === 2) return 'послезавтра';
    if (diff === -1) return 'вчера';
    return '';
  }
  function minutesOf(t) { var p = String(t).split(':'); return (+p[0]) * 60 + (+p[1] || 0); }
  function normTime(t) { var p = String(t).split(':'); return (+p[0]) + ':' + pad(+p[1] || 0); }
  function cmpRu(a, b) { return a.localeCompare(b, 'ru', { numeric: true }); }
  function splitCode(disc) {
    var m = /^([А-ЯЁA-Z]{2,5}[а-яё]?\.?\s?\d{2}(?:\.\d{2})?)\.?\s+(.+)$/.exec(disc);
    return m ? { code: m[1], name: m[2] } : { code: '', name: disc };
  }
  function discHTML(disc) {
    var dc = splitCode(disc);
    return (dc.code ? '<span class="code">' + esc(dc.code) + '</span>' : '') + esc(dc.name);
  }
  function roomLabel(room) { if (!room) return ''; return /^\d/.test(room) ? 'ауд. ' + room : room; }
  var TRANSLIT = { 'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e', 'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm', 'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u', 'ф': 'f', 'х': 'h', 'ц': 'c', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch', 'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya' };
  function slug(s) {
    return String(s).toLowerCase().split('').map(function (c) { return TRANSLIT[c] != null ? TRANSLIT[c] : c; }).join('')
      .replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
  }
  function validDate(y, m, d) {
    var t = new Date(Date.UTC(y, m - 1, d));
    return y > 1900 && t.getUTCFullYear() === y && t.getUTCMonth() === m - 1 && t.getUTCDate() === d;
  }

  // ---------- Google Sheets ----------
  // headers=0 keeps Google from gluing the top rows of a sheet into one header row
  function sheetUrl(name) {
    return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(CFG.sheetId) +
      '/gviz/tq?tqx=out:csv&headers=0&sheet=' + encodeURIComponent(name);
  }
  function fetchSheet(name) {
    return fetch(sheetUrl(name), { cache: 'no-store' }).then(function (res) {
      var type = res.headers.get('content-type') || '';
      if (!res.ok || type.indexOf('csv') < 0) throw new Error('Google Sheets answered ' + res.status + ' ' + type);
      return res.text();
    }).then(function (csv) {
      cachePut(name, csv);
      return csv;
    });
  }
  function cacheGet(name) {
    try {
      var v = JSON.parse(localStorage.getItem(CACHE_KEY + name) || 'null');
      return v && typeof v.csv === 'string' ? v : null;
    } catch (e) { return null; }
  }
  function cachePut(name, csv) {
    try { localStorage.setItem(CACHE_KEY + name, JSON.stringify({ t: Date.now(), csv: csv })); }
    catch (e) { /* storage full or blocked: the page still works, only without the saved copy */ }
  }
  function parseCSV(text) {
    var rows = [], row = [], field = '', quoted = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (quoted) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false;
        } else field += c;
      } else if (c === '"') quoted = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += c;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows;
  }

  // ---------- reading a group sheet ----------
  function weekdayIndex(v) { return WEEKDAY_NAMES.indexOf(String(v).trim().toLowerCase()); }
  function parityOf(v) {
    var s = String(v).trim().toLowerCase().replace(/ё/g, 'е');
    if (/^четн(ая|ые)?( неделя)?$/.test(s)) return 1;
    if (/^нечетн(ая|ые)?( неделя)?$/.test(s)) return 2;
    return 0;
  }
  function cellDayUtc(v) {
    var s = String(v).trim();
    var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);
    if (m) return validDate(+m[3], +m[2], +m[1]) ? Date.UTC(+m[3], +m[2] - 1, +m[1]) : null;
    if (/^\d{5}(\.0+)?$/.test(s)) {
      var n = parseFloat(s); // a date left as an Excel day number
      if (n > 30000 && n < 80000) return Date.UTC(1899, 11, 30) + Math.floor(n) * DAY;
    }
    return null;
  }
  function modeMonday(days) {
    var counts = {}, best = null;
    days.forEach(function (d) {
      var m = mondayOfUtc(d);
      counts[m] = (counts[m] || 0) + 1;
      if (best === null || counts[m] > counts[best]) best = m;
    });
    return best;
  }
  function parseSlotCell(v) {
    var m = /(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})/.exec(String(v).replace(/\s+/g, ' '));
    return m ? { start: (+m[1]) + ':' + m[2], end: (+m[3]) + ':' + m[4] } : null;
  }
  function buildLessonEntry(lines, parity) {
    var ti = -1;
    for (var i = 1; i < lines.length; i++) if (TEACHER_RE.test(lines[i]) || /^вакансия(?![а-яё])/i.test(lines[i])) { ti = i; break; }
    var disc, teacher = '', rest;
    if (ti > 0) { disc = lines.slice(0, ti).join(' '); teacher = lines[ti]; rest = lines.slice(ti + 1).join(' '); }
    else { disc = lines[0]; rest = lines.slice(1).join(' '); }
    if (!/[а-яёa-z]/i.test(disc)) return null;
    if (/^вакансия/i.test(teacher)) teacher = '';
    var sub = 0;
    rest = rest.replace(/\(?\s*(\d)\s*(?:п\/)?\s*гр\.?\s*\)?/i, function (_, n) { sub = +n; return ''; }).replace(/\s+/g, ' ').trim();
    return { disc: disc.replace(/\s+/g, ' ').trim(), teacher: teacher, room: rest, sub: sub, parity: parity };
  }
  // A cell holds "discipline / teacher / room" lines; a blank line separates subgroups,
  // and a "Чётная"/"Нечётная" line marks the entry that follows it.
  function parseLessonCell(v) {
    if (!/[а-яё]/i.test(v)) return [];
    var lines = String(v).split(/\r?\n/).map(function (l) { return l.replace(/\s+/g, ' ').trim(); });
    var out = [], cur = [], pend = 0;
    function flush() {
      if (cur.length) { var e = buildLessonEntry(cur, pend); if (e) out.push(e); pend = 0; }
      cur = [];
    }
    lines.forEach(function (l) {
      if (!l) { flush(); return; }
      var p = parityOf(l);
      if (p) { flush(); pend = p; return; }
      cur.push(l);
    });
    flush();
    return out;
  }
  // Splits a sheet into week blocks: a row of dates, a row of weekday names, then one row per pair.
  function scanLessonSheet(rows) {
    var blocks = [], cur = null, dayCols = null, timeCol = -1, pendingMonday = null, dateFresh = false, hasDates = false;
    rows.forEach(function (raw) {
      var row = raw.slice(0, 9);
      var wds = [], days = [], pw = 0;
      row.forEach(function (v, c) {
        var w = weekdayIndex(v); if (w >= 0) wds.push([c, w]);
        var d = cellDayUtc(v); if (d != null) days.push(d);
        if (c < 2 && !pw) pw = parityOf(v);
      });
      if (days.length >= 3) {
        pendingMonday = modeMonday(days); dateFresh = true; hasDates = true;
        return;
      }
      if (wds.length >= 3) {
        dayCols = {};
        wds.forEach(function (x) { if (x[1] < 6) dayCols[x[0]] = x[1]; });
        timeCol = wds[0][0] - wds[0][1] - 1;
        cur = { monday: dateFresh ? pendingMonday : null, votes: [0, 0, 0], cells: [] };
        blocks.push(cur);
        dateFresh = false;
        if (pw) cur.votes[pw]++;
        return;
      }
      if (!cur) return;
      var slot = timeCol >= 0 && row[timeCol] != null ? parseSlotCell(row[timeCol]) : null;
      for (var c = 0; !slot && c < 3 && c < row.length; c++) slot = parseSlotCell(row[c]);
      if (!slot) return;
      if (pw) cur.votes[pw]++;
      Object.keys(dayCols).forEach(function (col) {
        var entries = row[col] != null ? parseLessonCell(row[col]) : [];
        if (entries.length) cur.cells.push({ wd: dayCols[col], slot: slot, entries: entries });
      });
    });
    return { blocks: blocks, hasDates: hasDates };
  }

  // ---------- data store ----------
  var DB = {
    groups: [],      // from the group list sheet, in its order
    calendar: {},    // monday (UTC ms) → 1 even, 2 odd, 0 unknown
    sheets: {},      // group → { scan, rows, byDay, template, source: 'live' | 'cache', at }
    weeks: [],       // [{ ms, iso, par }]
    parity: {},      // monday ms → parity
    slots: [],       // [{ n, start, end }]
    slotByStart: {},
    teachers: []
  };
  var net = { pending: 0, failed: {}, listLive: false, listFailed: false, bulk: null };

  function applyGroupList(csv) {
    var seen = {}, groups = [];
    parseCSV(csv).forEach(function (row) {
      var g = String(row[0] || '').trim();
      if (GROUP_RE.test(g) && !seen[g]) { seen[g] = true; groups.push(g); }
    });
    DB.groups = groups;
  }
  function applyCalendar(csv) {
    var cal = {};
    parseCSV(csv).forEach(function (row) {
      var days = [], p = 0;
      row.forEach(function (v) { var d = cellDayUtc(v); if (d != null) days.push(d); if (!p) p = parityOf(v); });
      if (days.length >= 3) { var m = modeMonday(days); cal[m] = p || cal[m] || 0; }
    });
    DB.calendar = cal;
  }
  function applySheet(group, csv, source, at) {
    DB.sheets[group] = { scan: scanLessonSheet(parseCSV(csv)), source: source, at: at, rows: [], byDay: {}, template: false };
  }

  function recompute() {
    var set = {}, votes = {};
    Object.keys(DB.calendar).forEach(function (m) { set[m] = true; });
    Object.keys(DB.sheets).forEach(function (g) {
      DB.sheets[g].scan.blocks.forEach(function (b) {
        if (b.monday == null) return;
        set[b.monday] = true;
        var v = votes[b.monday] || (votes[b.monday] = [0, 0, 0]);
        v[1] += b.votes[1]; v[2] += b.votes[2];
      });
    });
    var mondays = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
    var par = mondays.map(function (m) {
      if (DB.calendar[m]) return DB.calendar[m];
      var v = votes[m];
      return v && v[1] !== v[2] ? (v[1] > v[2] ? 1 : 2) : 0;
    });
    var known = [];
    par.forEach(function (p, i) { if (p) known.push(i); });
    par = par.map(function (p, i) {
      if (p || !known.length) return p;
      var j = known.reduce(function (a, b) { return Math.abs(b - i) < Math.abs(a - i) ? b : a; });
      return (i - j) % 2 === 0 ? par[j] : 3 - par[j];
    });
    DB.weeks = mondays.map(function (m, i) { return { ms: m, iso: isoOfUtc(m), par: par[i] }; });
    DB.parity = {};
    DB.weeks.forEach(function (w) { DB.parity[w.ms] = w.par; });

    Object.keys(DB.sheets).forEach(function (g) { buildRows(g); });

    var slotMap = {};
    (CFG.slots || []).forEach(function (s) { slotMap[normTime(s[0])] = { start: normTime(s[0]), end: normTime(s[1]) }; });
    var teacherSet = {};
    Object.keys(DB.sheets).forEach(function (g) {
      DB.sheets[g].rows.forEach(function (r) {
        if (!slotMap[r.start]) slotMap[r.start] = { start: r.start, end: r.end };
        if (r.teacher) teacherSet[r.teacher] = true;
      });
    });
    DB.slots = Object.keys(slotMap).map(function (k) { return slotMap[k]; })
      .sort(function (a, b) { return minutesOf(a.start) - minutesOf(b.start); });
    DB.slotByStart = {};
    DB.slots.forEach(function (s, i) { s.n = i + 1; DB.slotByStart[s.start] = s; });
    DB.teachers = Object.keys(teacherSet).sort(cmpRu);
  }

  function buildRows(group) {
    var sheet = DB.sheets[group], rows = [], seen = {};
    function emit(mondayMs, cells, template) {
      var wp = DB.parity[mondayMs] || 0;
      cells.forEach(function (cell) {
        var ps = cell.entries.map(function (e) { return e.parity; });
        var both = ps.indexOf(1) >= 0 && ps.indexOf(2) >= 0;
        cell.entries.forEach(function (e) {
          var keep = 0;
          // an entry marked for one parity is dropped on the other weeks only when the cell
          // holds both variants or the sheet is a template; a lone mark on a dated week stays
          if (e.parity && (both || template)) {
            if (wp) { if (e.parity !== wp) return; } else keep = e.parity;
          }
          rows.push({
            group: group, day: isoOfUtc(mondayMs + cell.wd * DAY), start: cell.slot.start, end: cell.slot.end,
            disc: e.disc, teacher: e.teacher, room: e.room, sub: e.sub, par: keep
          });
        });
      });
    }
    if (sheet.scan.hasDates) {
      sheet.scan.blocks.forEach(function (b) {
        if (b.monday == null || seen[b.monday]) return;
        seen[b.monday] = true;
        emit(b.monday, b.cells, false);
      });
    } else {
      var cells = [].concat.apply([], sheet.scan.blocks.map(function (b) { return b.cells; }));
      DB.weeks.forEach(function (w) { emit(w.ms, cells, true); });
    }
    sheet.rows = rows;
    sheet.template = !sheet.scan.hasDates && rows.length > 0;
    sheet.byDay = {};
    rows.forEach(function (r) { (sheet.byDay[r.day] || (sheet.byDay[r.day] = [])).push(r); });
  }

  // ---------- state ----------
  var state = { mode: 'group', group: '', teacher: '', week: '' };
  var pendingHash = ''; // a link to a group or teacher that is not loaded yet
  try {
    var saved = JSON.parse(localStorage.getItem(PREFS_KEY) || 'null');
    if (saved && typeof saved === 'object') {
      ['mode', 'group', 'teacher'].forEach(function (k) { if (typeof saved[k] === 'string') state[k] = saved[k]; });
    }
  } catch (e) { /* storage unavailable */ }
  if (state.mode !== 'teacher') state.mode = 'group';
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify({ mode: state.mode, group: state.group, teacher: state.teacher })); }
    catch (e) { /* storage unavailable */ }
  }
  function readHash() {
    try { return decodeURIComponent(location.hash.replace(/^#/, '')).toLowerCase(); } catch (e) { return ''; }
  }
  function allLive() { return DB.groups.length > 0 && DB.groups.every(isLive); }
  function applyHash(h) {
    h = h || pendingHash;
    pendingHash = '';
    if (!h) return;
    if (h.indexOf('p-') === 0) {
      state.mode = 'teacher';
      var t = DB.teachers.filter(function (x) { return 'p-' + slug(x) === h; })[0];
      if (t) state.teacher = t;
      else if (net.bulk || !allLive()) pendingHash = h; // the teacher list fills in as groups load
      return;
    }
    var g = DB.groups.filter(function (x) { return slug(x) === h || x.toLowerCase() === h; })[0];
    if (g) { state.mode = 'group'; state.group = g; }
    else if (!net.listLive && !net.listFailed) pendingHash = h;
  }
  function writeHash() {
    if (pendingHash) return;
    var target = '';
    if (state.mode === 'teacher' && state.teacher) target = '#p-' + slug(state.teacher);
    else if (state.mode === 'group' && state.group) target = '#' + slug(state.group);
    if ((location.hash || '') === target) return;
    try { history.replaceState(null, '', target || location.pathname + location.search); } catch (e) { /* ignore */ }
  }

  // ---------- loading ----------
  function loadSheet(group) {
    net.pending++;
    return fetchSheet(group).then(function (csv) {
      applySheet(group, csv, 'live', Date.now());
      delete net.failed[group];
    }, function () {
      net.failed[group] = true;
      if (!DB.sheets[group]) {
        var c = cacheGet(group);
        if (c) applySheet(group, c.csv, 'cache', c.t);
      }
    }).then(function () { net.pending--; });
  }
  function isLive(group) { return DB.sheets[group] && DB.sheets[group].source === 'live'; }
  function ensureGroup(group) {
    if (!group || isLive(group) || net.loading === group) return;
    if (!DB.sheets[group]) {
      var c = cacheGet(group);
      if (c) { applySheet(group, c.csv, 'cache', c.t); recompute(); }
    }
    net.loading = group;
    render();
    loadSheet(group).then(function () {
      net.loading = null;
      recompute();
      render();
    });
  }
  function ensureAll() {
    var todo = DB.groups.filter(function (g) { return !isLive(g); });
    if (!todo.length || net.bulk) return;
    todo.forEach(function (g) {
      if (!DB.sheets[g]) { var c = cacheGet(g); if (c) applySheet(g, c.csv, 'cache', c.t); }
    });
    recompute();
    net.bulk = { done: 0, total: todo.length };
    render();
    var queue = todo.slice();
    function worker() {
      var g = queue.shift();
      if (!g) return Promise.resolve();
      return loadSheet(g).then(function () {
        net.bulk.done++;
        if (net.bulk.done % 5 === 0) { recompute(); render(); }
        return worker();
      });
    }
    var workers = [];
    for (var i = 0; i < CONCURRENCY; i++) workers.push(worker());
    Promise.all(workers).then(function () {
      net.bulk = null;
      recompute();
      applyHash();
      fillOptions();
      render();
    });
  }
  function ensureView() {
    if (state.mode === 'teacher') ensureAll();
    else if (state.group) ensureGroup(state.group);
  }
  function loadLists() {
    net.pending++;
    net.listFailed = false;
    render();
    var cal = CFG.calendarSheet ? fetchSheet(CFG.calendarSheet).catch(function () { return null; }) : Promise.resolve(null);
    return Promise.all([fetchSheet(CFG.groupListSheet), cal]).then(function (res) {
      applyGroupList(res[0]);
      if (res[1]) applyCalendar(res[1]);
      net.listLive = true;
    }, function () {
      net.listFailed = true;
    }).then(function () {
      net.pending--;
      recompute();
      applyHash();
      fillOptions();
      render();
      ensureView();
    });
  }
  function retry() {
    net.failed = {};
    if (!net.listLive) loadLists(); else ensureView();
  }

  // ---------- rendering ----------
  function courseOf(group) {
    var m = /-(\d{2})\D*$/.exec(group);
    if (!m) return 0;
    var t = new Date(), start = t.getMonth() >= 6 ? t.getFullYear() : t.getFullYear() - 1;
    var c = start - (2000 + (+m[1])) + 1;
    return c >= 1 && c <= 6 ? c : 0;
  }
  function groupsByCourse() {
    var out = {};
    DB.groups.forEach(function (g) { var c = courseOf(g); (out[c] || (out[c] = [])).push(g); });
    return Object.keys(out).map(Number).sort(function (a, b) { return (a || 99) - (b || 99); }).map(function (c) {
      return { label: c ? c + ' курс' : 'Другие группы', groups: out[c].sort(cmpRu) };
    });
  }
  function fillOptions() {
    var courses = groupsByCourse();
    $('group').innerHTML = '<option value="">Выберите группу</option>' + courses.map(function (c) {
      return '<optgroup label="' + esc(c.label) + '">' + c.groups.map(function (g) {
        return '<option value="' + esc(g) + '">' + esc(g) + '</option>';
      }).join('') + '</optgroup>';
    }).join('');
    var teachers = DB.teachers.slice();
    if (state.teacher && teachers.indexOf(state.teacher) < 0) teachers.unshift(state.teacher);
    $('teacher').innerHTML = '<option value="">' + (net.bulk ? 'Загружаю список…' : 'Выберите преподавателя') + '</option>' + teachers.map(function (t) {
      return '<option value="' + esc(t) + '">' + esc(t) + '</option>';
    }).join('');
    $('picker-body').innerHTML = courses.map(function (c) {
      return '<div class="course"><h3 class="eyebrow">' + esc(c.label) + '</h3><div class="chips">' + c.groups.map(function (g) {
        return '<button type="button" class="chip" data-group="' + esc(g) + '">' + esc(g) + '</button>';
      }).join('') + '</div></div>';
    }).join('');
  }

  function currentMondayIso() {
    var t = new Date();
    var ms = Date.UTC(t.getFullYear(), t.getMonth(), t.getDate());
    if (t.getDay() === 0) ms += DAY; // on Sunday show the coming week
    return isoOfUtc(mondayOfUtc(ms));
  }
  function nearestWeekIndex(iso) {
    var weeks = DB.weeks;
    if (!weeks.length) return -1;
    for (var i = 0; i < weeks.length; i++) if (weeks[i].iso >= iso) return i;
    return weeks.length - 1;
  }
  function selection() {
    if (state.mode === 'teacher') return state.teacher ? { mode: 'teacher', name: state.teacher } : null;
    return state.group ? { mode: 'group', name: state.group } : null;
  }
  function rowsForDay(iso, sel) {
    if (sel.mode === 'group') {
      var s = DB.sheets[sel.name];
      return s && s.byDay[iso] ? s.byDay[iso] : [];
    }
    var out = [];
    Object.keys(DB.sheets).forEach(function (g) {
      (DB.sheets[g].byDay[iso] || []).forEach(function (r) { if (r.teacher === sel.name) out.push(r); });
    });
    return out;
  }
  function slotOf(r) { return DB.slotByStart[r.start] || { n: 0, start: r.start, end: r.end }; }

  // In teacher mode the same class given to several groups at once becomes one entry.
  function pairEntries(rows, sel) {
    var map = {}, order = [];
    rows.forEach(function (r) {
      var key = sel.mode === 'teacher' ? [r.disc, r.room, r.sub, r.par].join('|') : [r.disc, r.teacher, r.room, r.sub, r.par].join('|');
      if (!map[key]) { map[key] = { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub, par: r.par, groups: [] }; order.push(key); }
      if (map[key].groups.indexOf(r.group) < 0) map[key].groups.push(r.group);
    });
    return order.map(function (k) { return map[k]; }).sort(function (a, b) { return (a.sub || 0) - (b.sub || 0); });
  }
  function entryHTML(e, sel) {
    var meta = [];
    if (sel.mode === 'teacher') e.groups.sort(cmpRu).forEach(function (g) { meta.push('<span class="grp">' + esc(g) + '</span>'); });
    else if (e.teacher) meta.push('<span>' + esc(e.teacher) + '</span>');
    if (e.room) meta.push('<span>' + esc(roomLabel(e.room)) + '</span>');
    if (e.sub) meta.push('<span class="tag">' + e.sub + ' подгруппа</span>');
    if (e.par) meta.push('<span class="tag tag-par">' + (e.par === 1 ? 'по чётным' : 'по нечётным') + '</span>');
    return '<div class="entry"><p class="disc">' + discHTML(e.disc) + '</p>' + (meta.length ? '<p class="meta">' + meta.join('') + '</p>' : '') + '</div>';
  }
  function dayCardHTML(dayIso, wd, rows, sel, today) {
    var diff = daysBetween(today, dayIso);
    var cls = diff === 0 ? ' is-today' : (diff < 0 ? ' is-past' : '');
    var rel = relLabel(diff);
    var head = '<header class="dhead"><span class="dname">' + WEEKDAYS_MON[wd] + '</span><span class="ddate">' + fmtDayMonth(dayIso) + '</span>' +
      (rel ? '<span class="rel">' + rel + '</span>' : '') + '</header>';
    if (!rows.length) return '<article class="dcard' + cls + '">' + head + '<p class="dempty">Пар нет</p></article>';
    var byPair = {};
    rows.forEach(function (r) { var n = slotOf(r).n; (byPair[n] || (byPair[n] = [])).push(r); });
    var nums = Object.keys(byPair).map(Number).sort(function (a, b) { return a - b; });
    var nowMin = new Date().getHours() * 60 + new Date().getMinutes();
    var items = [];
    for (var p = nums[0]; p <= nums[nums.length - 1]; p++) {
      var s = DB.slots[p - 1];
      if (!byPair[p]) {
        items.push('<li class="pair gap"><span class="pnum"><b>' + p + '</b></span><span>Окно' + (s ? ', ' + esc(s.start) + '–' + esc(s.end) : '') + '</span></li>');
        continue;
      }
      var isNow = diff === 0 && s && nowMin >= minutesOf(s.start) && nowMin < minutesOf(s.end);
      items.push('<li class="pair' + (isNow ? ' is-now' : '') + '"><span class="pnum"><b>' + p + ' пара</b><span>' + (s ? esc(s.start) + '<br>' + esc(s.end) : '') + '</span></span><div>' +
        pairEntries(byPair[p], sel).map(function (e) { return entryHTML(e, sel); }).join('') + '</div></li>');
    }
    return '<article class="dcard' + cls + '">' + head + '<ol class="pairs">' + items.join('') + '</ol></article>';
  }

  function renderNow(sel, today) {
    var el = $('now'), weeks = DB.weeks;
    if (!weeks.length || today < weeks[0].iso || today > isoOfUtc(weeks[weeks.length - 1].ms + 6 * DAY)) { el.hidden = true; return; }
    var who = sel.mode === 'group' ? 'у ' + sel.name : 'у преподавателя ' + sel.name;
    var rows = rowsForDay(today, sel);
    var lines = [];
    if (!rows.length) {
      var next = null;
      for (var k = 1; k <= 14 && !next; k++) {
        var iso = isoOfUtc(utcOfIso(today) + k * DAY), r2 = rowsForDay(iso, sel);
        if (r2.length) next = { k: k, iso: iso, rows: r2 };
      }
      lines.push('<p class="now-title">Сегодня пар нет</p>');
      if (next) {
        var nums = next.rows.map(function (r) { return slotOf(r).n; });
        var first = DB.slots[Math.min.apply(null, nums) - 1];
        var np = nums.filter(function (n, i) { return nums.indexOf(n) === i; }).length;
        lines.push('<p class="now-line">Следующие занятия ' + esc(who) + ': <strong>' + (next.k === 1 ? 'завтра' : WEEKDAYS[isoToDate(next.iso).getDay()] + ', ' + fmtDayMonth(next.iso)) + '</strong>, ' +
          count(np, 'пара', 'пары', 'пар') + (first ? ', начало в ' + esc(first.start) : '') + '.</p>');
      }
    } else {
      var ps = [];
      rows.forEach(function (r) { var n = slotOf(r).n; if (ps.indexOf(n) < 0) ps.push(n); });
      ps.sort(function (a, b) { return a - b; });
      var s0 = DB.slots[ps[0] - 1], s1 = DB.slots[ps[ps.length - 1] - 1];
      lines.push('<p class="now-title">Сегодня ' + esc(who) + ' ' + count(ps.length, 'пара', 'пары', 'пар') + (s0 && s1 ? ', ' + esc(s0.start) + '–' + esc(s1.end) : '') + '</p>');
      var nowMin = new Date().getHours() * 60 + new Date().getMinutes(), cur = null, upcoming = null;
      ps.forEach(function (p) {
        var s = DB.slots[p - 1]; if (!s) return;
        if (nowMin >= minutesOf(s.start) && nowMin < minutesOf(s.end)) cur = p;
        else if (minutesOf(s.start) > nowMin && upcoming === null) upcoming = p;
      });
      var describe = function (p) {
        return pairEntries(rows.filter(function (r) { return slotOf(r).n === p; }), sel).map(function (e) {
          return '<strong>' + esc(splitCode(e.disc).name) + '</strong>' + (e.room ? ', ' + esc(roomLabel(e.room)) : '') + (e.sub ? ' (' + e.sub + ' подгр.)' : '');
        }).join('; ');
      };
      if (cur !== null) lines.push('<p class="now-line">Сейчас идёт ' + cur + ' пара, до ' + esc(DB.slots[cur - 1].end) + ': ' + describe(cur) + '.</p>');
      if (upcoming !== null) lines.push('<p class="now-line">' + (cur !== null ? 'Дальше' : 'Следующая') + ' — ' + upcoming + ' пара в ' + esc(DB.slots[upcoming - 1].start) + ': ' + describe(upcoming) + '.</p>');
      if (cur === null && upcoming === null) lines.push('<p class="now-line">Пары на сегодня закончились.</p>');
    }
    el.innerHTML = '<span class="eyebrow">' + esc(WEEKDAYS[new Date().getDay()]) + ', ' + esc(fmtDayMonth(today)) + '</span>' + lines.join('');
    el.hidden = false;
  }

  function showEmpty(text, action, data) {
    $('empty').hidden = false;
    $('empty-text').textContent = text;
    var btn = $('empty-action');
    btn.hidden = !action;
    if (action) { btn.textContent = action; btn.setAttribute('data-action', data || ''); }
  }

  function renderStatus() {
    var el = $('status'), text, st;
    var shown = [];
    if (state.mode === 'group' && state.group && DB.sheets[state.group]) shown = [DB.sheets[state.group]];
    if (state.mode === 'teacher') shown = Object.keys(DB.sheets).map(function (g) { return DB.sheets[g]; });
    var anyFailed = net.listFailed || Object.keys(net.failed).length > 0;
    if (net.bulk) { st = 'loading'; text = 'Загружаю расписание групп из Google-таблицы: ' + net.bulk.done + ' из ' + net.bulk.total + '…'; }
    else if (net.pending > 0) { st = 'loading'; text = 'Загружаю свежее расписание из Google-таблицы…'; }
    else if (anyFailed) {
      var cached = shown.filter(function (s) { return s.source === 'cache'; });
      var hasData = state.mode === 'group' ? DB.groups.length > 0 && (!state.group || !!DB.sheets[state.group]) : shown.length > 0;
      if (hasData) {
        st = 'stale';
        var oldest = cached.length ? Math.min.apply(null, cached.map(function (s) { return s.at; })) : 0;
        text = 'Google-таблица не ответила. ' + (oldest ? 'Показана сохранённая копия, загруженная ' + fmtStamp(oldest) + '.' : 'Показана сохранённая копия.');
      } else { st = 'error'; text = 'Не удалось загрузить расписание из Google-таблицы.'; }
    } else if (net.listLive) {
      var live = shown.filter(function (s) { return s.source === 'live'; });
      var at = live.length ? Math.max.apply(null, live.map(function (s) { return s.at; })) : Date.now();
      st = 'ok'; text = 'Свежие данные из Google-таблицы, загружены ' + fmtStamp(at) + '.';
    } else { st = 'loading'; text = 'Загружаю расписание из Google-таблицы…'; }
    el.setAttribute('data-state', st);
    $('status-text').textContent = text;
    $('retry').hidden = !(st === 'stale' || st === 'error');
  }

  function render() {
    var td = new Date(), today = todayISO();
    $('today-label').textContent = WEEKDAYS[td.getDay()] + ', ' + td.getDate() + ' ' + MONTHS_GEN[td.getMonth()];
    $('stamp').textContent = DB.groups.length
      ? count(DB.groups.length, 'группа', 'группы', 'групп') + (DB.weeks.length ? ' · ' + count(DB.weeks.length, 'неделя', 'недели', 'недель') + ' с ' + fmtDayMonth(DB.weeks[0].iso) : '')
      : '';
    renderStatus();

    var sel = selection();
    $('group-field').hidden = state.mode !== 'group';
    $('teacher-field').hidden = state.mode !== 'teacher';
    $('group').value = state.group;
    $('teacher').value = state.teacher;
    Array.prototype.forEach.call($('mode').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === state.mode ? 'true' : 'false');
    });

    var weekEl = $('week');
    $('empty').hidden = true;
    $('now').hidden = true;
    $('picker').hidden = true;
    weekEl.innerHTML = '';
    var sheet = sel && sel.mode === 'group' ? DB.sheets[sel.name] : null;
    $('tpl-note').hidden = !(sheet && sheet.template);

    var weeks = DB.weeks, wi = nearestWeekIndex(state.week || currentMondayIso());
    $('weeknav').hidden = !(sel && wi >= 0);
    savePrefs();
    writeHash();

    if (!DB.groups.length) {
      if (net.listFailed) showEmpty('Не удалось загрузить расписание из Google-таблицы. Проверьте подключение к интернету. Если ошибка повторяется, проверьте, что таблица открыта по ссылке.', 'Повторить', 'retry');
      else showEmpty('Загружаю расписание из Google-таблицы…');
      return;
    }
    if (!sel) {
      if (state.mode === 'group') $('picker').hidden = false;
      else showEmpty(net.bulk ? 'Собираю список преподавателей: загружено ' + net.bulk.done + ' из ' + net.bulk.total + ' групп…' : 'Выберите преподавателя, чтобы увидеть его пары.');
      return;
    }
    if (sel.mode === 'group' && !sheet) {
      if (net.failed[sel.name]) showEmpty('Не удалось загрузить расписание группы ' + sel.name + '. Проверьте подключение к интернету.', 'Повторить', 'retry');
      else showEmpty('Загружаю расписание группы ' + sel.name + '…');
      return;
    }
    if (wi < 0) { showEmpty('В таблице нет ни одной недели с датами.'); return; }

    state.week = weeks[wi].iso;
    var monday = weeks[wi].iso, mondayMs = weeks[wi].ms, saturday = isoOfUtc(mondayMs + 5 * DAY);
    $('wk-label').textContent = fmtRange(monday, saturday);
    $('wk-sub').textContent = 'Неделя ' + (Math.round((mondayMs - weeks[0].ms) / (7 * DAY)) + 1) +
      (weeks[wi].par === 1 ? ' · чётная' : weeks[wi].par === 2 ? ' · нечётная' : '');
    $('wk-prev').disabled = wi <= 0;
    $('wk-next').disabled = wi >= weeks.length - 1;
    $('wk-today').hidden = nearestWeekIndex(currentMondayIso()) === wi;

    var total = 0, cards = [];
    for (var wd = 0; wd < 6; wd++) {
      var dayIso = isoOfUtc(mondayMs + wd * DAY), rows = rowsForDay(dayIso, sel);
      total += rows.length;
      cards.push(dayCardHTML(dayIso, wd, rows, sel, today));
    }
    if (total) weekEl.innerHTML = cards.join('');
    else {
      var other = findWeekWithLessons(sel, wi);
      showEmpty(sel.mode === 'teacher' && net.bulk
        ? 'Загружаю расписание групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…'
        : 'На эту неделю пар нет: расписание ещё не заполнено или идёт практика.',
        other >= 0 ? 'Показать ' + fmtRange(weeks[other].iso, isoOfUtc(weeks[other].ms + 5 * DAY)) : '', other >= 0 ? 'week:' + weeks[other].iso : '');
    }
    renderNow(sel, today);
  }

  function findWeekWithLessons(sel, from) {
    var best = -1, bestDist = 1e9;
    DB.weeks.forEach(function (w, i) {
      var has = false;
      for (var wd = 0; wd < 6 && !has; wd++) has = rowsForDay(isoOfUtc(w.ms + wd * DAY), sel).length > 0;
      if (has && Math.abs(i - from) < bestDist) { best = i; bestDist = Math.abs(i - from); }
    });
    return best;
  }

  // ---------- events ----------
  $('mode').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    state.mode = b.getAttribute('data-value');
    render();
    ensureView();
  });
  $('group').addEventListener('change', function (e) { state.group = e.target.value; render(); ensureView(); });
  $('teacher').addEventListener('change', function (e) { state.teacher = e.target.value; render(); });
  $('picker-body').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-group]');
    if (b) { state.group = b.getAttribute('data-group'); render(); ensureView(); }
  });
  function shiftWeek(delta) {
    var wi = nearestWeekIndex(state.week) + delta;
    if (wi >= 0 && wi < DB.weeks.length) { state.week = DB.weeks[wi].iso; render(); }
  }
  $('wk-prev').addEventListener('click', function () { shiftWeek(-1); });
  $('wk-next').addEventListener('click', function () { shiftWeek(1); });
  $('wk-today').addEventListener('click', function () { state.week = ''; render(); });
  $('empty-action').addEventListener('click', function () {
    var a = this.getAttribute('data-action') || '';
    if (a === 'retry') retry();
    else if (a.indexOf('week:') === 0) { state.week = a.slice(5); render(); }
  });
  $('retry').addEventListener('click', retry);
  window.addEventListener('hashchange', function () { applyHash(readHash()); render(); ensureView(); });

  // ---------- start ----------
  (function boot() {
    var lc = cacheGet(CFG.groupListSheet), cc = CFG.calendarSheet ? cacheGet(CFG.calendarSheet) : null;
    if (lc) applyGroupList(lc.csv);
    if (cc) applyCalendar(cc.csv);
    applyHash(readHash());
    if (state.mode === 'group' && state.group) {
      var c = cacheGet(state.group);
      if (c) applySheet(state.group, c.csv, 'cache', c.t);
    }
    recompute();
    fillOptions();
    render();
    loadLists();
    // keep "today" and the current pair fresh while the page stays open
    setInterval(render, 60000);
    // when the site is opened from this computer, expose the data for checks in the console
    if (/^(localhost|127\.0\.0\.1)$/.test(location.hostname)) window.__raspisanie = { DB: DB, net: net, ensureAll: ensureAll };
  })();
})();
