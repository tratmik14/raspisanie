(function () {
  'use strict';

  var CFG = window.SCHEDULE_CONFIG || {};
  var DB_KEY = String(CFG.sheetId || '') + ':';
  var DAY = 864e5;
  var PREFS_KEY = 'raspisanie.prefs.v1';
  var CACHE_KEY = 'raspisanie.sheet.v2:' + DB_KEY;
  var CONCURRENCY = 4;
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  var MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  var WEEKDAYS = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  var WEEKDAYS_MON = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'];
  var WEEKDAY_NAMES = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];
  var TEACHER_RE = /^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.?)?$/;
  var GROUP_RE = /^[А-ЯЁA-Z][А-ЯЁA-Zа-яё0-9]*-\d{2}[А-ЯЁа-яё]?$/;
  var TYPE_KEY = { 'Экзамен': 'exam', 'Диф. зачёт': 'diff', 'Зачёт': 'credit' };
  var SITE_VERSION = '6.1';
  var COLS_KEY = 'raspisanie.columns.v1';
  var FAVS_KEY = 'raspisanie.favs.v1';
  var THEME_KEY = 'raspisanie.theme.v1';
  var STATE_KEY = 'raspisanie.state.v2:' + DB_KEY;
  var GOOD_KEY = 'raspisanie.good.v2:' + DB_KEY;
  var DIFF_KEY = 'raspisanie.diff.v2:' + DB_KEY;
  var SOURCES_KEY = 'raspisanie.sources.v2:' + DB_KEY;
  var VIEWS = ['lessons', 'hours', 'balance', 'grafik', 'exams', 'history'];
  var VIEW_TITLES = {
    lessons: ['Расписание занятий', 'Пары по группам, преподавателям и аудиториям на каждый день и неделю.'],
    hours: ['Часы по предметам', 'Сколько часов по плану, сколько уже проведено и сколько осталось.'],
    balance: ['Вычитка часов', 'Сколько пар не хватает или лишние до конца семестра. Пересчитывается сразу, как только пару добавили, отменили или перенесли.'],
    grafik: ['График учебного процесса', 'Теория, практика, сессии и каникулы каждой группы по неделям учебного года.'],
    exams: ['Расписание экзаменов', 'Экзамены, зачёты и пересдачи по группам и преподавателям.'],
    history: ['Изменения в расписании', 'Замены и отмены пар, которые внесли администраторы, и журнал всех правок.']
  };
  // week types of the PDF schedule (tools/parse_grafik.py), in the colours of its legend
  var WEEK_TYPES = {
    T: 'теоретическое обучение', A: 'промежуточная аттестация', K: 'каникулы', U: 'учебная практика',
    P: 'производственная практика', X: 'преддипломная практика', D: 'подготовка к ГИА', G: 'государственная итоговая аттестация'
  };
  var THEMES = [
    { id: 'cosmos', name: 'Космос', note: 'тёмная, ночное небо и оранжевый «Восток»' },
    { id: 'aurora', name: 'Аврора', note: 'тёмная, северное сияние' },
    { id: 'sunset', name: 'Закат', note: 'тёмная, стратосфера на закате' },
    { id: 'neon', name: 'Неон', note: 'тёмная, ночной город и неон' },
    { id: 'dawn', name: 'Рассвет', note: 'светлая, утреннее небо' },
    { id: 'dune', name: 'Дюна', note: 'светлая, тёплый песок и терракота' }
  ];
  var Api = window.RaspApi;
  var Hours = window.RaspHours;
  var Rules = window.RaspRules;
  var PLAN_OVERRIDES = [], CATALOG = [];
  function catalogName(type, value) { return window.RaspCatalog.name(CATALOG, type, value); }
  function catalogTeacher(value) { return window.RaspCatalog.teachers(CATALOG, value); }
  function catalogDisc(value) {
    var dc = splitCode(value), renamed = catalogName('subject', dc.name);
    return renamed === dc.name ? value : (dc.code ? dc.code + ' ' : '') + renamed;
  }
  function catalogNames(type) {
    var names = CATALOG.filter(function (r) { return r.type === type; }).map(function (r) { return r.name; });
    Object.keys(DB.sheets).forEach(function (g) { DB.sheets[g].rows.forEach(function (r) { names = names.concat(type === 'teacher' ? splitTeachers(r.teacher) : [splitCode(r.disc).name]); }); });
    DB.groups.forEach(function (g) { var p = planFor(g); if (p) p.subjects.forEach(function (r) { names = names.concat(type === 'teacher' ? (String(r.teacher || '').match(/[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s*[А-ЯЁ]\.?/g) || []) : [splitCode(r.name).name]); }); });
    EX.records.forEach(function (r) { names = names.concat(type === 'teacher' ? splitTeachers(r.teacher) : [splitCode(r.discipline).name]); });
    return uniq(names.map(function (v) { return catalogName(type, v); })).sort(cmpRu);
  }
  function refreshSuggestions() {
    datalistFill('dl-disc', catalogNames('subject'));
    datalistFill('dl-teacher', catalogNames('teacher'));
    datalistFill('dl-room', DB.rooms || []);
  }
  function exDefaults() {
    return { group: [], teacher: [], disc: [], type: 'exam', status: 'all', half: 'all', from: '', to: '', search: '', ordering: '', page: 1, pageSize: 20, rpage: 1 };
  }
  var EX_FILTER_KEYS = ['group', 'teacher', 'disc', 'type', 'status', 'half', 'from', 'to', 'search'];

  function $(id) { return document.getElementById(id); }

  // ---------- small helpers ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function hm(h, m) { return pad(h) + ':' + pad(m); }
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
  // dates go through Intl.DateTimeFormat; the hand-made forms stay for browsers without it
  var DF = (function () {
    try {
      return {
        dayMonth: new Intl.DateTimeFormat('ru-RU', { day: 'numeric', month: 'long' }),
        short: new Intl.DateTimeFormat('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric' }),
        time: new Intl.DateTimeFormat('ru-RU', { hour: '2-digit', minute: '2-digit' })
      };
    } catch (e) { return null; }
  })();
  function fmtDayMonth(iso) { var d = isoToDate(iso); return DF ? DF.dayMonth.format(d) : d.getDate() + ' ' + MONTHS_GEN[d.getMonth()]; }
  function fmtShort(iso) {
    if (DF) return DF.short.format(isoToDate(iso));
    var p = isoParts(iso);
    return pad(p.d) + '.' + pad(p.m) + '.' + p.y;
  }
  function fmtRange(a, b) {
    var x = isoToDate(a), y = isoToDate(b);
    if (DF && DF.dayMonth.formatRange) return DF.dayMonth.formatRange(x, y);
    if (x.getMonth() === y.getMonth()) return x.getDate() + '–' + y.getDate() + ' ' + MONTHS_GEN[y.getMonth()];
    return x.getDate() + ' ' + MONTHS_GEN[x.getMonth()] + ' – ' + y.getDate() + ' ' + MONTHS_GEN[y.getMonth()];
  }
  function fmtStamp(ms) {
    var d = new Date(ms), t = DF ? DF.time.format(d) : pad(d.getHours()) + ':' + pad(d.getMinutes());
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
  function inDays(diff) {
    if (diff === 0) return 'сегодня';
    if (diff === 1) return 'завтра';
    if (diff === 2) return 'послезавтра';
    return 'через ' + count(diff, 'день', 'дня', 'дней');
  }
  function minutesOf(t) { var p = String(t).split(':'); return (+p[0]) * 60 + (+p[1] || 0); }
  function normTime(t) { var p = String(t).split(':'); return (+p[0]) + ':' + pad(+p[1] || 0); }
  function cmpRu(a, b) { return String(a).localeCompare(String(b), 'ru', { numeric: true }); }
  function uniq(arr) { return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; }); }
  function splitTeachers(s) { return window.RaspCatalog.splitTeachers(s); }
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
  function cellText(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

  // ---------- Google Sheets ----------
  // headers=0 keeps Google from gluing the top rows of a sheet into one header row
  // range limits a group sheet to its timetable columns: the side tables of the head of studies
  // («Пары», «Часы», the plan) stay out, so the answer is smaller and Google builds it faster.
  function sheetUrl(name, range) {
    return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(CFG.sheetId) +
      '/gviz/tq?tqx=out:csv&headers=0&sheet=' + encodeURIComponent(name) + (range ? '&range=' + encodeURIComponent(range) : '');
  }
  // Google answers 35 sheets at once with an occasional error or 429; such a sheet is asked again
  // twice with a pause, so one hiccup no longer leaves a group without its timetable.
  function fetchSheet(name, range) {
    if (CFG.dataSource === 'database') {
      if (!Api || !Api.loadTable) return Promise.reject(new Error('База расписания не подключена'));
      return Api.loadTable(name);
    }
    var tries = 0;
    function attempt() {
      return fetch(sheetUrl(name, range), { cache: 'no-store' }).then(function (res) {
        var type = res.headers.get('content-type') || '';
        if (!res.ok || type.indexOf('csv') < 0) {
          var err = new Error('Google Sheets answered ' + res.status + ' ' + type);
          err.final = res.ok || (res.status >= 400 && res.status < 500 && res.status !== 429);
          throw err;
        }
        return res.text();
      }).catch(function (err) {
        tries++;
        if (err.final || tries >= 3) throw err;
        return new Promise(function (r) { setTimeout(r, 700 * tries * tries + Math.random() * 400); }).then(attempt);
      });
    }
    return attempt();
  }
  function groupSheetRange() { return CFG.groupSheetRange == null ? 'A:I' : String(CFG.groupSheetRange); }
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
  function serialToUtc(n) { return Date.UTC(1899, 11, 30) + Math.floor(n) * DAY; }
  function cellDayUtc(v) {
    var s = String(v).trim();
    var m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(s);
    if (m) return validDate(+m[3], +m[2], +m[1]) ? Date.UTC(+m[3], +m[2] - 1, +m[1]) : null;
    if (/^\d{5}(\.0+)?$/.test(s)) {
      var n = parseFloat(s); // a date left as an Excel day number
      if (n > 30000 && n < 80000) return serialToUtc(n);
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
  // «Неделя 12» in one of the first columns
  function weekLabelOf(row) {
    for (var c = 0; c < 3 && c < row.length; c++) {
      var m = /^\s*неделя\s*(\d{1,2})\s*$/i.exec(String(row[c] == null ? '' : row[c]));
      if (m) return +m[1];
    }
    return 0;
  }
  // Splits a sheet into week blocks: a row of dates, a row of weekday names, then one row per pair.
  // Real sheets are not always tidy, and nothing that is on the sheet should get lost:
  // - a weekday row typed twice in a row (ПКп-24) is one week, not an empty week and a lost one;
  // - a week without its row of dates (ТМ-26, week 3; ОДЛу-24, weeks 19–36) gets its Monday from its
  //   «Неделя N» label, else it follows the week before. Such weeks are counted in `guessed`.
  function scanLessonSheet(rows) {
    var blocks = [], cur = null, dayCols = null, timeCol = -1, pendingMonday = null, dateFresh = false, hasDates = false;
    var label = 0, ref = null, lastMonday = null, guessed = 0;
    rows.forEach(function (raw) {
      var row = raw.slice(0, 9);
      var wds = [], days = [], pw = 0;
      row.forEach(function (v, c) {
        var w = weekdayIndex(v); if (w >= 0) wds.push([c, w]);
        var d = cellDayUtc(v); if (d != null) days.push(d);
        if (c < 2 && !pw) pw = parityOf(v);
      });
      var wl = weekLabelOf(row);
      if (wl) label = wl;
      if (days.length >= 3) {
        pendingMonday = modeMonday(days); dateFresh = true; hasDates = true;
        if (wl && !ref) ref = { n: wl, ms: pendingMonday };
        return;
      }
      if (wds.length >= 3) {
        var cols = {};
        wds.forEach(function (x) { if (x[1] < 6) cols[x[0]] = x[1]; });
        if (cur && !cur.cells.length && !dateFresh && !label) {
          dayCols = cols;
          timeCol = wds[0][0] - wds[0][1] - 1;
          if (pw) cur.votes[pw]++;
          return;
        }
        dayCols = cols;
        timeCol = wds[0][0] - wds[0][1] - 1;
        var monday = dateFresh ? pendingMonday : null, guess = false;
        if (monday == null && hasDates) {
          if (label && ref) monday = ref.ms + (label - ref.n) * 7 * DAY;
          else if (lastMonday != null) monday = lastMonday + 7 * DAY;
          guess = monday != null;
          if (guess) guessed++;
        }
        cur = { monday: monday, votes: [0, 0, 0], cells: [], guessed: guess };
        if (monday != null) lastMonday = monday;
        blocks.push(cur);
        dateFresh = false;
        label = 0;
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
    return { blocks: blocks, hasDates: hasDates, guessed: guessed };
  }

  // ---------- reading the exam sheet (ТЗ БП 5.5: columns A–H) ----------
  function monthFromWord(w) {
    var stems = [['янв', 1], ['фев', 2], ['мар', 3], ['апр', 4], ['мая', 5], ['май', 5], ['июн', 6], ['июл', 7], ['авг', 8], ['сен', 9], ['окт', 10], ['ноя', 11], ['дек', 12]];
    w = w.toLowerCase();
    for (var i = 0; i < stems.length; i++) if (w.indexOf(stems[i][0]) === 0) return stems[i][1];
    return 0;
  }
  // Returns null for an empty cell, {date, time} when it reads, {error} otherwise.
  function parseExamDate(v) {
    var s = cellText(v).replace(/ё/gi, 'е');
    if (!s) return null;
    var y, mo, d, H = null, M = null, m;
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      var n = parseFloat(s), iso = isoOfUtc(serialToUtc(n)), mins = Math.round((n - Math.floor(n)) * 1440) % 1440;
      return { date: iso, time: mins ? hm(Math.floor(mins / 60), mins % 60) : '' };
    }
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::\d{2})?)?/.exec(s))) {
      y = +m[1]; mo = +m[2]; d = +m[3];
      if (m[4] != null) { H = +m[4]; M = +m[5]; }
    } else if ((m = /^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s*(?:г\.?)?,?\s*(?:в\s*)?(\d{1,2})[:.](\d{2}))?/i.exec(s))) {
      d = +m[1]; mo = +m[2]; y = +m[3]; if (y < 100) y += 2000;
      // a workbook saved with English settings writes the month first: «12/23/2026»
      if (mo > 12 && d <= 12 && /\//.test(m[0])) { var sw = d; d = mo; mo = sw; }
      if (m[4] != null && /\bpm\b/i.test(s) && +m[4] < 12) m[4] = String(+m[4] + 12);
      if (m[4] != null) { H = +m[4]; M = +m[5]; }
    } else if ((m = /^(\d{1,2})\s+([а-я]+)\.?\s+(\d{4})/i.exec(s))) {
      d = +m[1]; mo = monthFromWord(m[2]); y = +m[3];
      var t = /(\d{1,2})\s*(?::|час[а-я]*|ч\.?)\s*(\d{1,2})?/i.exec(s.slice(m[0].length));
      if (t) { H = +t[1]; M = t[2] != null ? +t[2] : 0; }
    } else return { error: true };
    if (!mo || !validDate(y, mo, d) || (H != null && (H > 23 || M > 59))) return { error: true };
    return { date: ymd(y, mo, d), time: H != null ? hm(H, M) : '' };
  }
  function parseTimeText(v) {
    var t = /(\d{1,2})\s*(?::|\.|час[а-я]*|ч\.?)\s*(\d{1,2})?/i.exec(cellText(v));
    return t && +t[1] < 24 && (t[2] == null || +t[2] < 60) ? hm(+t[1], t[2] != null ? +t[2] : 0) : '';
  }
  function normType(v) {
    var s = cellText(v).toLowerCase().replace(/ё/g, 'е');
    if (!s) return '';
    if (/диф|с оценкой|^дз$/.test(s)) return 'Диф. зачёт';
    if (/зач/.test(s)) return 'Зачёт';
    if (/экз|^э$/.test(s)) return 'Экзамен';
    return null;
  }
  function mapExamHeader(row) {
    var map = {};
    row.forEach(function (cell, i) {
      var h = cellText(cell).toLowerCase().replace(/ё/g, 'е');
      if (!h) return;
      if (/пересд/.test(h)) {
        if (/время/.test(h) && !/дата/.test(h)) { if (map.retakeTime == null) map.retakeTime = i; }
        else if (map.retake == null) map.retake = i;
        return;
      }
      if (/групп/.test(h)) { if (map.group == null) map.group = i; return; }
      if (/дисципл|модул|предмет|наименован/.test(h)) { if (map.discipline == null) map.discipline = i; return; }
      if (/ауд|кабин|помещ/.test(h)) { if (map.room == null) map.room = i; return; }
      if (/^№\s*(п\s*\/\s*п)?\.?$/.test(h)) { if (map.num == null) map.num = i; return; }
      if (/препод|фио|экзаменатор|педагог/.test(h)) { if (map.teacher == null) map.teacher = i; return; }
      if (/тип|вид|форм/.test(h)) { if (map.type == null) map.type = i; return; }
      if (/дата|срок/.test(h)) { if (map.date == null) map.date = i; return; }
      if (/время|начал/.test(h)) { if (map.time == null) map.time = i; return; }
    });
    return map;
  }
  // «график экзаменационной сессии» → Экзамен, «расписание зачетной недели» → Зачёт
  function examKindOf(text) {
    var low = cellText(text).toLowerCase().replace(/ё/g, 'е');
    if (/зачетн\S*\s+недел/.test(low)) return 'Зачёт';
    if (/экзаменацион/.test(low)) return 'Экзамен';
    return '';
  }
  // a word broken at a dash by a narrow cell («дизайн-» / «макетов») is joined back
  function joinDisc(s) { return cellText(s).replace(/([А-ЯЁа-яёA-Za-z])- (?=[а-яё])/g, '$1-'); }
  // Returns null when the sheet does not look like an exam list (Google answers with the
  // first sheet of the workbook when the requested one does not exist).
  // The college's own Word schedules are read too: one table per group under «группы № ПСОп-23»,
  // no «Тип» column (the kind comes from the title of the document, «диф» in a name makes it a
  // diff. credit), and a long name that runs over into a row without a date and a number.
  // defaults: { group, kind, map } — what the text above the table said, and the columns of the
  // table before it (a table that breaks over a page has no header the second time).
  function parseExamSheet(rows, defaults) {
    defaults = defaults || {};
    var headerIdx = -1, map = null;
    for (var i = 0; i < Math.min(rows.length, 20); i++) {
      var mp = mapExamHeader(rows[i] || []);
      var hits = ['group', 'discipline', 'date', 'teacher', 'type'].filter(function (k) { return mp[k] != null; }).length;
      var more = ['group', 'teacher', 'type', 'time', 'room', 'num'].filter(function (k) { return mp[k] != null; }).length;
      if (mp.discipline != null && mp.date != null && (hits >= 3 || more >= 1)) { headerIdx = i; map = mp; break; }
    }
    if (!map && defaults.map) map = defaults.map;
    if (!map) return null;
    var above = rows.slice(0, Math.max(headerIdx, 0)).map(function (r) { return (r || []).join(' '); }).join(' ');
    var dGroup = groupFromText(above) || defaults.group || '', dKind = examKindOf(above) || defaults.kind || '';
    var out = { records: [], skipped: 0, bad: [], map: map, header: headerIdx >= 0 }, seen = {}, last = null;
    function get(row, key) { return map[key] == null ? '' : row[map[key]]; }
    for (var r = headerIdx + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var discipline = joinDisc(get(row, 'discipline')), teacherCell = cellText(get(row, 'teacher'));
      var group = cellText(get(row, 'group')) || dGroup, groupGuess = false;
      if (!group && xiGroup) { group = xiGroup; groupGuess = true; }
      var dateRaw = get(row, 'date'), num = cellText(get(row, 'num'));
      if (!discipline && !cellText(dateRaw) && !teacherCell) continue;
      if (last && !cellText(dateRaw) && !/^\d/.test(num) && (!cellText(get(row, 'group')) || cellText(get(row, 'group')) === last.group)) {
        if (discipline) last.discipline = joinDisc(last.discipline + ' ' + discipline);
        if (teacherCell) last.teacher = [last.teacher, teacherCell].filter(Boolean).join(', ');
        continue;
      }
      if (!discipline && !cellText(dateRaw)) continue;
      var dt = parseExamDate(dateRaw), type = normType(get(row, 'type'));
      var typeGuess = false;
      if (type === '') type = /(^|[^а-яё])диф/i.test(discipline) ? 'Диф. зачёт' : dKind;
      if (!type && type !== null && discipline && dt && !dt.error) { type = 'Экзамен'; typeGuess = true; }
      var why = !group ? 'не указана группа' : !discipline ? 'нет дисциплины' : !dt ? 'нет даты' : dt.error ? 'не читается дата «' + cellText(dateRaw) + '»' : !type ? 'не понятно, экзамен это или зачёт' : '';
      if (why) { out.skipped++; out.bad.push({ text: [group, discipline, cellText(dateRaw)].filter(Boolean).join(' · '), why: why }); last = null; continue; }
      var time = dt.time || (map.time != null ? parseTimeText(get(row, 'time')) : '');
      var rt = parseExamDate(get(row, 'retake')), retakeDate = '', retakeTime = '';
      if (rt && !rt.error) {
        retakeDate = rt.date;
        retakeTime = rt.time || (map.retakeTime != null ? parseTimeText(get(row, 'retakeTime')) : '');
        if ((retakeDate + retakeTime) <= (dt.date + time)) { retakeDate = ''; retakeTime = ''; }
      }
      var key = [group, discipline, dt.date, time].join('|').toLowerCase();
      if (seen[key]) { last = null; continue; }
      seen[key] = true;
      last = {
        group: group, discipline: discipline, date: dt.date, time: time, room: cellText(get(row, 'room')),
        teacher: teacherCell, type: type, retakeDate: retakeDate, retakeTime: retakeTime, typeGuess: typeGuess, groupGuess: groupGuess
      };
      out.records.push(last);
    }
    return out;
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
  // exam list: status 'loading' | 'ok' | 'none' (no such sheet) | 'error'.
  // EX joins the sheet «Экзамены» (EXS), the college's Word schedules parsed by tools/parse_exams.py
  // into data/exams.json (EXF) and the list an admin loaded from a document (EXI).
  var EXS = { status: 'loading', records: [], skipped: 0, source: '', at: 0 };
  var EXF = { status: 'loading', records: [], period: '' };
  var EXI = { records: [], at: 0, raw: [] };
  var EX = { status: 'loading', records: [], skipped: 0, source: '', at: 0, imported: 0 };
  // guards: sheet name → why its fresh copy was not taken (ТЗ, этап 2: «данные не затираются»)
  // notes: sheets taken although they shrank a lot; synced: when the open sheets were last re-read
  var net = { pending: 0, failed: {}, listLive: false, listFailed: false, bulk: null, guards: {}, notes: {}, refreshing: false, synced: Date.now(), syncing: false };
  // admin edits laid over the timetable: status 'loading' | 'ok' | 'cache' | 'error' | 'off'
  var EDITS = { status: Api.mode === 'off' ? 'off' : 'loading', list: [], at: 0 };
  // planned hours from the «0 Учебные графики» sheet: status 'loading' | 'ok' | 'error'
  var PLAN = { status: 'loading', groups: {}, order: [], source: '', at: 0 };
  // the PDF schedule parsed into data/grafik.json
  var GRAFIK = { status: 'loading', data: null };
  var HIST = { status: 'idle', items: [], full: false, group: '' };
  // the PDF schedule corrected by admins: PDF group key → week Monday → { days: 'TTUU.', note }
  var GOV = { list: [], map: {} };
  // the shared journal of imports: what changed in the college table, parser warnings for admins
  var TJ = { status: 'idle', items: [], full: false };
  var histType = 'all'; // the record-type filter of the changes view
  var showSrc = false;  // show the type, source and change time under every lesson
  try { showSrc = localStorage.getItem(SOURCES_KEY) === '1'; } catch (e) { /* storage unavailable */ }

  function groupListOf(csv) {
    var seen = {}, groups = [];
    parseCSV(csv).forEach(function (row) {
      var g = String(row[0] || '').trim();
      if (GROUP_RE.test(g) && !seen[g]) { seen[g] = true; groups.push(g); }
    });
    return groups;
  }
  function applyGroupList(csv) { DB.groups = groupListOf(csv); }
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
  function applyExams(csv, source, at) {
    var res = parseExamSheet(parseCSV(csv));
    EXS = res ? { status: 'ok', records: res.records, skipped: res.skipped, source: source, at: at } : { status: 'none', records: [], skipped: 0, source: source, at: at };
    rebuildExams();
  }
  function rebuildExams() {
    var seen = {}, records = [], byGroup = {};
    // a table an admin uploaded wins over the sheet and data/exams.json for its groups in that session
    EXI.records.forEach(function (r) { (byGroup[r.group] || (byGroup[r.group] = [])).push(utcOfIso(r.date)); });
    function overridden(r) {
      var ds = byGroup[r.group], t = utcOfIso(r.date);
      return !!ds && ds.some(function (d) { return Math.abs(d - t) <= 30 * DAY; });
    }
    EXS.records.concat(EXF.records).filter(function (r) { return !overridden(r); }).concat(EXI.records).forEach(function (r) {
      r = Object.assign({}, r, { discipline: catalogDisc(r.discipline), teacher: catalogTeacher(r.teacher) });
      var k = [r.group, r.discipline, r.date, r.time].join('|').toLowerCase();
      if (!seen[k]) { seen[k] = true; records.push(r); }
    });
    var extra = EXF.records.length + EXI.records.length > 0;
    EX = {
      status: CFG.dataSource === 'database' || extra ? 'ok' : EXS.status === 'loading' || EXF.status === 'loading' ? 'loading' : EXS.status, records: records, skipped: EXS.skipped,
      source: EXS.source || (extra ? 'live' : ''), at: EXS.at || EXI.at, imported: EXI.records.length
    };
  }
  function applyPlan(csv, source, at) {
    var res = Hours.parsePlan(parseCSV(csv));
    PLAN = { status: res.order.length ? 'ok' : 'error', groups: res.groups, order: res.order, source: source, at: at };
  }
  function setEdits(list, status) {
    EDITS = { status: status, list: list.filter(function (e) { return e && e.status === 'active'; }), at: Date.now() };
    if (status === 'ok') persistState();
  }
  // edits, schedule corrections and imported exams from the server (state), kept for offline visits
  function setState(st, status) {
    CATALOG = st.catalog || [];
    PLAN_OVERRIDES = st.plans || [];
    GOV = { list: st.grafik || [], map: {} };
    GOV.list.forEach(function (g) { (GOV.map[g.group] || (GOV.map[g.group] = {}))[g.week] = g; });
    EXI = {
      raw: st.exams || [], at: Date.now(),
      records: (st.exams || []).map(function (x) {
        return { group: x.group, discipline: x.discipline, date: x.date, time: x.time, room: x.room, teacher: x.teacher, type: x.type, retakeDate: x.retakeDate, retakeTime: x.retakeTime, imported: true };
      })
    };
    rebuildExams();
    setEdits(st.edits || [], status);
  }
  function persistState() {
    try { localStorage.setItem(STATE_KEY, JSON.stringify({ t: Date.now(), state: { edits: EDITS.list, grafik: GOV.list, exams: EXI.raw, plans: PLAN_OVERRIDES, catalog: CATALOG } })); }
    catch (e) { /* storage unavailable */ }
  }
  function setGrafik(items) {
    items.forEach(function (it) {
      GOV.list = GOV.list.filter(function (g) { return !(g.group === it.group && g.week === it.week); });
      if (it.days) GOV.list.push(it);
    });
    GOV.map = {};
    GOV.list.forEach(function (g) { (GOV.map[g.group] || (GOV.map[g.group] = {}))[g.week] = g; });
    persistState();
  }

  // ---------- one name, several spellings ----------
  // The plan and the PDF sometimes name a group differently from the timetable (ГДп-24 is «ГД-24»
  // in the PDF, ПДп-26 is «ПДп-22»): same letters and year first, then the letters without «п».
  function lettersOf(g) { return String(g).replace(/\s+/g, '').replace(/-.*$/, ''); }
  function yearOf(g) { var m = /-(\d{2})/.exec(g); return m ? m[1] : ''; }
  function resolveKey(group, keys, courseOfKey) {
    if (!group) return null;
    if (keys.indexOf(group) >= 0) return group;
    var flat = group.replace(/\s+/g, '');
    var hit = keys.filter(function (k) { return k.replace(/\s+/g, '') === flat; })[0];
    if (hit) return hit;
    var L = lettersOf(group), Y = yearOf(group), c = courseOf(group);
    hit = keys.filter(function (k) { return lettersOf(k) === L && courseOfKey && courseOfKey(k) === c && c; })[0];
    if (hit) return hit;
    var base = L.replace(/п$/, '');
    return keys.filter(function (k) { return lettersOf(k).replace(/п$/, '') === base && yearOf(k) === Y; })[0] || null;
  }
  function planFor(group) {
    var key = resolveKey(group, PLAN.order, function (k) { return PLAN.groups[k].course; });
    var base = key ? PLAN.groups[key] : null;
    var overrides = PLAN_OVERRIDES.filter(function (r) { return r.group === group; });
    if (!base && !overrides.length) return null;
    var subjects = base ? base.subjects.map(function (x) { return Object.assign({}, x); }) : [];
    overrides.forEach(function (r) {
      var i = subjects.findIndex(function (x) { return catalogDisc(x.name) === catalogDisc(r.name); });
      var next = Object.assign({}, i >= 0 ? subjects[i] : {}, {
        name: r.name, code: Hours.codeOf(r.name).code, s1: +r.s1, s2: +r.s2,
        year: +r.s1 + +r.s2, teacher: r.teacher, form: [r.f1 || '', r.f2 || '']
      });
      if (i >= 0) subjects[i] = next; else subjects.push(next);
    });
    return Object.assign({}, base || { group: group, title: '', course: courseOf(group) }, { subjects: subjects.map(function (r) { return Object.assign({}, r, { name: catalogDisc(r.name), teacher: catalogTeacher(r.teacher) }); }) });
  }
  // The PDF schedule of a group with the admins' corrections laid over it («.» keeps the PDF day).
  function grafikFor(group) {
    var d = GRAFIK.data;
    if (!d) return null;
    var keys = Object.keys(d.groups);
    var key = keys.indexOf(group) >= 0 ? group : resolveKey(group, keys, function (k) { return d.groups[k].course; });
    if (!key) return null;
    var g = d.groups[key], ov = GOV.map[key], edited = {};
    if (!ov) return { key: key, g: g, edited: edited, orig: g };
    var days = g.days.map(function (str, i) {
      var o = ov[d.weeks[i]];
      if (!o) return str;
      edited[i] = o;
      return str.split('').map(function (c, j) { var x = o.days.charAt(j); return x && x !== '.' ? x : c; }).join('');
    });
    return { key: key, g: Object.assign({}, g, { days: days }), edited: edited, orig: g };
  }
  // first Monday of the spring half: from the PDF schedule, else the Monday after 8 January
  function half2Of(group) {
    var gf = grafikFor(group);
    var iso = gf ? Hours.secondHalfStart(gf.g, GRAFIK.data.weeks) : null;
    if (iso) return iso;
    var t = new Date(), y = t.getMonth() >= 6 ? t.getFullYear() + 1 : t.getFullYear();
    var ms = Date.UTC(y, 0, 8);
    while (new Date(ms).getUTCDay() !== 1) ms += DAY;
    return isoOfUtc(ms);
  }

  // ---------- admin edits over the timetable ----------
  function slotEndOf(start) {
    var s = (CFG.slots || []).filter(function (x) { return normTime(x[0]) === start; })[0];
    if (s) return normTime(s[1]);
    var m = minutesOf(start) + 90;
    return Math.floor(m / 60) + ':' + pad(m % 60);
  }
  // A replacement or cancellation owns one slot of one group (and subgroup): the lessons there give way.
  // An added pair («add») stands next to them, so two pairs can share a time. Removed lessons stay as
  // hidden "ghost" rows, so the page can show what was there before.
  function wasKey(list) {
    return list.map(function (r) { return [r.disc, r.teacher, r.room, +r.sub || 0].join('|'); })
      .filter(function (k) { return k.charAt(0) !== '|'; }).sort().join('\n');
  }
  function savedWas(e) { try { var w = JSON.parse(e.was || 'null'); return Array.isArray(w) ? w : null; } catch (err) { return null; } }
  // An edit whose result is already in the table (the head of studies typed the same change into the
  // sheet) is not laid over it again: an added pair would stand twice and count twice in the hours.
  // Such edits are listed in «Изменения» → «Уже внесено в таблицу», where they can be removed at once.
  var IN_TABLE = {};
  function lessonKey(r) { return [bare(r.disc), bare(r.teacher)].join('|'); }
  function applyEdits(group, rows) {
    var list = EDITS.list.filter(function (e) { return e.group === group; });
    if (!list.length) return rows;
    var out = rows.slice();
    list.slice().sort(function (a, b) { return (a.kind === 'add' ? 1 : 0) - (b.kind === 'add' ? 1 : 0); }).forEach(function (e) {
      var start = normTime(e.start), was = [], sub = +e.sub || 0;
      var here = out.filter(function (r) { return !r.edit && r.day === e.date && r.start === start && (!sub || !r.sub || r.sub === sub); });
      if (e.kind === 'add' || e.kind === 'replace') {
        var want = lessonKey(e);
        if (here.some(function (r) { return lessonKey(r) === want; }) && (e.kind === 'add' || here.every(function (r) { return lessonKey(r) === want; }))) { IN_TABLE[e.id] = true; return; }
      } else if (e.kind === 'cancel' && !here.length) {
        var then = savedWas(e);
        if (then && then.length) IN_TABLE[e.id] = true;
      }
      if (e.kind !== 'add') {
        out = out.filter(function (r) {
          var hit = !r.edit && r.day === e.date && r.start === start && (!sub || r.sub === sub);
          if (hit) was.push(r);
          return !hit;
        });
      }
      // ТЗ, этап 4: the table changed under the edit, so the lessons there now differ from those replaced
      var then = e.kind === 'add' ? null : savedWas(e);
      var conflict = then && wasKey(then) !== wasKey(was) ? { then: then, now: was.slice() } : null;
      var end = e.end ? normTime(e.end) : (was[0] ? was[0].end : slotEndOf(start));
      if (e.kind === 'cancel') {
        if (!was.length) out.push({ group: group, day: e.date, start: start, end: end, disc: '', teacher: '', room: '', sub: sub, par: 0, ghost: true, cancelled: true, edit: e, was: [], conflict: conflict });
        was.forEach(function (r) { out.push(Object.assign({}, r, { ghost: true, cancelled: true, edit: e, was: [r], conflict: conflict })); });
        return;
      }
      out.push({ group: group, day: e.date, start: start, end: end, disc: e.disc, teacher: e.teacher, room: e.room, sub: sub, par: 0, edit: e, was: was, conflict: conflict });
      // the teacher who lost the pair still sees it, marked as handed over
      was.forEach(function (r) {
        if (r.teacher && r.teacher !== e.teacher) out.push(Object.assign({}, r, { ghost: true, away: true, edit: e, was: [r], to: e }));
      });
    });
    return out;
  }

  function recompute() {
    var set = {}, votes = {};
    Object.keys(DB.calendar).forEach(function (m) { set[m] = true; });
    if (GRAFIK.data) GRAFIK.data.weeks.forEach(function (w) { set[utcOfIso(w)] = true; });
    EDITS.list.forEach(function (e) { set[mondayOfUtc(utcOfIso(e.date))] = true; });
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

    IN_TABLE = {};
    balMemo = {};
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
    DB.teachers = catalogNames('teacher');
    DB.rooms = uniq([].concat.apply([], Object.keys(DB.sheets).map(function (g) { return DB.sheets[g].rows.map(function (r) { return r.room; }); }))).sort(cmpRu);
    if (state) { state.teacher = catalogTeacher(state.teacher); state.exTeacher = catalogTeacher(state.exTeacher); }
    rebuildExams();
  }

  // the lessons of one sheet, week by week, before any admin edit
  function baseRows(group, scan) {
    var rows = [], seen = {};
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
    if (scan.hasDates) {
      scan.blocks.forEach(function (b) {
        if (b.monday == null || seen[b.monday]) return;
        seen[b.monday] = true;
        emit(b.monday, b.cells, false);
      });
    } else {
      // a weekly pattern without dates runs only on the theory days of the PDF schedule: not in
      // practice, session or holiday weeks, and not on public holidays
      var cells = [].concat.apply([], scan.blocks.map(function (b) { return b.cells; }));
      var gf = grafikFor(group);
      DB.weeks.forEach(function (w) {
        var wi = gf ? GRAFIK.data.weeks.indexOf(w.iso) : -1, letters = wi >= 0 ? gf.g.days[wi] : '';
        emit(w.ms, cells.filter(function (cell) {
          var iso = isoOfUtc(w.ms + cell.wd * DAY), c = letters ? letters.charAt(Math.min(cell.wd, 4)) : '';
          return !isHoliday(iso) && (!c || c === 'T');
        }), true);
      });
    }
    return rows;
  }
  function buildRows(group) {
    var sheet = DB.sheets[group], rows = baseRows(group, sheet.scan);
    sheet.base = rows;
    sheet.template = !sheet.scan.hasDates && rows.length > 0;
    // this browser's own «что изменилось»: the copy from the last visit against the fresh one
    if (sheet.prev) {
      var changes = diffRows(baseRows(group, scanLessonSheet(parseCSV(sheet.prev.csv))), rows, group);
      if (changes.length) {
        try { localStorage.setItem(DIFF_KEY + group, JSON.stringify({ at: Date.now(), since: sheet.prev.t, changes: changes.slice(0, 300) })); } catch (e) { /* storage unavailable */ }
      }
      delete sheet.prev;
    }
    sheet.rows = applyEdits(group, rows).map(function (r) { return Object.assign({}, r, { disc: catalogDisc(r.disc), teacher: catalogTeacher(r.teacher) }); });
    sheet.byDay = {};
    sheet.rows.forEach(function (r) { (sheet.byDay[r.day] || (sheet.byDay[r.day] = [])).push(r); });
  }

  // ---------- state and the address bar (ТЗ БП 1.5: filters live in query parameters) ----------
  // span: 'week' | 'today' | 'tomorrow' | 'day' (a date from the calendar, in state.day)
  // exPage: 'list' | 'book' (зачётка) | 'teacher'
  var state = { view: 'lessons', mode: 'group', group: '', teacher: '', room: '', week: '', half: null, span: 'week', day: '', exPage: 'list', bookGroup: '', bookHalf: 0, exTeacher: '', balAll: false };
  var ex = exDefaults();
  var pendingHash = ''; // an old #group link that waits for the group list
  var wantLogin = false; // #admin opens the login right away
  try {
    var saved = JSON.parse(localStorage.getItem(PREFS_KEY) || 'null');
    if (saved && typeof saved === 'object') {
      ['mode', 'group', 'teacher', 'room'].forEach(function (k) { if (typeof saved[k] === 'string') state[k] = saved[k]; });
      if (saved.span === 'today' || saved.span === 'tomorrow') state.span = saved.span;
    }
  } catch (e) { /* storage unavailable */ }
  if (['teacher', 'room'].indexOf(state.mode) < 0) state.mode = 'group';
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify({ mode: state.mode, group: state.group, teacher: state.teacher, room: state.room, span: state.span === 'day' ? 'week' : state.span })); }
    catch (e) { /* storage unavailable */ }
  }
  // Several values of one filter are repeated parameters (?group=ИС2-25&group=ГД-26):
  // discipline names can contain commas, so a comma-joined list would break them.
  function listParam(p, k) {
    var out = [];
    p.getAll(k).forEach(function (v) { v = v.trim(); if (v && out.indexOf(v) < 0) out.push(v); });
    return out;
  }
  function parseExParams(p) {
    var e = exDefaults();
    e.group = listParam(p, 'group');
    e.teacher = listParam(p, 'teacher');
    e.disc = listParam(p, 'discipline');
    e.type = /^(exam|credit|diff|retake)$/.test(p.get('type')) ? p.get('type') : 'exam';
    e.status = /^(future|past)$/.test(p.get('status')) ? p.get('status') : 'all';
    e.half = /^[12]$/.test(p.get('half')) ? p.get('half') : 'all';
    e.from = /^\d{4}-\d{2}-\d{2}$/.test(p.get('date_from') || '') ? p.get('date_from') : '';
    e.to = /^\d{4}-\d{2}-\d{2}$/.test(p.get('date_to') || '') ? p.get('date_to') : '';
    e.search = p.get('search') || '';
    e.ordering = p.get('ordering') || '';
    e.page = Math.max(1, +p.get('page') || 1);
    e.rpage = Math.max(1, +p.get('rpage') || 1);
    e.pageSize = [10, 20, 50, 100].indexOf(+p.get('page_size')) >= 0 ? +p.get('page_size') : 20;
    return e;
  }
  function exParams(e, p, withPaging) {
    e.group.forEach(function (v) { p.append('group', v); });
    e.teacher.forEach(function (v) { p.append('teacher', v); });
    e.disc.forEach(function (v) { p.append('discipline', v); });
    if (e.type !== 'exam') p.set('type', e.type);
    if (e.status !== 'all') p.set('status', e.status);
    if (e.half !== 'all') p.set('half', e.half);
    if (e.from) p.set('date_from', e.from);
    if (e.to) p.set('date_to', e.to);
    if (e.search) p.set('search', e.search);
    if (e.ordering) p.set('ordering', e.ordering);
    if (withPaging) {
      if (e.page > 1) p.set('page', e.page);
      if (e.rpage > 1) p.set('rpage', e.rpage);
      if (e.pageSize !== 20) p.set('page_size', e.pageSize);
    }
    return p;
  }
  function readUrl() {
    var p = new URLSearchParams(location.search);
    var v = p.get('view');
    if (v === 'exams') {
      state.view = 'exams';
      ex = parseExParams(p);
      state.exPage = p.get('page') === 'book' || p.get('page') === 'teacher' ? p.get('page') : 'list';
      if (state.exPage === 'book') { state.bookGroup = p.get('for') || ''; state.bookHalf = /^[12]$/.test(p.get('sem') || '') ? +p.get('sem') : 0; }
      if (state.exPage === 'teacher') state.exTeacher = p.get('for') || '';
    } else {
      state.view = VIEWS.indexOf(v) >= 0 ? v : 'lessons';
      if (p.get('room')) { state.mode = 'room'; state.room = p.get('room'); }
      else if (p.get('teacher')) { state.mode = 'teacher'; state.teacher = p.get('teacher'); }
      else if (p.get('group')) { state.mode = 'group'; state.group = p.get('group'); }
      // a link to a week opens that week, whatever the visitor chose last time
      if (/^\d{4}-\d{2}-\d{2}$/.test(p.get('week') || '')) { state.week = p.get('week'); state.span = 'week'; }
      if (/^[012]$/.test(p.get('half') || '')) state.half = +p.get('half');
      if (p.get('span') === 'today' || p.get('span') === 'tomorrow' || p.get('span') === 'week') state.span = p.get('span');
      if (/^\d{4}-\d{2}-\d{2}$/.test(p.get('day') || '')) { state.span = 'day'; state.day = p.get('day'); }
      if (state.view === 'balance') { state.balAll = p.get('all') === '1'; if (p.get('group')) state.mode = 'group'; }
    }
    if (p.get('embed') === '1') document.documentElement.classList.add('embed');
    var h = '';
    try { h = decodeURIComponent(location.hash.replace(/^#/, '')).toLowerCase(); } catch (e) { h = ''; }
    if (h === 'admin') wantLogin = true;
    else if (h) pendingHash = h;
  }
  function resolvePendingHash() {
    var h = pendingHash;
    if (!h) return;
    if (h.indexOf('p-') === 0) {
      var t = DB.teachers.filter(function (x) { return 'p-' + slug(x) === h; })[0];
      state.view = 'lessons'; state.mode = 'teacher';
      if (t) { state.teacher = t; pendingHash = ''; }
      else if (!net.bulk && allLive()) pendingHash = '';
      return;
    }
    var g = DB.groups.filter(function (x) { return slug(x) === h || x.toLowerCase() === h; })[0];
    if (g) { state.view = 'lessons'; state.mode = 'group'; state.group = g; pendingHash = ''; }
    else if (net.listLive || net.listFailed) pendingHash = '';
  }
  function writeUrl() {
    if (pendingHash) return;
    var p = new URLSearchParams();
    if (state.view === 'exams') {
      p.set('view', 'exams');
      if (state.exPage === 'book') { p.set('page', 'book'); if (state.bookGroup) p.set('for', state.bookGroup); if (state.bookHalf) p.set('sem', state.bookHalf); }
      else if (state.exPage === 'teacher') { p.set('page', 'teacher'); if (state.exTeacher) p.set('for', state.exTeacher); }
      else exParams(ex, p, true);
    } else {
      if (state.view !== 'lessons') p.set('view', state.view);
      if (state.view !== 'grafik' && state.view !== 'history') {
        if (state.mode === 'room' && state.room) p.set('room', state.room);
        else if (state.mode === 'teacher' && state.teacher) p.set('teacher', state.teacher);
        else if (state.mode === 'group' && state.group) p.set('group', state.group);
      }
      if (state.view === 'lessons') {
        if (state.span === 'day' && state.day) p.set('day', state.day);
        else if (state.span === 'today' || state.span === 'tomorrow') p.set('span', state.span);
        else if (state.week && state.week !== currentWeekIso()) p.set('week', state.week);
      }
      if ((state.view === 'hours' || state.view === 'balance') && state.half != null) p.set('half', state.half);
      if (state.view === 'balance' && state.balAll) { p.delete('group'); p.delete('teacher'); p.set('all', '1'); }
      if (state.view === 'balance' && !state.balAll && state.mode === 'teacher') { p.delete('teacher'); if (state.group) p.set('group', state.group); }
    }
    if (document.documentElement.classList.contains('embed')) p.set('embed', '1');
    var qs = p.toString(), target = location.pathname + (qs ? '?' + qs : '');
    if (target === location.pathname + location.search && !location.hash) return;
    try { history.replaceState(null, '', target); } catch (e) { /* ignore */ }
  }

  window.RaspAdminTools.init({
    user: isAdmin, groups: function () { return DB.groups; }, group: function () { return state.group || DB.groups[0] || ''; },
    plan: planFor, exams: function () { return EX.records; }, api: Api,
    navigate: function (view) { switchView(view); },
    examGroup: function () { return state.exPage === 'book' && state.bookGroup ? state.bookGroup : ex.group.length === 1 ? ex.group[0] : state.group || DB.groups[0] || ''; },
    showExam: function (record, retake) { state.view = 'exams'; state.exPage = 'list'; ex = exDefaults(); ex.group = [record.group]; ex.type = retake ? 'retake' : TYPE_KEY[record.type] || 'exam'; render(); },
    refresh: function () { HIST.status = 'idle'; return loadEdits().then(function () { recompute(); fillOptions(); refreshSuggestions(); render(); }); },
    toast: toast, esc: esc,
    catalogNames: catalogNames, catalog: function () { return CATALOG; }, suggestions: refreshSuggestions,
    prepare: function () { ensureAll(); refreshSuggestions(); },
    selection: selection, slots: function () { return DB.slots; },
    day: function () { return spanDay() || state.week || currentWeekIso(); },
    shift: Rules.shift, monday: function (d) { return isoOfUtc(mondayOfUtc(utcOfIso(d))); },
    fmt: fmtDayMonth, weekday: function (d) { return cap(WEEKDAYS[isoToDate(d).getDay()]); },
    until: function () { return Rules.shift(academicEnd(), -1); },
    ready: function (g) { return (isLive(g) ? Promise.resolve() : loadSheet(g)).then(function () { recompute(); if (!isLive(g)) throw new Error('Не удалось загрузить расписание группы. Повторите попытку.'); }); },
    rows: slotRows,
    expand: function (list, mode, until, study) {
      var out = [];
      list.forEach(function (r) {
        var end = mode === 'term' && r.date < half2Of(r.group) ? Rules.shift(half2Of(r.group), -1) : until;
        var dates = mode === 'once' ? [r.date] : Rules.dates(r.date, end, mode === 'alternate' ? 14 : 7, function (d) {
          if (!study) return true;
          var gw = grafikWeek(r.group, isoOfUtc(mondayOfUtc(utcOfIso(d)))), wd = (isoToDate(d).getDay() + 6) % 7;
          return !isHoliday(d) && wd !== 6 && (!GRAFIK.data || !!gw) && (!gw || wd >= 5 || gw.days.charAt(wd) === 'T');
        });
        out = out.concat(Rules.expand([r], dates, r.date, function (next) {
          var existing = EDITS.list.find(function (x) { return x.kind === 'add' && x.group === next.group && x.date === next.date && normTime(x.start) === next.start && +x.sub === +next.sub && catalogDisc(x.disc) === next.disc && catalogTeacher(x.teacher) === next.teacher && x.room === next.room; });
          return { id: existing ? existing.id : '', was: '' };
        }));
      });
      return out;
    },
    saveMany: function (list, batch) {
      var saved = 0, chain = Promise.resolve();
      for (var i = 0; i < list.length; i += 100) (function (chunk) { chain = chain.then(function () { return Api.save(chunk, {batch: batch, comment: 'добавлены пары на день или неделю'}).then(function (items) { saved += items.length; mergeEdits(items); }); }); })(list.slice(i, i + 100));
      return chain.catch(function (err) { err.message += saved ? ' Уже сохранено записей: ' + saved + '. Повторное сохранение обновит эти пары.' : ''; throw err; }).finally(function () { HIST.status = 'idle'; recompute(); fillOptions(); render(); });
    }
  });
  // ---------- loading ----------
  function isLive(group) { return !!DB.sheets[group] && (DB.sheets[group].source === 'live' || !!DB.sheets[group].fresh); }
  function allLive() { return DB.groups.length > 0 && DB.groups.every(isLive); }
  // ---------- protection against a broken sheet (ТЗ, этап 2) ----------
  // A sheet that suddenly parses into nothing (the layout changed, a tab was renamed) does not replace
  // the last good copy: the page keeps that copy, and admins see a warning and get it in the import journal.
  var good = {};
  try { good = JSON.parse(localStorage.getItem(GOOD_KEY) || '{}') || {}; } catch (e) { good = {}; }
  // Only a sheet that parsed into nothing keeps the old copy. A sheet that just got much shorter is
  // shown as it is now (the head of studies often clears whole weeks), and admins get a note about it:
  // before, such a sheet was held back and the site went on showing the old timetable.
  function brokenParse(name, size, what) {
    return size ? '' : 'На листе «' + name + '» не нашлось ' + what;
  }
  function shrinkNote(name, size, what) {
    var before = good[name] || 0;
    if (before >= 20 && size < before * 0.3) return 'На листе «' + name + '» ' + what.replace(/^ни одно[йг]о? /, '') + ' стало в ' + Math.round(before / size) + ' раз меньше, чем в прошлый раз. Показана новая версия: проверьте, что лист заполнен как нужно.';
    return '';
  }
  function scanSize(scan) { return scan.blocks.reduce(function (a, b) { return a + b.cells.length; }, 0); }
  function acceptSheet(name, csv, size, what) {
    var why = brokenParse(name, size, what);
    if (why) { net.guards[name] = why + '. Показана прошлая версия: проверьте, не изменился ли формат листа.'; return false; }
    delete net.guards[name];
    var shrink = shrinkNote(name, size, what);
    if (shrink) net.notes[name] = shrink; else delete net.notes[name];
    cachePut(name, csv);
    good[name] = size;
    try { localStorage.setItem(GOOD_KEY, JSON.stringify(good)); } catch (e) { /* storage unavailable */ }
    return true;
  }
  function loadSheet(group) {
    net.pending++;
    return fetchSheet(group, groupSheetRange()).then(function (csv) {
      var scan = scanLessonSheet(parseCSV(csv)), prev = cacheGet(group);
      if (acceptSheet(group, csv, scanSize(scan), 'ни одной пары')) {
        DB.sheets[group] = { scan: scan, source: 'live', at: Date.now(), rows: [], byDay: {}, template: false };
        if (prev && prev.csv !== csv) DB.sheets[group].prev = prev;
      } else if (prev) applySheet(group, prev.csv, 'cache', prev.t);
      else DB.sheets[group] = { scan: scan, source: 'live', at: Date.now(), rows: [], byDay: {}, template: false };
      DB.sheets[group].fresh = true;
      DB.sheets[group].guard = net.guards[group] || '';
      delete net.failed[group];
    }, function () {
      net.failed[group] = true;
      if (!DB.sheets[group]) {
        var c = cacheGet(group);
        if (c) applySheet(group, c.csv, 'cache', c.t);
      }
    }).then(function () { net.pending--; });
  }
  function ensureGroup(group) {
    if (!group || isLive(group) || net.loading === group || (DB.groups.length && DB.groups.indexOf(group) < 0)) return;
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
  function ensureAll() { ensureGroups(DB.groups); }
  // loads the given groups a few at a time; a group that failed waits for «Повторить»
  function ensureGroups(list) {
    var todo = list.filter(function (g) { return DB.groups.indexOf(g) >= 0 && !isLive(g) && !net.failed[g]; });
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
      resolvePendingHash();
      if (state.teacher && DB.teachers.indexOf(state.teacher) < 0) state.teacher = '';
      fillOptions();
      render();
      if ($('bulk-dialog').open) { bulkFillTeachers(); bulkPreview(); }
      if ($('xi-dialog').open && xi.review) renderExamReview();
      if ($('pi-dialog').open && pi.kind) renderPracticeReview();
      // a view that asked for more groups while this round was loading gets them now
      ensureView();
    });
  }
  function ensureView() {
    if (state.view === 'history') {
      loadHistory(); loadTableJournal();
      // conflicts and edits already in the table are found on the sheets of the groups with edits
      if (isAdmin()) ensureGroups(uniq(EDITS.list.map(function (e) { return e.group; })));
    }
    if (state.view === 'balance') { ensureGroups(balanceNeeds()); return; }
    if (state.view !== 'lessons' && state.view !== 'hours') return;
    if (state.mode === 'teacher' || state.mode === 'room') ensureAll();
    else if (state.group) ensureGroup(state.group);
  }
  function loadPlan() {
    if (!CFG.planSheet) { PLAN.status = 'error'; return Promise.resolve(); }
    return fetchSheet(CFG.planSheet).then(function (csv) {
      var res = Hours.parsePlan(parseCSV(csv));
      if (acceptSheet(CFG.planSheet, csv, res.order.length, 'ни одного учебного плана группы')) {
        PLAN = { status: 'ok', groups: res.groups, order: res.order, source: 'live', at: Date.now() };
        return;
      }
      var c = cacheGet(CFG.planSheet);
      if (c) applyPlan(c.csv, 'cache', c.t); else applyPlan(csv, 'live', Date.now());
    }, function () {
      var c = cacheGet(CFG.planSheet);
      if (c) applyPlan(c.csv, 'cache', c.t); else PLAN.status = 'error';
    });
  }
  function loadGrafik() {
    return fetch(CFG.grafikUrl || 'data/grafik.json', { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error(r.status);
      return r.json();
    }).then(function (d) {
      GRAFIK = { status: 'ok', data: d };
    }, function () {
      GRAFIK = { status: 'error', data: null };
    }).then(function () { recompute(); render(); });
  }
  function loadEdits() {
    if (Api.mode === 'off') return Promise.resolve();
    return Api.loadState().then(function (st) {
      setState(st, 'ok');
    }, function () {
      var c = null;
      try { c = JSON.parse(localStorage.getItem(STATE_KEY) || 'null'); } catch (e) { c = null; }
      if (c && c.state) { setState(c.state, 'cache'); EDITS.at = c.t; } else setEdits([], 'error');
    }).then(function () { recompute(); render(); });
  }
  function loadTableJournal(force) {
    if (Api.mode === 'off') { TJ.status = 'off'; return; }
    if (TJ.status === 'loading' || (TJ.status === 'ok' && !force && TJ.full === !!Api.user())) return;
    TJ.status = 'loading';
    Api.tableJournal('').then(function (r) {
      TJ = { status: 'ok', items: r.items || [], full: !!r.full };
    }, function (err) {
      TJ.status = 'error'; TJ.error = err.message;
    }).then(function () { if (state.view === 'history') render(); });
  }
  function loadHistory(force) {
    if (Api.mode === 'off') { HIST.status = 'off'; return; }
    if (HIST.status === 'loading' || (HIST.status === 'ok' && !force && HIST.full === !!Api.user())) return;
    HIST.status = 'loading';
    Api.history('').then(function (r) {
      HIST = { status: 'ok', items: r.items || [], full: !!r.full, group: HIST.group };
    }, function (err) {
      HIST.status = 'error'; HIST.error = err.message;
    }).then(function () { if (state.view === 'history') render(); });
    render();
  }
  function loadExams() {
    if (!CFG.examSheet) { EXS = { status: 'none', records: [], skipped: 0, source: '', at: 0 }; rebuildExams(); return Promise.resolve(); }
    return fetchSheet(CFG.examSheet).then(function (csv) {
      cachePut(CFG.examSheet, csv);
      applyExams(csv, 'live', Date.now());
    }, function () {
      var c = cacheGet(CFG.examSheet);
      if (c) applyExams(c.csv, 'cache', c.t);
      else { EXS = { status: 'error', records: [], skipped: 0, source: '', at: 0 }; rebuildExams(); }
    });
  }
  // exams and credits from the college's Word schedules (tools/parse_exams.py → data/exams.json)
  function loadExamFile() {
    if (!CFG.examsUrl) { EXF = { status: 'none', records: [], period: '' }; rebuildExams(); return Promise.resolve(); }
    return fetch(CFG.examsUrl, { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error(r.status);
      return r.json();
    }).then(function (d) {
      var recs = (d.records || []).filter(function (x) { return x && x.group && x.discipline && /^\d{4}-\d{2}-\d{2}$/.test(x.date) && TYPE_KEY[x.type]; });
      EXF = { status: 'ok', records: recs, period: d.period || '' };
    }, function () {
      EXF = { status: 'error', records: [], period: '' };
    }).then(function () { rebuildExams(); render(); });
  }
  function loadLists() {
    net.pending++;
    net.listFailed = false;
    render();
    var cal = CFG.calendarSheet ? fetchSheet(CFG.calendarSheet).catch(function () { return null; }) : Promise.resolve(null);
    var lists = Promise.all([fetchSheet(CFG.groupListSheet), cal]).then(function (res) {
      var groups = groupListOf(res[0]);
      if (acceptSheet(CFG.groupListSheet, res[0], groups.length, 'ни одной группы')) DB.groups = groups;
      else { var c = cacheGet(CFG.groupListSheet); DB.groups = c ? groupListOf(c.csv) : groups; }
      if (res[1]) { cachePut(CFG.calendarSheet, res[1]); applyCalendar(res[1]); }
      net.listLive = true;
    }, function () {
      net.listFailed = true;
    });
    return Promise.all([lists, loadExams(), loadExamFile(), loadPlan()]).then(function () {
      net.pending--;
      recompute();
      resolvePendingHash();
      if (net.listLive && state.group && DB.groups.indexOf(state.group) < 0) state.group = '';
      fillOptions();
      render();
      ensureView();
    });
  }
  function retry() {
    net.failed = {};
    if (EDITS.status === 'error' || EDITS.status === 'cache') loadEdits();
    if (!net.listLive || EX.status === 'error' || PLAN.status === 'error') loadLists(); else ensureView();
  }

  // ---------- lessons view ----------
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
  function optionsHTML(placeholder, values, current) {
    var list = values.slice();
    if (current && list.indexOf(current) < 0) list.unshift(current);
    return '<option value="">' + esc(placeholder) + '</option>' + list.map(function (v) {
      return '<option value="' + esc(v) + '">' + esc(v) + '</option>';
    }).join('');
  }
  function fillOptions() {
    var courses = groupsByCourse();
    $('group').innerHTML = '<option value="">Выберите группу</option>' + courses.map(function (c) {
      return '<optgroup label="' + esc(c.label) + '">' + c.groups.map(function (g) {
        return '<option value="' + esc(g) + '">' + esc(g) + '</option>';
      }).join('') + '</optgroup>';
    }).join('');
    $('room').innerHTML = optionsHTML('Выберите аудиторию', DB.rooms || [], state.room);
    $('teacher').innerHTML = optionsHTML(net.bulk ? 'Загружаю список…' : 'Выберите преподавателя', DB.teachers, state.teacher);
    $('picker-body').innerHTML = courses.map(function (c) {
      return '<div class="course"><h3 class="eyebrow">' + esc(c.label) + '</h3><div class="chips">' + c.groups.map(function (g) {
        return '<button type="button" class="chip" translate="no" data-group="' + esc(g) + '">' + esc(g) + '</button>';
      }).join('') + '</div></div>';
    }).join('');
    fillExamOptions();
  }

  function currentWeekIso() {
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
    if (state.mode === 'room') return state.room ? { mode: 'room', name: state.room } : null;
    if (state.mode === 'teacher') return state.teacher ? { mode: 'teacher', name: state.teacher } : null;
    return state.group ? { mode: 'group', name: state.group } : null;
  }
  // withGhosts adds the cancelled pairs (and, for a teacher, the pairs handed to someone else)
  function rowsForDay(iso, sel, withGhosts) {
    var out = [];
    if (sel.mode === 'group') {
      var s = DB.sheets[sel.name];
      ((s && s.byDay[iso]) || []).forEach(function (r) { if (!r.ghost || (withGhosts && !r.away)) out.push(r); });
      return out;
    }
    Object.keys(DB.sheets).forEach(function (g) {
      (DB.sheets[g].byDay[iso] || []).forEach(function (r) { if ((sel.mode === 'room' ? r.room === sel.name : splitTeachers(r.teacher).indexOf(sel.name) >= 0) && (!r.ghost || withGhosts)) out.push(r); });
    });
    return out;
  }
  function slotOf(r) { return DB.slotByStart[r.start] || { n: 0, start: r.start, end: r.end }; }

  // In teacher mode the same class given to several groups at once becomes one entry.
  function pairEntries(rows, sel) {
    var map = {}, order = [];
    rows.forEach(function (r) {
      var key = (sel.mode === 'teacher' ? [r.disc, r.room, r.sub, r.par] : [r.disc, r.teacher, r.room, r.sub, r.par]).join('|') +
        '|' + (r.edit ? r.edit.id : '') + (r.cancelled ? 'c' : '') + (r.away ? 'a' : '');
      if (!map[key]) {
        map[key] = { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub, par: r.par, groups: [], edit: r.edit || null, cancelled: !!r.cancelled, away: !!r.away, was: r.was || [], conflict: r.conflict || null, day: r.day, start: r.start };
        order.push(key);
      }
      if (map[key].groups.indexOf(r.group) < 0) map[key].groups.push(r.group);
    });
    return order.map(function (k) { return map[k]; }).sort(function (a, b) { return (a.sub || 0) - (b.sub || 0) || (a.cancelled || a.away ? 1 : 0) - (b.cancelled || b.away ? 1 : 0); });
  }
  // ---------- record type, source and change time ----------
  // «Основная» — пара из базы расписания; «Замена» — администратор поставил вместо неё другую;
  // «Ручная правка» — администратор добавил, отменил или перенёс пару.
  var REC_TYPES = { main: 'Основная', replace: 'Замена', manual: 'Ручная правка' };
  function recordType(edit) {
    if (!edit) return 'main';
    return edit.kind === 'replace' && !edit.link ? 'replace' : 'manual';
  }
  function editTime(e) { var t = Date.parse(e.updated || e.created || ''); return isNaN(t) ? 0 : t; }
  function typeTag(type) { return '<span class="rt rt-' + type + '">' + REC_TYPES[type] + '</span>'; }
  // entry: { edit, groups } — a lesson card entry or an edit wrapped the same way
  function sourceInfo(entry) {
    if (!entry.edit) {
      var sheets = entry.groups.map(function (g) { return DB.sheets[g]; }).filter(Boolean);
      var at = sheets.reduce(function (a, s) { return Math.max(a, s.at || 0); }, 0);
      var cached = sheets.some(function (s) { return s.source === 'cache'; });
      return {
        type: 'main',
        source: 'Google-таблица расписания, ' + (entry.groups.length > 1 ? 'листы ' : 'лист ') + entry.groups.map(function (g) { return '«' + g + '»'; }).join(', '),
        time: at ? (cached ? 'сохранённая копия от ' : 'загружено ') + fmtStamp(at) : ''
      };
    }
    var e = entry.edit, t = editTime(e);
    var what = e.link ? 'перенос' : e.kind === 'cancel' ? 'отмена' : e.kind === 'add' ? 'новая пара' : 'замена';
    return {
      type: recordType(e),
      source: 'админка сайта, ' + what + (e.authorName ? ', ' + e.authorName : ''),
      time: t ? 'изменено ' + fmtStamp(t) : ''
    };
  }
  function sourceHTML(entry) {
    var s = sourceInfo(entry);
    return '<p class="src">' + typeTag(s.type) + '<span>' + esc(cap(s.source)) + (s.time ? ', ' + esc(s.time) : '') + '</span></p>';
  }
  // the journal keeps the type since this version; older rows get it from the saved snapshot
  function historyType(h) {
    if (h.type) return h.type === 'замена' ? 'replace' : 'manual';
    var o = null;
    try { o = JSON.parse(h.after || h.before || 'null'); } catch (e) { o = null; }
    return o ? recordType(o) : 'manual';
  }
  function wasText(was) {
    return uniq(was.map(function (r) { return splitCode(r.disc).name + (r.teacher ? ', ' + r.teacher : '') + (r.room ? ', ' + roomLabel(r.room) : ''); })).join('; ');
  }
  function entryHTML(e, sel) {
    var meta = [], cls = 'entry', flag = '', note = '';
    if (sel.mode !== 'group') e.groups.sort(cmpRu).forEach(function (g) { meta.push('<span class="grp" translate="no">' + esc(g) + '</span>'); });
    if (sel.mode !== 'teacher' && e.teacher) meta.push('<span>' + esc(e.teacher) + '</span>');
    if (e.room) meta.push('<span>' + esc(roomLabel(e.room)) + '</span>');
    if (e.sub) meta.push('<span class="tag">' + e.sub + ' подгруппа</span>');
    if (e.par) meta.push('<span class="tag tag-par">' + (e.par === 1 ? 'по чётным' : 'по нечётным') + '</span>');
    if (e.edit) {
      if (e.cancelled) { cls += ' is-cancelled'; flag = 'Пара отменена'; }
      else if (e.away) { cls += ' is-away'; flag = 'Пару ведёт ' + (e.edit.teacher || 'другой преподаватель'); }
      else if (e.was.length) { cls += ' is-changed'; flag = 'Замена'; note = 'Было: ' + wasText(e.was); }
      else { cls += ' is-added'; flag = 'Добавлена'; }
      if (e.edit.note) note += (note ? '. ' : '') + 'Причина: ' + e.edit.note;
    }
    var disc = e.disc ? '<p class="disc">' + discHTML(e.disc) + '</p>' : '';
    var extra = '';
    if (e.conflict && !e.away) {
      cls += ' has-conflict';
      extra += '<p class="conflict"><strong>Конфликт с таблицей.</strong> Когда вносили правку, здесь было: ' + esc(wasText(e.conflict.then) || 'пары не было') +
        '. Сейчас в таблице: ' + esc(wasText(e.conflict.now) || 'пары нет') + '. Пока правку не сняли, действует она.</p>';
      if (isAdmin()) extra += '<p class="entry-acts"><button type="button" class="btn btn-small" data-keep-edit="' + esc(e.edit.id) + '">Оставить правку</button>' +
        '<button type="button" class="btn btn-quiet btn-small" data-drop-edit="' + esc(e.edit.id) + '">Убрать правку</button></p>';
    }
    if (!e.edit && e.groups.some(function (g) { return changedSlots[g + '|' + e.day + '|' + e.start]; })) flag = flag || 'Изменено в таблице';
    if (!e.edit && flag) cls += ' is-tchanged';
    if (isAdmin() && e.edit && e.edit.kind === 'add' && !e.cancelled) extra += '<p class="entry-acts"><button type="button" class="btn btn-quiet btn-small" data-open-edit="' + esc(e.edit.id) + '">' + PENCIL + 'Изменить эту пару</button></p>';
    return '<div class="' + cls + '">' + (flag ? '<span class="flag">' + esc(flag) + '</span>' : '') + disc +
      (meta.length ? '<p class="meta">' + meta.join('') + '</p>' : '') + (note ? '<p class="was">' + esc(note) + '</p>' : '') + extra + (showSrc ? sourceHTML(e) : '') + '</div>';
  }
  var changedSlots = {}; // group|day|start that changed in the table since the last visit
  var PENCIL = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>';
  var PLUS = '<svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>';
  function isAdmin() { return !!Api.user(); }
  function dayCardHTML(dayIso, wd, rows, sel, today) {
    var diff = daysBetween(today, dayIso);
    var cls = diff === 0 ? ' is-today' : (diff < 0 ? ' is-past' : '');
    var rel = relLabel(diff);
    var admin = isAdmin(), canAdd = admin;
    var head = '<header class="dhead"><span class="dname">' + WEEKDAYS_MON[wd] + '</span><span class="ddate">' + fmtDayMonth(dayIso) + '</span>' +
      (rel ? '<span class="rel">' + rel + '</span>' : '') + '</header>';
    var addDayBtn = canAdd ? '<button type="button" class="btn btn-quiet btn-small" data-multi-day="' + dayIso + '">Заполнить день</button>' : '';
    var addBtn = canAdd ? '<button type="button" class="btn btn-quiet btn-small add-pair" data-add-day="' + dayIso + '">' + PLUS + 'Добавить пару</button>' : '';
    if (!rows.length && !DB.slots.length) return '<article class="dcard' + cls + '">' + head + '<p class="dempty">Пар нет</p>' + addBtn + addDayBtn + '</article>';
    var byPair = {};
    rows.forEach(function (r) { var n = slotOf(r).n; (byPair[n] || (byPair[n] = [])).push(r); });
    var nums = Object.keys(byPair).map(Number).sort(function (a, b) { return a - b; });
    var nowMin = new Date().getHours() * 60 + new Date().getMinutes();
    var items = [];
    for (var p = 1; p <= DB.slots.length; p++) {
      var s = DB.slots[p - 1];
      if (!byPair[p]) {
        items.push('<li class="pair gap"><span class="pnum"><b>' + p + '</b></span><span>Свободная пара' + (s ? ', ' + esc(s.start) + '–' + esc(s.end) : '') + '</span>' +
          (canAdd && s ? '<button type="button" class="pair-edit" data-add-day="' + dayIso + '" data-add-start="' + esc(s.start) + '" aria-label="Поставить пару на окно ' + p + '">' + PLUS + '</button>' : '') + '</li>');
        continue;
      }
      var isNow = diff === 0 && s && nowMin >= minutesOf(s.start) && nowMin < minutesOf(s.end);
      var live = byPair[p].filter(function (r) { return !r.ghost; });
      var start = byPair[p][0].start;
      items.push('<li class="pair' + (isNow ? ' is-now' : '') + (live.length ? '' : ' is-off') + '"><span class="pnum"><b>' + p + ' пара</b><span>' + (s ? esc(s.start) + '<br>' + esc(s.end) : '') + '</span></span><div>' +
        pairEntries(byPair[p], sel).map(function (e) { return entryHTML(e, sel); }).join('') + '</div>' +
        (admin ? '<button type="button" class="pair-edit" data-edit-day="' + dayIso + '" data-edit-start="' + esc(start) + '" aria-label="Изменить ' + p + ' пару, ' + esc(WEEKDAYS_MON[wd].toLowerCase()) + '">' + PENCIL + '</button>' : '') + '</li>');
    }
    return '<article class="dcard' + cls + '">' + head + '<ol class="pairs">' + items.join('') + '</ol>' + addBtn + addDayBtn + '</article>';
  }

  // ---------- the academic year of a group, week by week (PDF schedule) ----------
  function weekTypeText(days) {
    var counts = {}, order = [];
    days.split('').forEach(function (c) { if (!counts[c]) { counts[c] = 0; order.push(c); } counts[c]++; });
    if (order.length === 1) return WEEK_TYPES[order[0]] || '';
    return order.map(function (c) { return WEEK_TYPES[c] + ' ' + count(counts[c], 'день', 'дня', 'дней'); }).join(', ');
  }
  function grafikWeek(group, weekIso) {
    var gf = grafikFor(group);
    if (!gf) return null;
    var i = GRAFIK.data.weeks.indexOf(weekIso);
    return i >= 0 ? { i: i, days: gf.g.days[i] } : null;
  }
  function renderRibbon(sel, weekIso) {
    var el = $('ribbon');
    var gf = sel && sel.mode === 'group' ? grafikFor(sel.name) : null;
    if (!gf) { el.hidden = true; el.innerHTML = ''; return; }
    var weeks = GRAFIK.data.weeks, g = gf.g, cur = currentWeekIso();
    var ci = weeks.indexOf(cur), si = weeks.indexOf(weekIso);
    var used = {};
    var cells = weeks.map(function (w, i) {
      var days = g.days[i] || 'TTTTT';
      days.split('').forEach(function (c) { used[c] = true; });
      var label = 'Неделя ' + (i + 1) + ', ' + fmtRange(w, isoOfUtc(utcOfIso(w) + 5 * DAY)) + ': ' + weekTypeText(days);
      var cls = 'rw' + (i === ci ? ' is-current' : '') + (i === si ? ' is-selected' : '') + (w < cur && i !== ci ? ' is-past' : '');
      return '<button type="button" class="' + cls + '" data-rweek="' + w + '" title="' + esc(label) + '" aria-label="' + esc(label) + '"' + (i === si ? ' aria-current="true"' : '') + '>' +
        days.split('').map(function (c) { return '<i class="d-' + c + '"></i>'; }).join('') + '</button>';
    });
    // month names over the weeks they start in (by Thursday, like ISO weeks)
    var months = [], prev = -1;
    weeks.forEach(function (w, i) {
      var m = new Date(utcOfIso(w) + 3 * DAY).getUTCMonth();
      if (m !== prev) { months.push({ m: m, from: i, n: 1 }); prev = m; } else months[months.length - 1].n++;
    });
    var MONTHS_SHORT = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    var mrow = months.map(function (x) { return '<span style="grid-column:' + (x.from + 1) + ' / span ' + x.n + '">' + (x.n > 1 ? MONTHS_SHORT[x.m] : '') + '</span>'; }).join('');
    // what is going on now and what comes next
    var lines = [];
    if (ci >= 0) {
      var now = g.days[ci];
      lines.push('Сейчас ' + (ci + 1) + '-я неделя учебного года: ' + weekTypeText(now) + '.');
      for (var k = ci + 1; k < weeks.length; k++) {
        if (g.days[k] !== now) { lines.push('С ' + fmtDayMonth(weeks[k]) + ': ' + weekTypeText(g.days[k]) + '.'); break; }
      }
    } else if (cur < weeks[0]) lines.push('Учебный год начнётся ' + fmtDayMonth(weeks[0]) + '.');
    else lines.push('Учебный год по графику закончился.');
    var legend = Object.keys(WEEK_TYPES).filter(function (c) { return used[c]; }).map(function (c) {
      return '<span><i class="d-' + c + '"></i>' + esc(WEEK_TYPES[c]) + '</span>';
    }).join('');
    var name = gf.key !== sel.name ? ' (в PDF группа записана как ' + esc(gf.key) + ')' : '';
    el.innerHTML = '<div class="ribbon-head"><p class="ribbon-now">' + esc(lines.join(' ')) + '</p>' +
      '<p class="ribbon-src">' + esc(GRAFIK.data.title) + name + '. Нажмите на неделю, чтобы открыть её расписание.</p></div>' +
      '<div class="ribbon-scroll"><div class="ribbon-grid" style="--weeks:' + weeks.length + '"><div class="ribbon-months" aria-hidden="true">' + mrow + '</div>' +
      '<div class="ribbon-weeks">' + cells.join('') + '</div></div></div><p class="legend">' + legend + '</p>';
    el.hidden = false;
  }

  function renderNow(sel, today) {
    var el = $('now'), weeks = DB.weeks;
    if (!weeks.length || today < weeks[0].iso || today > isoOfUtc(weeks[weeks.length - 1].ms + 6 * DAY)) { el.hidden = true; return; }
    var who = sel.mode === 'group' ? 'у ' + sel.name : sel.mode === 'room' ? 'в аудитории ' + sel.name : 'у преподавателя ' + sel.name;
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

  // ТЗ БП 5.8: a reminder about exams within the next 3 calendar days.
  // Credits and retakes are not shown in the general schedule, so the reminder leaves them out too.
  function renderReminder(sel, today) {
    var el = $('reminder');
    if (EX.status !== 'ok' || !sel) { el.hidden = true; return; }
    var soon = examItems().filter(function (it) {
      if (it.retake || TYPE_KEY[it.rec.type] !== 'exam') return false;
      var d = daysBetween(today, it.date);
      if (d < 0 || d > 3) return false;
      return sel.mode === 'group' ? it.rec.group === sel.name : sel.mode === 'room' ? it.rec.room === sel.name : splitTeachers(it.rec.teacher).indexOf(sel.name) >= 0;
    }).sort(byWhen);
    if (!soon.length) { el.hidden = true; return; }
    el.innerHTML = '<span class="eyebrow">Напоминание об аттестации</span>' + soon.slice(0, 3).map(function (it) {
      var r = it.rec, bits = [fmtDayMonth(it.date) + (it.time ? ', ' + it.time : ''), roomLabel(r.room), sel.mode === 'teacher' ? r.group : r.teacher].filter(Boolean);
      return '<p class="now-line"><strong>' + esc(inDays(daysBetween(today, it.date))[0].toUpperCase() + inDays(daysBetween(today, it.date)).slice(1)) + '</strong> — ' +
        esc((it.retake ? 'пересдача' : r.type.toLowerCase())) + ' по дисциплине «' + esc(splitCode(r.discipline).name) + '»: ' + esc(bits.join(', ').replace(/\.$/, '')) + '.</p>';
    }).join('') + '<p class="now-line"><a href="#" id="reminder-link">Всё расписание экзаменов →</a></p>';
    el.hidden = false;
  }

  function showEmpty(text, action, data) {
    $('empty').hidden = false;
    $('empty-text').textContent = text;
    var btn = $('empty-action');
    btn.hidden = !action;
    if (action) { btn.textContent = action; btn.setAttribute('data-action', data || ''); }
  }

  function renderLessons() {
    var today = todayISO();
    var sel = selection();
    var weekEl = $('week');
    $('empty').hidden = true;
    $('now').hidden = true;
    $('reminder').hidden = true;
    $('picker').hidden = true;
    $('ribbon').hidden = true;
    $('admin-bar').hidden = true;
    weekEl.innerHTML = '';
    var sheet = sel && sel.mode === 'group' ? DB.sheets[sel.name] : null;
    $('tpl-note').hidden = !(sheet && sheet.template);

    var dayIso = spanDay();
    if (dayIso) state.week = isoOfUtc(mondayOfUtc(utcOfIso(dayIso)));
    var weeks = DB.weeks, wi = nearestWeekIndex(state.week || currentWeekIso());
    $('weeknav').hidden = !(sel && wi >= 0);
    $('tdiff').hidden = true;
    renderSpan(dayIso);

    if (!DB.groups.length) {
      if (net.listFailed) showEmpty('Не удалось загрузить расписание из базы расписания. Проверьте подключение к интернету. Если ошибка повторяется, проверьте подключение сервера базы; открывать таблицу для всех не нужно.', 'Повторить', 'retry');
      else showEmpty('Загружаю расписание из базы расписания…');
      return;
    }
    if (!sel) {
      if (state.mode === 'group') $('picker').hidden = false;
      else showEmpty(state.mode === 'room' ? (net.bulk ? 'Собираю список аудиторий…' : 'Выберите аудиторию, чтобы увидеть её пары.') : net.bulk ? 'Собираю список преподавателей: загружено ' + net.bulk.done + ' из ' + net.bulk.total + ' групп…' : 'Выберите преподавателя, чтобы увидеть его пары.');
      return;
    }
    renderReminder(sel, today);
    if (sel.mode === 'group' && !sheet) {
      if (net.failed[sel.name]) showEmpty('Не удалось загрузить расписание группы ' + sel.name + '. Проверьте подключение к интернету.', 'Повторить', 'retry');
      else showEmpty('Загружаю расписание группы ' + sel.name + '…');
      return;
    }
    if (wi < 0) { showEmpty('В таблице нет ни одной недели с датами.'); return; }
    renderTableDiff(sel);
    if (dayIso) { renderDayView(sel, dayIso, today); return; }

    state.week = weeks[wi].iso;
    var mondayMs = weeks[wi].ms;
    var gw = sel.mode === 'group' ? grafikWeek(sel.name, weeks[wi].iso) : null;
    $('wk-label').textContent = fmtRange(weeks[wi].iso, isoOfUtc(mondayMs + 5 * DAY));
    $('wk-sub').textContent = 'Неделя ' + (Math.round((mondayMs - weeks[0].ms) / (7 * DAY)) + 1) +
      (weeks[wi].par === 1 ? ', чётная' : weeks[wi].par === 2 ? ', нечётная' : '') + (gw ? ', ' + weekTypeText(gw.days) : '');
    $('wk-prev').disabled = wi <= 0;
    $('wk-next').disabled = wi >= weeks.length - 1;
    $('wk-today').hidden = nearestWeekIndex(currentWeekIso()) === wi;
    $('wk-today').textContent = 'К текущей неделе';
    renderRibbon(sel, weeks[wi].iso);
    $('admin-bar').hidden = !isAdmin();
    weekEl.classList.remove('one-day');

    var total = 0, cards = [];
    for (var wd = 0; wd < 6; wd++) {
      var dayIso = isoOfUtc(mondayMs + wd * DAY), rows = rowsForDay(dayIso, sel, true);
      total += rows.length;
      cards.push(dayCardHTML(dayIso, wd, rows, sel, today));
    }
    if (sel.mode === 'group' || total || isAdmin()) weekEl.innerHTML = cards.join('');
    else {
      var other = findWeekWithLessons(sel, wi);
      var why = 'На эту неделю пар нет: расписание ещё не заполнено или идёт практика.';
      if (gw && !/T/.test(gw.days)) why = 'На этой неделе у группы ' + weekTypeText(gw.days) + ' по графику учебного процесса, пар в расписании нет.';
      showEmpty(sel.mode !== 'group' && net.bulk ? 'Загружаю расписание групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…' : why,
        other >= 0 ? 'Показать ' + fmtRange(weeks[other].iso, isoOfUtc(weeks[other].ms + 5 * DAY)) : '', other >= 0 ? 'week:' + weeks[other].iso : '');
    }
    renderNow(sel, today);
  }

  // ТЗ, этап 1: «сегодня, завтра, неделя»; этап 5: календарь, чтобы сразу перейти к дате
  function spanDay() {
    if (state.span === 'today') return todayISO();
    if (state.span === 'tomorrow') return isoOfUtc(utcOfIso(todayISO()) + DAY);
    if (state.span === 'day' && state.day) return state.day;
    return '';
  }
  function renderSpan(dayIso) {
    Array.prototype.forEach.call($('span').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === state.span ? 'true' : 'false');
    });
    $('goto-date').value = dayIso || state.week || '';
  }
  function renderDayView(sel, dayIso, today) {
    var weekEl = $('week'), ms = utcOfIso(dayIso), monday = mondayOfUtc(ms), wd = (new Date(ms).getUTCDay() + 6) % 7;
    var wk = DB.weeks.filter(function (w) { return w.ms === monday; })[0];
    var gw = sel.mode === 'group' ? grafikWeek(sel.name, isoOfUtc(monday)) : null;
    $('wk-label').textContent = cap(WEEKDAYS[new Date(ms).getUTCDay()]) + ', ' + fmtDayMonth(dayIso);
    $('wk-sub').textContent = (wk ? 'Неделя ' + (Math.round((monday - DB.weeks[0].ms) / (7 * DAY)) + 1) + (wk.par === 1 ? ', чётная' : wk.par === 2 ? ', нечётная' : '') : 'Вне расписания семестра') +
      (gw && wd < 5 ? ', ' + WEEK_TYPES[gw.days.charAt(wd)] : '');
    $('wk-prev').disabled = false;
    $('wk-next').disabled = false;
    $('wk-today').hidden = dayIso === today;
    $('wk-today').textContent = 'К сегодня';
    renderRibbon(sel, isoOfUtc(monday));
    $('admin-bar').hidden = !isAdmin();
    weekEl.classList.add('one-day');
    weekEl.innerHTML = dayCardHTML(dayIso, wd, rowsForDay(dayIso, sel, true), sel, today);
    if (dayIso === today) renderNow(sel, today);
  }
  function changeText(c) {
    return (c.before ? 'было «' + c.before + '»' : 'пары не было') + ', ' + (c.after ? 'стало «' + c.after + '»' : 'пару убрали');
  }
  function localDiff(group) {
    var d = null;
    try { d = JSON.parse(localStorage.getItem(DIFF_KEY + group) || 'null'); } catch (e) { d = null; }
    if (!d || !Array.isArray(d.changes) || Date.now() - d.at > 5 * DAY) return null;
    var today = todayISO();
    d.upcoming = d.changes.filter(function (c) { return c.date >= today; });
    return d.upcoming.length ? d : null;
  }
  // this browser's «что изменилось»: lessons of the group that differ from the copy of the last visit
  function renderTableDiff(sel) {
    var el = $('tdiff'), d = sel && sel.mode === 'group' ? localDiff(sel.name) : null;
    changedSlots = {};
    if (!d) { el.hidden = true; return; }
    d.upcoming.forEach(function (c) { changedSlots[sel.name + '|' + c.date + '|' + c.start] = true; });
    el.innerHTML = '<p class="tdiff-title">С вашего прошлого просмотра (' + esc(fmtStamp(d.since)) + ') в таблице изменилось ' + count(d.upcoming.length, 'пара', 'пары', 'пар') + '</p>' +
      '<ul>' + d.upcoming.slice(0, 6).map(function (c) {
        return '<li><strong>' + esc(cap(WEEKDAYS[isoToDate(c.date).getDay()])) + ', ' + esc(fmtDayMonth(c.date)) + ', ' + esc(slotLabel(c.start)) + ':</strong> ' + esc(changeText(c)) + '</li>';
      }).join('') + '</ul>' + (d.upcoming.length > 6 ? '<p class="hint">И ещё ' + (d.upcoming.length - 6) + '.</p>' : '') +
      '<p class="entry-acts"><button type="button" class="btn btn-quiet btn-small" data-tdiff-hide>Понятно, скрыть</button></p>';
    el.hidden = false;
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

  // ---------- import: snapshots of the table and «что изменилось» (ТЗ, этап 2) ----------
  function rowText(r) { return [r.disc, r.teacher, r.room ? roomLabel(r.room) : '', r.par === 1 ? 'по чётным' : r.par === 2 ? 'по нечётным' : ''].filter(Boolean).join(', '); }
  function slotTexts(rows) {
    var m = {};
    rows.forEach(function (r) { var k = r.day + '|' + r.start + '|' + (+r.sub || 0); (m[k] || (m[k] = [])).push(rowText(r)); });
    Object.keys(m).forEach(function (k) { m[k] = uniq(m[k]).sort(cmpRu).join('; '); });
    return m;
  }
  function diffRows(oldRows, newRows, group) {
    var a = slotTexts(oldRows), b = slotTexts(newRows), out = [];
    uniq(Object.keys(a).concat(Object.keys(b))).sort().forEach(function (k) {
      if ((a[k] || '') === (b[k] || '')) return;
      var part = k.split('|');
      out.push({ group: group, date: part[0], start: part[1], sub: +part[2] || 0, before: a[k] || '', after: b[k] || '' });
    });
    return out;
  }
  function snapshotOf(rows) {
    return rows.map(function (r) { return [r.day, r.start, +r.sub || 0, r.par || 0, r.disc, r.teacher, r.room].join('\t'); }).sort().join('\n');
  }
  function rowsOfSnapshot(data) {
    return data ? data.split('\n').map(function (l) {
      var f = l.split('\t');
      return { day: f[0], start: f[1], sub: +f[2] || 0, par: +f[3] || 0, disc: f[4] || '', teacher: f[5] || '', room: f[6] || '' };
    }) : [];
  }
  function hashOf(str) {
    var h = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) { h ^= str.charCodeAt(i); h = Math.imul(h, 0x01000193) >>> 0; }
    return ('0000000' + h.toString(16)).slice(-8) + ':' + str.length;
  }
  function runQueue(items, fn) {
    var queue = items.slice(), workers = [];
    function worker() { var x = queue.shift(); return x == null ? Promise.resolve() : fn(x).then(worker); }
    for (var i = 0; i < CONCURRENCY; i++) workers.push(worker());
    return Promise.all(workers);
  }
  // An admin's import compares every sheet with the snapshot of the previous import on the server and
  // writes what changed to the shared journal; a sheet the parser could not read keeps its old snapshot.
  function importToServer() {
    return Api.snapshots().then(function (snaps) {
      var groups = [], changes = [], warnings = [], first = !Object.keys(snaps).length;
      DB.groups.forEach(function (g) {
        var sh = DB.sheets[g];
        if (!sh || !sh.fresh) { warnings.push({ group: g, note: 'Лист «' + g + '» не загрузился, его прошлая версия осталась' }); return; }
        if (sh.guard) { warnings.push({ group: g, note: sh.guard }); return; }
        var data = snapshotOf(sh.base || []), hash = hashOf(data), old = snaps[g];
        if (old && old.hash === hash) return;
        if (old) changes = changes.concat(diffRows(rowsOfSnapshot(old.data), sh.base || [], g));
        groups.push({ group: g, hash: hash, data: data });
      });
      Object.keys(net.guards).forEach(function (n) { if (DB.groups.indexOf(n) < 0) warnings.push({ group: '', note: net.guards[n] }); });
      if (!groups.length && !warnings.length) return { same: true };
      return Api.importTable({ groups: groups, changes: changes.slice(0, 20000), warnings: warnings }).then(function (r) { r.first = first; return r; });
    });
  }
  // The open page keeps up with the table by itself: every syncMinutes (config.js) and whenever the
  // tab comes back to the screen, it re-reads the group list, the plan and every sheet it has loaded.
  // A sheet that changed is shown at once, with «С вашего прошлого просмотра в таблице изменилось…».
  function syncMinutes() { var m = CFG.syncMinutes == null ? 5 : +CFG.syncMinutes; return m > 0 ? m : 0; }
  function syncNow() {
    if (net.syncing || net.refreshing || net.bulk || !DB.groups.length) return;
    net.syncing = true;
    var groups = Object.keys(DB.sheets).filter(function (g) { return DB.groups.indexOf(g) >= 0; });
    var before = {};
    groups.forEach(function (g) { var c = cacheGet(g); before[g] = c ? c.csv : ''; });
    loadLists().then(function () {
      return runQueue(groups, function (g) { return loadSheet(g); });
    }).then(function () {
      net.synced = Date.now();
      var changed = groups.filter(function (g) { var c = cacheGet(g); return c && before[g] && c.csv !== before[g]; });
      recompute();
      fillOptions();
      render();
      if ($('xi-dialog').open && xi.review) renderExamReview();
      if (changed.length) toast('Таблица изменилась, расписание обновлено: ' + changed.slice(0, 6).join(', ') + (changed.length > 6 ? ' и ещё ' + (changed.length - 6) : '') + '.');
    }, function () { net.synced = Date.now(); }).then(function () { net.syncing = false; });
  }
  function syncDue() { var m = syncMinutes(); return m && Date.now() - net.synced >= m * 60000; }
  // ТЗ, этап 2: «импорт из таблицы и учебного плана по нажатию кнопки»
  function refreshAll() {
    if (net.refreshing) return;
    var admin = isAdmin();
    net.refreshing = { done: 0, total: 0 };
    net.failed = {};
    net.synced = Date.now();
    render();
    Promise.all([loadLists(), loadGrafik(), loadEdits()]).then(function () {
      var groups = admin ? DB.groups.slice() : Object.keys(DB.sheets);
      net.refreshing.total = groups.length;
      return runQueue(groups, function (g) {
        return loadSheet(g).then(function () { net.refreshing.done++; if (net.refreshing.done % 5 === 0) render(); });
      });
    }).then(function () {
      recompute();
      fillOptions();
      return admin ? importToServer() : null;
    }).then(function (res) {
      var guards = Object.keys(net.guards).length;
      if (res && res.same) toast('Таблица не изменилась с прошлого импорта.');
      else if (res) {
        toast(res.first ? 'Первый импорт: версия таблицы сохранена. Со следующего импорта журнал покажет, что изменилось.'
          : 'Импорт готов: ' + count(res.changes, 'изменение', 'изменения', 'изменений') + (res.warnings ? ', предупреждений: ' + res.warnings : '') + '.');
      } else toast(guards ? 'Обновлено. Часть листов не разобралась, для них показана прошлая версия.' : 'Расписание обновлено.');
      TJ.status = 'idle';
      if (state.view === 'history') loadTableJournal(true);
    }, function (err) { toast('Не удалось обновить: ' + err.message, true); }).then(function () {
      net.refreshing = false;
      recompute();
      render();
    });
  }

  // ---------- hours view ----------
  function nowMinutes() { var t = new Date(); return t.getHours() * 60 + t.getMinutes(); }
  function currentHalf(group) { return todayISO() < half2Of(group) ? 1 : 2; }
  function balanceHalf(group) { return state.half === 1 || state.half === 2 ? state.half : currentHalf(group); }
  function balanceKey(group) { return group + '|' + balanceHalf(group); }
  function academicEnd() { return (+acadStart().slice(0, 4) + 1) + '-08-01'; }
  function inAcademicPeriod(day, group, half) { return Rules.inPeriod(day, acadStart(), half2Of(group), academicEnd(), half); }
  function countFor(group, teacher) {
    var sheet = DB.sheets[group];
    if (!sheet) return null;
    var rows = sheet.rows.filter(function (r) { return !r.ghost && r.disc && inAcademicPeriod(r.day, group, 0) && (!teacher || splitTeachers(r.teacher).indexOf(teacher) >= 0); });
    return Hours.countHours(rows, { today: todayISO(), nowMin: nowMinutes(), half2: half2Of(group), hoursPerPair: CFG.hoursPerPair || 2 });
  }
  // Share of the theory days of the half that are already behind us, from the PDF schedule.
  function expectedShare(group, half) {
    var gf = grafikFor(group);
    if (!gf) return null;
    var weeks = GRAFIK.data.weeks, h2 = half2Of(group), today = todayISO(), all = 0, past = 0;
    weeks.forEach(function (w, i) {
      var inHalf = half === 0 || (half === 1 ? w < h2 : w >= h2);
      if (!inHalf) return;
      gf.g.days[i].split('').forEach(function (c, d) {
        if (c !== 'T') return;
        all++;
        if (isoOfUtc(utcOfIso(w) + d * DAY) < today) past++;
      });
    });
    return all ? past / all : null;
  }
  // the last week of the half that the timetable of the group covers
  function coverage(group, half) {
    var sheet = DB.sheets[group], h2 = half2Of(group), last = '';
    if (!sheet) return '';
    sheet.rows.forEach(function (r) {
      if (r.ghost) return;
      var inHalf = half === 0 || (half === 1 ? r.day < h2 : r.day >= h2);
      if (inHalf && r.day > last) last = r.day;
    });
    return last;
  }
  function lastTheoryDay(group, half) {
    var gf = grafikFor(group);
    if (!gf) return '';
    var weeks = GRAFIK.data.weeks, h2 = half2Of(group), last = '';
    weeks.forEach(function (w, i) {
      var inHalf = half === 0 || (half === 1 ? w < h2 : w >= h2);
      if (inHalf && /T/.test(gf.g.days[i])) last = isoOfUtc(utcOfIso(w) + 4 * DAY);
    });
    return last;
  }
  function hoursWord(n) { return n.toLocaleString('ru-RU') + ' ч'; }
  function barHTML(done, plan, all, expect) {
    if (!plan) return '<span class="bar bar-none" aria-hidden="true"></span>';
    var d = Math.min(100, done / plan * 100), a = Math.min(100, all / plan * 100);
    return '<span class="bar" aria-hidden="true"><span class="bar-all" style="width:' + a.toFixed(1) + '%"></span><span class="bar-done" style="width:' + d.toFixed(1) + '%"></span>' +
      (expect != null ? '<span class="bar-mark" style="left:' + (expect * 100).toFixed(1) + '%"></span>' : '') + '</span>';
  }
  function formHTML(f) {
    if (!f) return '';
    var full = Hours.FORM_NAMES[f];
    return full ? '<abbr title="' + esc(full) + '">' + esc(f) + '</abbr>' : esc(f);
  }
  function hoursTableHTML(sum, half, opts) {
    var complete = opts.complete, expect = opts.expect;
    var rows = sum.lines.map(function (l) {
      var s = l.subject, forms = half === 0 ? uniq(s.form) : [s.form[half - 1]];
      var gap = l.all - l.plan, warn = '';
      if (complete && l.plan && gap < 0) warn = '<span class="warn">в расписании на ' + hoursWord(-gap) + ' меньше плана</span>';
      else if (complete && l.plan && gap > 0) warn = '<span class="warn">в расписании на ' + hoursWord(gap) + ' больше плана</span>';
      var bare = function (x) { return String(x).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, ''); };
      var other = l.names.filter(function (n) { return bare(n) !== bare(s.name); });
      var names = other.length ? '<span class="aka">в расписании: ' + esc(other.join('; ')) + '</span>' : '';
      var teacher = Object.keys(l.teachers).join(', ') || s.teacher;
      return '<tr><td class="title" data-label="Предмет">' + discHTML(s.name) + names + warn + '</td>' +
        '<td class="col-minor" data-label="Преподаватель">' + esc(teacher || '—') + '</td>' +
        '<td data-label="Форма">' + (forms.filter(Boolean).map(formHTML).join(', ') || '—') + '</td>' +
        '<td class="n" data-label="План">' + (l.plan ? hoursWord(l.plan) : '—') + '</td>' +
        '<td class="n" data-label="Проведено">' + hoursWord(l.done) + '</td>' +
        '<td class="n strong" data-label="Осталось">' + (l.plan ? hoursWord(l.left) : '—') + '</td>' +
        '<td class="barcell" data-label="Ход">' + barHTML(l.done, l.plan, l.all, expect) + '</td></tr>';
    });
    var extra = sum.extra.map(function (l) {
      return '<tr class="is-extra"><td class="title" data-label="Предмет">' + discHTML(l.names.join('; ')) + '<span class="aka">нет в учебном плане группы</span></td>' +
        '<td class="col-minor" data-label="Преподаватель">' + esc(Object.keys(l.teachers).join(', ') || '—') + '</td><td data-label="Форма">—</td><td class="n" data-label="План">—</td>' +
        '<td class="n" data-label="Проведено">' + hoursWord(l.done) + '</td><td class="n" data-label="Осталось">—</td><td class="barcell" data-label="Ход"></td></tr>';
    });
    return '<div class="table-wrap"><table class="smart hours"><thead><tr><th scope="col">Предмет</th><th scope="col" class="col-minor">Преподаватель</th><th scope="col">Форма</th>' +
      '<th scope="col" class="n">План</th><th scope="col" class="n">Проведено</th><th scope="col" class="n">Осталось</th><th scope="col">Ход</th></tr></thead><tbody>' +
      rows.join('') + extra.join('') + '</tbody></table></div>';
  }
  function halfName(h) { return h === 0 ? 'учебный год' : h === 1 ? 'первый семестр' : 'второй семестр'; }
  function renderHours() {
    var sel = selection(), box = $('hours-body'), empty = $('h-empty');
    box.innerHTML = '';
    empty.hidden = true;
    var msg = function (t, action, data) { empty.hidden = false; $('h-empty-text').textContent = t; var b = $('h-empty-action'); b.hidden = !action; if (action) { b.textContent = action; b.setAttribute('data-action', data); } };
    if (!sel) { msg(state.mode === 'group' ? 'Выберите группу, чтобы увидеть часы по её предметам.' : 'Выберите преподавателя, чтобы увидеть часы по его предметам.'); return; }
    var group = sel.mode === 'group' ? sel.name : '';
    var half = state.half != null ? state.half : currentHalf(group || DB.groups[0] || '');
    Array.prototype.forEach.call($('half').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', +b.getAttribute('data-value') === half ? 'true' : 'false');
    });
    if (PLAN.status === 'loading') { msg('Загружаю учебный план с листа «' + CFG.planSheet + '»…'); return; }
    if (PLAN.status === 'error') { msg('Не удалось прочитать учебный план с листа «' + CFG.planSheet + '». Проверьте подключение к интернету.', 'Повторить', 'retry'); return; }
    if (sel.mode === 'group') {
      if (!DB.sheets[group]) { msg(net.failed[group] ? 'Не удалось загрузить расписание группы ' + group + '.' : 'Загружаю расписание группы ' + group + '…', net.failed[group] ? 'Повторить' : '', 'retry'); return; }
      var plan = planFor(group);
      var sum = Hours.summarize(plan, countFor(group, ''), half);
      var tot = sum.lines.reduce(function (a, l) { a.plan += l.plan; a.done += l.done; a.all += l.all; return a; }, { plan: 0, done: 0, all: 0 });
      var expect = expectedShare(group, half);
      var last = coverage(group, half), end = lastTheoryDay(group, half);
      var complete = !!last && !!end && daysBetween(last, end) <= 7;
      var notes = [];
      if (!plan) notes.push('Для группы ' + group + ' на листе «' + CFG.planSheet + '» нет учебного плана, поэтому плановых часов нет.');
      else if (plan.group !== group) notes.push('Учебный план взят из блока «' + plan.group + '».');
      if (last && !complete) notes.push('Расписание этой группы заполнено до ' + fmtDayMonth(last) + ', поэтому сравнивать его с планом на весь семестр ещё рано.');
      if (sheetTemplate(group)) notes.push('У группы только недельный шаблон без дат: проведённые часы посчитаны так, будто шаблон шёл каждую неделю семестра.');
      var head = '<section class="hsum" aria-label="Итого">' +
        '<p class="hsum-line">' + esc(cap(halfName(half))) + ': проведено <strong>' + hoursWord(tot.done) + '</strong> из ' + hoursWord(tot.plan) + ' по плану, осталось <strong>' + hoursWord(Math.max(0, tot.plan - tot.done)) + '</strong>.</p>' +
        barHTML(tot.done, tot.plan, tot.all, expect) +
        '<p class="hsum-note">' + (expect != null ? 'Чёрточка на полосе показывает, сколько теоретических дней ' + (half === 0 ? 'учебного года' : 'семестра') + ' уже прошло по графику (' + Math.round(expect * 100) + '%). ' : '') +
        'Одна пара считается за ' + count(CFG.hoursPerPair || 2, 'час', 'часа', 'часов') + '.' + (plan && plan.title ? ' ' + esc(plan.title) + '.' : '') + '</p>' +
        (notes.length ? '<p class="hsum-note">' + esc(notes.join(' ')) + '</p>' : '') + '</section>';
      box.innerHTML = head + (sum.lines.length || sum.extra.length ? '<section class="tablebox">' + hoursTableHTML(sum, half, { complete: complete, expect: expect }) + '</section>' : '');
      return;
    }
    // teacher: every group where the teacher has pairs
    if (net.bulk) { msg('Собираю расписание всех групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…'); return; }
    var parts = [], grand = { plan: 0, done: 0 };
    Object.keys(DB.sheets).sort(cmpRu).forEach(function (g) {
      var counted = countFor(g, sel.name);
      if (!counted || !Object.keys(counted).length) return;
      var h = state.half != null ? state.half : currentHalf(g);
      var sum = Hours.summarize(planFor(g), counted, h);
      sum.lines = sum.lines.filter(function (l) { return l.names.length; });
      if (!sum.lines.length && !sum.extra.length) return;
      sum.lines.forEach(function (l) { grand.plan += l.plan; grand.done += l.done; });
      parts.push('<section class="tablebox"><div class="table-head"><h2>' + esc(g) + '</h2><span class="count">' + esc(halfName(h)) + '</span></div>' +
        hoursTableHTML(sum, h, { complete: false, expect: expectedShare(g, h) }) + '</section>');
    });
    if (!parts.length) { msg('В загруженном расписании у преподавателя ' + sel.name + ' нет пар.'); return; }
    box.innerHTML = '<section class="hsum"><p class="hsum-line">' + esc(sel.name) + ': проведено <strong>' + hoursWord(grand.done) + '</strong> из ' + hoursWord(grand.plan) +
      ' по плану, осталось <strong>' + hoursWord(Math.max(0, grand.plan - grand.done)) + '</strong>.</p>' + barHTML(grand.done, grand.plan, grand.done, null) +
      '<p class="hsum-note">План — часы предмета в группе по учебному плану. Если предмет в группе ведут двое, план общий на обоих.</p></section>' + parts.join('');
  }
  function sheetTemplate(group) { return DB.sheets[group] && DB.sheets[group].template; }
  function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

  // ---------- вычитка часов ----------
  // For the head of studies: plan of the half against every pair that is held or still stands in the
  // timetable up to the last day before the session, after all admin edits. Nothing here reads the
  // «Пары / Часы» formulas of the college table: a pair added, cancelled or moved on the site
  // changes the numbers at once (recompute → render).
  var bal = { open: '', until: {}, wantAll: false };
  function normGroup(g) { return String(g).toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ''); }
  function urgentGroups() {
    return (CFG.urgentGroups || []).map(function (u) {
      var hit = DB.groups.filter(function (g) { return normGroup(g) === normGroup(u); })[0];
      return hit || String(u).trim();
    }).filter(Boolean);
  }
  function isManualUrgent(g) { return urgentGroups().some(function (u) { return normGroup(u) === normGroup(g); }); }
  // Groups catch fire by themselves too: theory ends and practice starts in urgentDays (config.js) or
  // less, or theory is already over, and pairs are still missing. The dates come from the PDF schedule
  // with the admins' corrections, so moving a practice on the «График» tab moves the fire too.
  function urgentDays() { var d = CFG.urgentDays == null ? 14 : +CFG.urgentDays; return d > 0 ? d : 0; }
  // the next practice of the group in this half: its first day, the last theory day before it and how many
  // theory days are left until then; with no practice ahead, the end of theory of the half
  function practiceAhead(group) {
    var half = balanceHalf(group), days = halfDays(group, half), today = todayISO();
    if (!days.length) return null;
    var start = '', now = false;
    for (var w = 0; w + 4 < days.length; w += 5) {
      var wk = days.slice(w, w + 5), letters = wk.map(function (d) { return d.c; }).join('');
      if (/T/.test(letters) || !/[UPX]/.test(letters) || wk[4].iso < today) continue;
      var first = wk.filter(function (d) { return /[UPX]/.test(d.c); })[0];
      start = first.iso; now = first.iso <= today;
      break;
    }
    var lastT = '';
    days.forEach(function (d) { if (d.c === 'T' && (!start || d.iso < start)) lastT = d.iso; });
    var left = days.filter(function (d) { return d.c === 'T' && d.iso > today && d.iso <= lastT && !isHoliday(d.iso); }).length;
    return { start: start, now: now, lastT: lastT, left: left, days: start ? daysBetween(today, start) : null };
  }
  // why a group burns: '' when it does not
  function fireReason(g) {
    if (isManualUrgent(g)) return 'manual';
    var d = urgentDays(), pa = d ? practiceAhead(g) : null;
    if (!pa) return '';
    var soon = pa.start && !pa.now && pa.days <= d, over = pa.lastT && pa.lastT < todayISO();
    if (!soon && !over) return '';
    var b = balanceFor(g);
    if (!b || !b.plan || !b.total.add) return b ? '' : 'maybe';
    return soon ? 'soon' : 'over';
  }
  function isUrgent(g) { var r = fireReason(g); return r && r !== 'maybe'; }
  function fireCandidates() {
    var d = urgentDays(), today = todayISO();
    if (!d) return [];
    return DB.groups.filter(function (g) {
      var pa = practiceAhead(g);
      return pa && ((pa.start && !pa.now && pa.days <= d) || (pa.lastT && pa.lastT < today));
    });
  }
  function burningGroups() {
    var manual = urgentGroups(), auto = fireCandidates().filter(function (g) { return manual.indexOf(g) < 0 && fireReason(g) !== ''; });
    auto.sort(function (a, b) { var A = practiceAhead(a), B = practiceAhead(b); return (A.start || 'z') < (B.start || 'z') ? -1 : (A.start || 'z') > (B.start || 'z') ? 1 : cmpRu(a, b); });
    return manual.concat(auto);
  }
  function practiceLine(g) {
    var pa = practiceAhead(g);
    if (!pa) return '';
    if (pa.start && pa.now) return 'идёт практика с ' + fmtDayMonth(pa.start);
    if (pa.start) return 'практика с ' + fmtDayMonth(pa.start) + (pa.left ? ', до неё ' + count(pa.left, 'учебный день', 'учебных дня', 'учебных дней') : ', учебных дней до неё нет');
    if (pa.lastT && pa.lastT < todayISO()) return 'теория закончилась ' + fmtDayMonth(pa.lastT);
    return pa.lastT ? 'теория до ' + fmtDayMonth(pa.lastT) + ', ' + count(pa.left, 'учебный день', 'учебных дня', 'учебных дней') : '';
  }
  function balanceGroup() {
    if (state.group && (!DB.groups.length || DB.groups.indexOf(state.group) >= 0)) return state.group;
    return urgentGroups()[0] || '';
  }
  function balanceNeeds() {
    if (state.balAll || bal.wantAll) return DB.groups.slice();
    var list = urgentGroups().concat(fireCandidates()), g = balanceGroup();
    if (g && list.indexOf(g) < 0) list.push(g);
    return list;
  }
  // Mon–Fri of the half with the letter of the PDF schedule (with the admins' corrections)
  function halfDays(group, half) {
    var gf = grafikFor(group);
    if (!gf) return [];
    var h2 = half2Of(group), out = [];
    GRAFIK.data.weeks.forEach(function (w, i) {
      if (half !== 0 && (half === 1 ? w >= h2 : w < h2)) return;
      for (var d = 0; d < 5; d++) out.push({ iso: isoOfUtc(utcOfIso(w) + d * DAY), c: gf.g.days[i].charAt(d) });
    });
    return out;
  }
  function pairKey(r) { return r.day + '|' + r.start + '|' + r.disc; }
  // public holidays (ТК РФ, ст. 112) as MM-DD or YYYY-MM-DD, config.js → holidays
  var HOLIDAYS_DEFAULT = ['01-01', '01-02', '01-03', '01-04', '01-05', '01-06', '01-07', '01-08', '02-23', '03-08', '05-01', '05-09', '06-12', '11-04'];
  function isHoliday(iso) {
    var list = CFG.holidays || HOLIDAYS_DEFAULT;
    return list.indexOf(iso) >= 0 || list.indexOf(iso.slice(5)) >= 0;
  }
  var balMemo = {};
  function balanceFor(group) {
    var sheet = DB.sheets[group];
    if (!sheet) return null;
    var mk = balanceKey(group) + '|' + (bal.until[balanceKey(group)] || '') + '|' + sheet.at + '|' + EDITS.at + '|' + todayISO();
    if (balMemo[group] && balMemo[group].key === mk) return balMemo[group].value;
    var value = balanceCompute(group, sheet);
    balMemo[group] = { key: mk, value: value };
    return value;
  }
  function balanceCompute(group, sheet) {
    var half = balanceHalf(group), h2 = half2Of(group), per = CFG.hoursPerPair || 2, today = todayISO();
    var days = halfDays(group, half);
    var dl = days.length ? Hours.balanceDeadline(days) : { theoryEnd: '', until: '' };
    var covered = coverage(group, half);
    // by default the hours are read by the end of theory in the PDF schedule: the weeks after it are
    // practice, and a timetable stretched over them (a weekly pattern pasted to December) is not real.
    // «Перед сессией» moves the deadline to the last day before the session, for pairs held during practice.
    var autoUntil = dl.theoryEnd || dl.until || covered;
    var until = bal.until[balanceKey(group)] || autoUntil;
    var inHalf = function (iso) { return inAcademicPeriod(iso, group, half); };
    var rows = sheet.rows.filter(function (r) { return !r.ghost && r.disc && inHalf(r.day); });
    var holiday = {};
    rows.forEach(function (r) { if (isHoliday(r.day) && (!until || r.day <= until)) holiday[pairKey(r)] = r.day; });
    var termEnds = [null, Hours.balanceDeadline(halfDays(group, 1)), Hours.balanceDeadline(halfDays(group, 2))];
    var counted = rows.filter(function (r) {
      var h = r.day < h2 ? 1 : 2, limit = termEnds[h].theoryEnd || termEnds[h].until;
      return (!until || r.day <= until) && !isHoliday(r.day) &&
        (half !== 0 || bal.until[balanceKey(group)] || !limit || r.day <= limit);
    });
    var after = {}, onPractice = {}, types = {};
    // whole weeks of practice only: groups with a practice day inside a theory week (СПм-25: «TUTTT»)
    // often hold that day on another weekday than the PDF says
    for (var w = 0; w < days.length; w += 5) {
      var wk = days.slice(w, w + 5), letters = wk.map(function (d) { return d.c; }).join('');
      if (!/T/.test(letters)) wk.forEach(function (d) { types[d.iso] = d.c; });
    }
    rows.forEach(function (r) {
      if (until && r.day > until) after[pairKey(r)] = true;
      else if (/[UPX]/.test(types[r.day] || '') && !Hours.isPractice(r.disc)) onPractice[pairKey(r)] = r.day;
    });
    var practiceDays = Object.keys(onPractice).map(function (k) { return onPractice[k]; }).sort();
    var sum = Hours.summarize(planFor(group), Hours.countHours(counted, { today: today, nowMin: nowMinutes(), half2: h2, hoursPerPair: per }), half);
    var b = Hours.balance(sum, per);
    var left = 0;
    if (days.length) left = days.filter(function (d) { return d.iso > today && d.iso <= until && d.c !== 'K' && d.c !== 'A' && d.c !== 'G' && !isHoliday(d.iso); }).length;
    else if (until) for (var t = utcOfIso(today) + DAY; t <= utcOfIso(until); t += DAY) { var wd = new Date(t).getUTCDay(); if (wd > 0 && wd < 6 && !isHoliday(isoOfUtc(t))) left++; }
    var holidayDays = uniq(Object.keys(holiday).map(function (k) { return holiday[k]; })).sort();
    var weeksLeft = Math.max(1, Math.ceil(left / 5));
    b.lines.forEach(function (l) { l.perWeek = l.status === 'add' ? Math.ceil(l.pairs / weeksLeft) : 0; });
    return {
      group: group, half: half, plan: planFor(group), lines: b.lines, total: b.total, theoryEnd: dl.theoryEnd, until: until, autoUntil: autoUntil, sessionUntil: dl.until, passed: !!until && until < today,
      left: left, covered: covered, after: Object.keys(after).length, onPractice: practiceDays.length, practiceFrom: practiceDays[0] || '', onHoliday: Object.keys(holiday).length, holidayDays: holidayDays, edits: EDITS.list.filter(function (e) { return e.group === group; }).length,
      guessed: sheet.scan.guessed || 0, template: sheet.template, source: sheet.source, at: sheet.at
    };
  }
  function pairsWord(n) { return count(n, 'пара', 'пары', 'пар'); }
  function signed(n) { return n > 0 ? '+' + n : n < 0 ? '−' + (-n) : '0'; }
  function balTone(b) { return b.total.add > 0 ? 'bad' : b.total.cut > 0 ? 'warn' : 'good'; }
  function pairsAcc(n) { return count(n, 'пару', 'пары', 'пар'); }
  // «12 пар стоят», «1 пара стоит»
  function pairsStand(n) { return pairsWord(n) + (plural(n, 'x', 'y', 'z') === 'x' ? ' стоит' : ' стоят'); }
  function fireTag(g) {
    var r = fireReason(g);
    if (r === 'manual') return '<span class="bal-tag">горит</span>';
    if (r === 'soon') return '<span class="bal-tag" title="До практики осталось мало времени, а пары ещё не вычитаны">горит: практика скоро</span>';
    if (r === 'over') return '<span class="bal-tag" title="Теория закончилась, а пары не вычитаны">горит: теория прошла</span>';
    return '';
  }
  function balCardHTML(g, cur) {
    var b = balanceFor(g), on = g === cur ? ' is-on' : '';
    var inner, pl = practiceLine(g);
    if (!b) {
      inner = '<span class="bal-card-big">' + (net.failed[g] ? 'не загрузилась' : 'загружаю…') + '</span>';
      return '<button type="button" class="bal-card' + on + '" data-bal-group="' + esc(g) + '" aria-pressed="' + (on ? 'true' : 'false') + '"><span class="bal-card-head"><span class="bal-card-name" translate="no">' + esc(g) + '</span>' + fireTag(g) + '</span>' + inner + '</button>';
    }
    var t = b.total, tone = balTone(b);
    var big = !b.plan ? '—' : t.add ? '+' + t.add : t.cut ? '−' + t.cut : '0';
    var cap = !b.plan ? 'нет учебного плана' : t.add ? plural(t.add, 'пару', 'пары', 'пар') + ' добавить' : t.cut ? plural(t.cut, 'пару', 'пары', 'пар') + ' снять' : 'часы сходятся';
    var sub = [];
    if (t.add && t.cut) sub.push('и снять ' + pairsAcc(t.cut));
    if (b.until) sub.push(b.passed ? 'срок вычитки прошёл ' + fmtDayMonth(b.until) : 'до ' + fmtDayMonth(b.until) + ', ' + count(b.left, 'учебный день', 'учебных дня', 'учебных дней'));
    return '<button type="button" class="bal-card is-' + tone + on + '" data-bal-group="' + esc(g) + '" aria-pressed="' + (on ? 'true' : 'false') + '">' +
      '<span class="bal-card-head"><span class="bal-card-name" translate="no">' + esc(g) + '</span>' + fireTag(g) + '</span>' +
      '<span class="bal-card-big"><strong>' + big + '</strong> ' + esc(cap) + '</span>' +
      (sub.length ? '<span class="bal-card-sub">' + esc(sub.join(' · ')) + '</span>' : '') +
      (pl ? '<span class="bal-card-sub bal-card-practice">' + esc(pl) + '</span>' : '') + '</button>';
  }
  function lineTeacher(l) {
    var t = l.teachers.filter(function (x) { return !/^вак/i.test(x); })[0];
    if (t) return t;
    var p = splitTeachers(String(l.planTeacher).replace(/\n/g, ',').replace(/\//g, ','))[0] || '';
    return /^вак/i.test(p) ? '' : p;
  }
  function balActionHTML(b, l, i) {
    var admin = isAdmin(), key = b.group + '|' + i;
    if (l.status === 'add') {
      return '<span class="bal-act is-bad">Добавить ' + pairsAcc(l.pairs) + '</span>' +
        (b.left ? '<span class="bal-sub">≈ ' + count(l.perWeek, 'пара', 'пары', 'пар') + ' в неделю</span>' : '<span class="bal-sub">учебные дни до срока закончились</span>') +
        (admin ? '<button type="button" class="btn btn-small" data-bal-open="' + esc(key) + ':add" aria-expanded="' + (bal.open === key + ':add') + '">Подобрать время</button>' : '');
    }
    if (l.status === 'cut') {
      var per = CFG.hoursPerPair || 2, can = Math.min(l.pairs, Math.floor(l.future / per)), over = l.done - l.plan;
      if (!can) return '<span class="bal-act is-warn">Перечитано на ' + hoursWord(-l.gap) + '</span><span class="bal-sub">пары уже проведены, снимать нечего</span>';
      return '<span class="bal-act is-warn">Снять ' + pairsAcc(can) + '</span>' +
        (over > 0 ? '<span class="bal-sub">и ещё ' + hoursWord(over) + ' уже проведено сверх плана</span>' : '') +
        (admin ? '<button type="button" class="btn btn-small" data-bal-open="' + esc(key) + ':cut" aria-expanded="' + (bal.open === key + ':cut') + '">Какие снять</button>' : '');
    }
    if (l.status === 'ok') return '<span class="bal-act is-good">Сходится</span>';
    if (l.status === 'practice') return '<span class="bal-act">Практика по графику</span><span class="bal-sub">идёт неделями, парами не вычитывается</span>';
    if (l.status === 'noplan') return '<span class="bal-act is-warn">В этом семестре по плану часов нет</span>';
    return '<span class="bal-act is-warn">Нет в учебном плане</span><span class="bal-sub">проверьте название в таблице</span>';
  }
  // free time of the group and of the teacher, from tomorrow to the deadline: first a pair right next to
  // the pairs the group already has that day (no gap), then other free pairs on such days, then empty days
  function freeSlots(b, l) {
    var group = b.group, teacher = lineTeacher(l), out = [], today = todayISO();
    var slots = (CFG.slots || []).map(function (x, i) { return { n: i + 1, start: normTime(x[0]), end: normTime(x[1]) }; });
    var types = {};
    halfDays(group, b.half).forEach(function (d) { types[d.iso] = d.c; });
    var sheet = DB.sheets[group];
    for (var t = utcOfIso(today) + DAY; b.until && t <= utcOfIso(b.until); t += DAY) {
      var iso = isoOfUtc(t), wd = new Date(t).getUTCDay();
      if (wd === 0 || isHoliday(iso)) continue;
      var c = types[iso] || (wd === 6 ? types[isoOfUtc(t - DAY)] : '') || '';
      if (c === 'K' || c === 'A' || c === 'G') continue;
      var mine = (sheet.byDay[iso] || []).filter(function (r) { return !r.ghost; });
      var busy = {}, nums = [];
      mine.forEach(function (r) { busy[r.start] = true; var s = DB.slotByStart[r.start]; if (s) nums.push(s.n); });
      var lo = nums.length ? Math.min.apply(null, nums) : 0, hi = nums.length ? Math.max.apply(null, nums) : 0;
      slots.forEach(function (s) {
        if (busy[s.start]) return;
        if (teacher && teacherBusy(teacher, iso, s.start)) return;
        var rank = !nums.length ? 3 : (s.n === hi + 1 || s.n === lo - 1) ? 0 : 1;
        if (nums.length >= 4) rank++;
        out.push({ day: iso, start: s.start, end: s.end, n: s.n, rank: rank, empty: !nums.length });
      });
    }
    out.sort(function (a, b2) { return a.rank - b2.rank || (a.day < b2.day ? -1 : a.day > b2.day ? 1 : a.n - b2.n); });
    return { teacher: teacher, list: out.slice(0, 8), total: out.length };
  }
  function teacherBusy(teacher, iso, start) {
    return Object.keys(DB.sheets).some(function (g) {
      return (DB.sheets[g].byDay[iso] || []).some(function (r) { return !r.ghost && r.start === start && splitTeachers(r.teacher).indexOf(teacher) >= 0; });
    });
  }
  function lineRows(b, l) {
    var sheet = DB.sheets[b.group], today = todayISO(), seen = {};
    return sheet.rows.filter(function (r) {
      if (r.ghost || l.names.indexOf(r.disc) < 0 || r.day <= today || (b.until && r.day > b.until) || isHoliday(r.day) || seen[pairKey(r)]) return false;
      seen[pairKey(r)] = true;
      return true;
    }).sort(function (a, c) { return a.day < c.day ? 1 : a.day > c.day ? -1 : minutesOf(c.start) - minutesOf(a.start); });
  }
  function dayLabel(iso) { return ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'][isoToDate(iso).getDay()] + ', ' + fmtDayMonth(iso); }
  function balMoreHTML(b, l, i, kind) {
    var key = b.group + '|' + i;
    if (kind === 'add') {
      var fs = freeSlots(b, l), loading = net.bulk || DB.groups.some(function (g) { return !DB.sheets[g]; });
      var head = '<p>Свободное время у группы <span class="grp" translate="no">' + esc(b.group) + '</span>' + (fs.teacher ? ' и у преподавателя ' + esc(fs.teacher) : '') + ' до ' + esc(fmtDayMonth(b.until)) + ':</p>';
      var notes = [];
      if (loading) notes.push(net.bulk ? 'Проверяю занятость преподавателя по всем группам: ' + net.bulk.done + ' из ' + net.bulk.total + '…' : 'Занятость преподавателя в остальных группах ещё не загружена.');
      if (!fs.teacher) notes.push('Преподаватель не указан, его занятость не проверялась.');
      notes.push('Первыми идут пары встык к тем, что у группы уже стоят, чтобы не было окон. Пара ставится как правка администратора: таблицу колледжа сайт не меняет.');
      if (!fs.list.length) return head.replace(/:<\/p>$/, '') + ': свободного времени нет.</p>' + '<p class="hint">' + esc(notes.join(' ')) + '</p>';
      return head + '<ul class="bal-slots">' + fs.list.map(function (s) {
        return '<li><span>' + esc(dayLabel(s.day)) + ', ' + s.n + ' пара, ' + esc(s.start) + '–' + esc(s.end) + (s.empty ? ' <span class="hint">(других пар в этот день нет)</span>' : '') + '</span>' +
          '<button type="button" class="btn btn-small btn-primary" data-bal-put="' + esc(key + '|' + s.day + '|' + s.start) + '">Поставить</button></li>';
      }).join('') + '</ul><p class="hint">' + esc(notes.join(' ')) + (fs.total > fs.list.length ? ' Показаны 8 лучших вариантов из ' + fs.total + '.' : '') + '</p>';
    }
    var rows = lineRows(b, l).slice(0, l.pairs + 4);
    if (!rows.length) return '<p>Будущих пар этого предмета до ' + esc(fmtDayMonth(b.until)) + ' нет.</p>';
    return '<p>Последние пары предмета до ' + esc(fmtDayMonth(b.until)) + ', снять нужно ' + esc(pairsAcc(l.pairs)) + ':</p><ul class="bal-slots">' + rows.map(function (r) {
      var s = DB.slotByStart[r.start], shared = slotRows(b.group, r.day, r.start).some(function (x) { return !x.ghost && x.disc !== r.disc && (!r.sub || !x.sub || x.sub === r.sub); });
      var id = r.day + '|' + r.start + '|' + (r.sub || 0);
      return '<li><span>' + esc(dayLabel(r.day)) + ', ' + (s ? s.n + ' пара' : esc(r.start)) + (r.sub ? ', ' + r.sub + ' подгруппа' : '') + (r.teacher ? ', ' + esc(r.teacher) : '') + (r.edit ? ' <span class="hint">(поставлена правкой)</span>' : '') + '</span>' +
        (shared ? '<span class="hint">в это время у группы ещё пара — отмените в расписании</span>' : '<button type="button" class="btn btn-small btn-danger" data-bal-cut="' + esc(key + '|' + id) + '">Отменить</button>') + '</li>';
    }).join('') + '</ul><p class="hint">Отмена — правка администратора, её можно убрать в разделе «Изменения».</p>';
  }
  function balTableHTML(b) {
    var rows = b.lines.map(function (l, i) {
      var key = b.group + '|' + i, quiet = l.status === 'practice' || l.status === 'ok';
      var bare = function (x) { return String(x).toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, ''); };
      var other = l.names.filter(function (n) { return bare(n) !== bare(l.name); });
      var tr = '<tr class="' + (quiet ? 'is-quiet' : '') + '"><td class="title" data-label="Предмет">' + discHTML(l.name) + (other.length ? '<span class="aka">в расписании: ' + esc(other.join('; ')) + '</span>' : '') + '</td>' +
        '<td class="col-minor" data-label="Преподаватель">' + esc(l.teachers.join(', ') || l.planTeacher || '—') + '</td>' +
        '<td class="n" data-label="План">' + (l.plan ? hoursWord(l.plan) : '—') + '</td>' +
        '<td class="n" data-label="Проведено">' + hoursWord(l.done) + '</td>' +
        '<td class="n" data-label="Ещё стоит">' + hoursWord(l.future) + '</td>' +
        '<td class="n strong" data-label="План − расписание">' + (l.plan || l.status === 'extra' ? signed(l.gap) + ' ч' : '—') + '</td>' +
        '<td class="bal-do" data-label="Что сделать">' + balActionHTML(b, l, i) + '</td></tr>';
      if (bal.open === key + ':add' && l.status === 'add') tr += '<tr class="bal-more"><td colspan="7">' + balMoreHTML(b, l, i, 'add') + '</td></tr>';
      if (bal.open === key + ':cut' && l.status === 'cut') tr += '<tr class="bal-more"><td colspan="7">' + balMoreHTML(b, l, i, 'cut') + '</td></tr>';
      return tr;
    });
    return '<section class="tablebox"><div class="table-wrap"><table class="smart hours bal"><thead><tr><th scope="col">Предмет</th><th scope="col" class="col-minor">Преподаватель</th>' +
      '<th scope="col" class="n">План</th><th scope="col" class="n">Проведено</th><th scope="col" class="n">Ещё стоит</th><th scope="col" class="n" title="Больше нуля — пары нужно добавить, меньше нуля — снять">План − расписание</th><th scope="col">Что сделать</th></tr></thead><tbody>' +
      rows.join('') + '</tbody></table></div></section>';
  }
  function balDetailHTML(g) {
    var b = balanceFor(g);
    if (!b) return '<section class="empty"><p>' + (net.failed[g] ? 'Не удалось загрузить расписание группы ' + esc(g) + '.' : 'Загружаю расписание группы ' + esc(g) + '…') + '</p>' + (net.failed[g] ? '<button type="button" class="btn" data-bal-retry>Повторить</button>' : '') + '</section>';
    var t = b.total, facts = [];
    if (!b.plan) facts.push('Для группы ' + g + ' на листе «' + CFG.planSheet + '» нет учебного плана, сравнивать не с чем.');
    else if (b.plan.group !== g) facts.push('Учебный план взят из блока «' + b.plan.group + '».');
    if (b.theoryEnd) facts.push('Теория по графику учебного процесса — до ' + fmtDayMonth(b.theoryEnd) + '.');
    var pa = practiceAhead(g);
    if (pa && pa.start && !pa.now && pa.start <= b.until) facts.push('Практика начинается ' + fmtDayMonth(pa.start) + (pa.left ? ': до неё ' + count(pa.left, 'учебный день', 'учебных дня', 'учебных дней') + '.' : ': учебных дней до неё не осталось.') + (fireReason(g) === 'soon' ? ' Группа горит: пары нужно вычитать до практики.' : ''));
    if (b.covered && b.until && b.covered < b.until) facts.push('Расписание в таблице заполнено только до ' + fmtDayMonth(b.covered) + ': всё, что не стоит в таблице до ' + fmtDayMonth(b.until) + ', посчитано как недостающее.');
    if (b.passed) facts.push('Срок вычитки (' + fmtDayMonth(b.until) + ') уже прошёл: «Добавить» показывает, сколько пар так и не проведено. Чтобы вычитать их в дни практики, нажмите «Перед сессией» или выберите свою дату.');
    if (b.after) facts.push(pairsStand(b.after) + ' в таблице после ' + fmtDayMonth(b.until) + (b.until === b.sessionUntil ? ' — в сессию или каникулы по графику' : ' — в дни практики, сессии или каникул по графику') + ', они не засчитаны.' +
      (b.sessionUntil && b.until < b.sessionUntil ? ' Если эти пары правда пройдут, нажмите «Перед сессией».' : ''));
    if (b.onHoliday) facts.push(pairsStand(b.onHoliday) + ' в праздничные дни (' + b.holidayDays.map(fmtDayMonth).join(', ') + '), их не будет — они не засчитаны. Перенесите их на другие дни.');
    if (b.onPractice) facts.push(pairsStand(b.onPractice) + ' в дни практики по графику учебного процесса (первая — ' + fmtDayMonth(b.practiceFrom) + '). Они засчитаны. Если этих пар не будет, снимите их или поправьте график на вкладке «График».');
    facts.push(b.edits ? 'Учтены правки администраторов по этой группе: ' + b.edits + '. После каждой новой правки таблица пересчитывается сама.' : 'Правок администраторов по этой группе нет. Как только пару добавят, отменят или перенесут на сайте, числа пересчитаются сами.');
    if (b.guessed) facts.push('На листе группы ' + count(b.guessed, 'неделя', 'недели', 'недель') + ' без строки дат: даты взяты по номеру недели.');
    if (b.template) facts.push('У группы только недельный шаблон без дат: он посчитан на каждую неделю семестра.');
    if (b.source === 'cache') facts.push('Google-таблица не ответила, посчитано по сохранённой копии от ' + fmtStamp(b.at) + '.');
    var head = '<section class="hsum" aria-label="Итог вычитки">' +
      '<p class="hsum-line"><span translate="no">' + esc(g) + '</span>, ' + esc(halfName(b.half)) + ': ' +
      (!b.plan ? 'нет учебного плана.' : !t.add && !t.cut ? '<strong>часы сходятся</strong> с учебным планом.' :
        [t.add ? 'добавить <strong>' + esc(pairsAcc(t.add)) + '</strong> (не хватает ' + hoursWord(t.addH) + ' по ' + count(t.nAdd, 'предмету', 'предметам', 'предметам') + ')' : '',
          t.cut ? 'снять <strong>' + esc(pairsAcc(t.cut)) + '</strong> (лишние ' + hoursWord(t.cutH) + ' по ' + count(t.nCut, 'предмету', 'предметам', 'предметам') + ')' : ''].filter(Boolean).join(', ') + '.') + '</p>' +
      '<div class="bal-until"><label for="bal-until">Вычитать до</label><input type="date" id="bal-until" value="' + esc(b.until) + '"' + (b.half === 1 ? ' max="' + esc(isoOfUtc(utcOfIso(half2Of(g)) - DAY)) + '"' : '') + '>' +
      (b.theoryEnd && b.until !== b.theoryEnd ? '<button type="button" class="btn btn-small" data-bal-until-reset>До конца теории (' + esc(fmtDayMonth(b.theoryEnd)) + ')</button>' : '') +
      (b.sessionUntil && b.until !== b.sessionUntil ? '<button type="button" class="btn btn-small" data-bal-until-set="' + esc(b.sessionUntil) + '">Перед сессией (' + esc(fmtDayMonth(b.sessionUntil)) + ')</button>' : '') +
      '<span class="hint">' + (b.left ? 'осталось ' + count(b.left, 'учебный день', 'учебных дня', 'учебных дней') : 'учебные дни до срока закончились') + '</span></div>' +
      '<ul class="bal-facts">' + facts.map(function (f) { return '<li>' + esc(f) + '</li>'; }).join('') + '</ul>' +
      '<div class="bal-tools"><button type="button" class="btn btn-small" data-bal-csv="' + esc(g) + '">Скачать для Excel</button>' +
      (isAdmin() ? '' : '<span class="hint">Ставить и снимать пары отсюда может администратор после входа.</span>') + '</div></section>';
    return head + (b.lines.length ? balTableHTML(b) : '');
  }
  function balOverviewHTML() {
    var rows = DB.groups.map(function (g) { return { g: g, b: balanceFor(g) }; });
    rows.sort(function (x, y) {
      var a = x.b ? x.b.total.add : -1, c = y.b ? y.b.total.add : -1;
      return (isUrgent(y.g) ? 1 : 0) - (isUrgent(x.g) ? 1 : 0) || c - a || cmpRu(x.g, y.g);
    });
    var body = rows.map(function (x) {
      var b = x.b;
      if (!b) return '<tr><th scope="row" translate="no">' + esc(x.g) + '</th><td colspan="5" class="hint">' + (net.failed[x.g] ? 'не загрузилась' : 'загружаю…') + '</td></tr>';
      return '<tr><th scope="row"><span translate="no">' + esc(x.g) + '</span>' + (isUrgent(x.g) ? ' <span class="bal-tag">горит</span>' : '') + '</th>' +
        '<td class="n" data-label="Добавить">' + (b.total.add ? '<span class="bal-act is-bad">+' + b.total.add + '</span>' : '0') + '</td>' +
        '<td class="n" data-label="Снять">' + (b.total.cut ? '<span class="bal-act is-warn">−' + b.total.cut + '</span>' : '0') + '</td>' +
        '<td data-label="Вычитать до">' + (b.until ? esc(fmtDayMonth(b.until)) : '—') + '</td>' +
        '<td data-label="Таблица до">' + (b.covered ? esc(fmtDayMonth(b.covered)) : '—') + (b.covered && b.until && b.covered < b.until ? ' <span class="bal-act is-warn">не до конца</span>' : '') + '</td>' +
        '<td><button type="button" class="btn btn-small" data-bal-group="' + esc(x.g) + '">Открыть</button></td></tr>';
    });
    var loading = net.bulk ? '<p class="hint pad">Загружаю расписание всех групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…</p>' : '';
    return '<section class="tablebox"><div class="table-head"><h2>Все группы</h2><span class="count">пары, которые нужно добавить или снять до сессии</span>' +
      '<button type="button" class="btn btn-small" data-bal-csv="*">Скачать для Excel</button></div>' + loading +
      '<div class="table-wrap"><table class="smart hours bal-all"><thead><tr><th scope="col">Группа</th><th scope="col" class="n">Добавить пар</th><th scope="col" class="n">Снять пар</th>' +
      '<th scope="col">Вычитать до</th><th scope="col">Таблица заполнена до</th><th scope="col"><span class="sr-only">Действие</span></th></tr></thead><tbody>' + body.join('') + '</tbody></table></div></section>';
  }
  function renderBalance() {
    state.half = balanceHalf(balanceGroup() || DB.groups[0] || '');
    balMemo = {};
    var box = $('balance-body');
    if (!DB.groups.length) { box.innerHTML = '<section class="empty"><p>' + (net.listFailed ? 'Не удалось загрузить список групп.' : 'Загружаю список групп…') + '</p></section>'; return; }
    if (PLAN.status === 'loading') { box.innerHTML = '<section class="empty"><p>Загружаю учебный план…</p></section>'; return; }
    var urgent = burningGroups(), cur = state.balAll ? '' : balanceGroup();
    var opts = '<option value="">Выберите группу</option>' + DB.groups.map(function (g) { return '<option value="' + esc(g) + '"' + (g === cur ? ' selected' : '') + '>' + esc(g) + (isUrgent(g) ? ' — горит' : '') + '</option>'; }).join('');
    var html = '<section class="filters bal-filters" aria-label="Группа">' +
      '<div class="field"><label for="bal-group">Группа</label><select id="bal-group">' + opts + '</select></div>' +
      '<div class="field"><label for="bal-half">Период</label><select id="bal-half">' + [1, 2].map(function (h) { return '<option value="' + h + '"' + (balanceHalf(cur || DB.groups[0]) === h ? ' selected' : '') + '>' + (h ? h + ' семестр' : 'Весь учебный год') + '</option>'; }).join('') + '</select></div>' +
      '<div class="field"><span class="label">Сводка</span><button type="button" class="btn' + (state.balAll ? ' btn-primary' : '') + '" data-bal-all aria-pressed="' + (state.balAll ? 'true' : 'false') + '">Все группы</button></div></section>';
    if (urgent.length) {
      html += '<section class="bal-urgent" aria-labelledby="bal-urgent-h"><h2 id="bal-urgent-h">Горящие группы: срочная вычитка часов</h2><div class="bal-cards">' +
        urgent.map(function (g) { return balCardHTML(g, cur); }).join('') + '</div>' +
        '<p class="hint">Число на карточке — сколько пар добавить (+) или снять (−), чтобы к сроку вычитать часы ровно по учебному плану. Срок по умолчанию — последний день теории по графику учебного процесса, его можно поменять в группе. ' +
        'Группа загорается сама, если до её практики осталось ' + count(urgentDays(), 'день', 'дня', 'дней') + ' или меньше (или теория уже закончилась), а пары ещё не вычитаны. Постоянный список горящих групп — в config.js, поле urgentGroups.</p></section>';
    }
    html += state.balAll ? balOverviewHTML() : cur ? balDetailHTML(cur) : '<section class="empty"><p>Выберите группу.</p></section>';
    box.innerHTML = html;
  }
  function balLine(key) {
    var p = key.split('|'), b = balanceFor(p[0]);
    return b ? { b: b, l: b.lines[+p[1]], i: +p[1] } : null;
  }
  function balPut(arg) {
    var p = arg.split('|'), x = balLine(p[0] + '|' + p[1]);
    if (!x || !x.l) return;
    var l = x.l, g = x.b.group, day = p[2], start = p[3];
    var disc = l.names[0] || l.name, teacher = lineTeacher(l), rooms = {};
    DB.sheets[g].rows.forEach(function (r) { if (!r.ghost && l.names.indexOf(r.disc) >= 0 && r.room) rooms[r.room] = (rooms[r.room] || 0) + 1; });
    var room = Object.keys(rooms).sort(function (a, c) { return rooms[c] - rooms[a]; })[0] || '';
    var e = { id: '', kind: 'add', group: g, date: day, start: start, end: slotEndOf(start), sub: 0, disc: disc, teacher: teacher, room: room, note: 'вычитка часов', was: '' };
    Api.save([e], {}).then(function (saved) {
      mergeEdits(saved);
      HIST.status = 'idle';
      recompute();
      var nb = balanceFor(g), nl = nb && nb.lines.filter(function (y) { return y.name === l.name; })[0];
      if (!nl || nl.status !== 'add') bal.open = '';
      render();
      toast('Пара поставлена: ' + dayLabel(day) + ', ' + slotLabel(start) + '. ' + (nl && nl.status === 'add' ? 'Осталось добавить ' + pairsAcc(nl.pairs) + '.' : 'По этому предмету часы теперь сходятся.'));
    }, function (err) { toast(err.message, true); });
  }
  function balCut(arg) {
    var p = arg.split('|'), x = balLine(p[0] + '|' + p[1]);
    if (!x || !x.l) return;
    var g = x.b.group, day = p[2], start = p[3], sub = +p[4] || 0;
    var row = slotRows(g, day, start).filter(function (r) { return !r.ghost && x.l.names.indexOf(r.disc) >= 0 && (+r.sub || 0) === sub; })[0];
    if (!row) return;
    if (!confirm('Отменить пару «' + splitCode(row.disc).name + '» ' + dayLabel(day) + ', ' + slotLabel(start) + '?')) return;
    var after = function () {
      var nb = balanceFor(g), nl = nb && nb.lines.filter(function (y) { return y.name === x.l.name; })[0];
      if (!nl || nl.status !== 'cut') bal.open = '';
    };
    if (row.edit && row.edit.kind === 'add') { removeEdit(row.edit.id, function () { after(); render(); }); return; }
    var was = JSON.stringify(slotRows(g, day, start).filter(function (r) { return !r.ghost && !r.edit && (!sub || !r.sub || r.sub === sub); }).map(function (r) { return { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub }; }));
    var e = { id: row.edit ? row.edit.id : '', kind: 'cancel', group: g, date: day, start: start, end: row.end, sub: sub, note: 'вычитка: лишние часы', was: row.edit ? row.edit.was : was };
    Api.save([e], {}).then(function (saved) {
      mergeEdits(saved);
      HIST.status = 'idle';
      recompute();
      after();
      render();
      toast('Пара отменена. Убрать отмену можно в разделе «Изменения».');
    }, function (err) { toast(err.message, true); });
  }
  function csvCell(v) { var s = String(v == null ? '' : v); return /[";\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
  function balCsv(which) {
    var rows = [];
    if (which === '*') {
      rows.push(['Группа', 'Горит', 'Добавить пар', 'Снять пар', 'Не хватает часов', 'Лишних часов', 'Вычитать до', 'Таблица заполнена до']);
      DB.groups.forEach(function (g) {
        var b = balanceFor(g);
        if (b) rows.push([g, isUrgent(g) ? 'да' : '', b.total.add, b.total.cut, b.total.addH, b.total.cutH, b.until ? fmtShort(b.until) : '', b.covered ? fmtShort(b.covered) : '']);
      });
    } else {
      var b = balanceFor(which);
      if (!b) return;
      rows.push(['Группа', 'Предмет', 'Преподаватель', 'План, ч', 'Проведено, ч', 'Ещё стоит, ч', 'План − расписание, ч', 'Пар добавить', 'Пар снять', 'Что сделать']);
      b.lines.forEach(function (l) {
        var what = { add: 'добавить пары', cut: 'снять пары', ok: 'сходится', practice: 'практика по графику', noplan: 'в семестре нет часов по плану', extra: 'нет в учебном плане' }[l.status];
        rows.push([b.group, l.name, l.teachers.join(', ') || l.planTeacher, l.plan, l.done, l.future, l.plan || l.status === 'extra' ? l.gap : '', l.status === 'add' ? l.pairs : '', l.status === 'cut' ? l.pairs : '', what]);
      });
    }
    var text = '﻿' + rows.map(function (r) { return r.map(csvCell).join(';'); }).join('\r\n');
    var a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
    a.download = 'vychitka-' + (which === '*' ? 'vse-gruppy' : slug(which)) + '-' + todayISO() + '.csv';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }

  // ---------- PDF schedule view ----------
  var gcourse = 'all';
  function renderGrafik() {
    var box = $('grafik-body');
    Array.prototype.forEach.call($('g-course').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === gcourse ? 'true' : 'false');
    });
    if (GRAFIK.status === 'loading') { box.innerHTML = '<p class="empty">Загружаю график учебного процесса…</p>'; return; }
    if (GRAFIK.status === 'error') { box.innerHTML = '<p class="empty">Не удалось загрузить график учебного процесса. Обновите страницу. Если ошибка повторяется, проверьте, что на сайт загружена папка data с файлом grafik.json.</p>'; return; }
    var d = GRAFIK.data, weeks = d.weeks, cur = currentWeekIso(), ci = weeks.indexOf(cur), admin = isAdmin();
    var byCourse = {};
    Object.keys(d.groups).forEach(function (k) { var c = d.groups[k].course || 0; (byCourse[c] || (byCourse[c] = [])).push(k); });
    var MONTHS_SHORT = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];
    var months = [], prev = -1;
    weeks.forEach(function (w, i) {
      var m = new Date(utcOfIso(w) + 3 * DAY).getUTCMonth();
      if (m !== prev) { months.push({ m: m, n: 1 }); prev = m; } else months[months.length - 1].n++;
    });
    var head = '<thead><tr class="g-months"><th scope="col" rowspan="2" class="g-name">Группа</th>' + months.map(function (x) {
      return '<th scope="colgroup" colspan="' + x.n + '">' + (x.n > 1 ? MONTHS_SHORT[x.m] : '') + '</th>';
    }).join('') + '</tr><tr class="g-nums">' + weeks.map(function (w, i) {
      return '<th scope="col" class="' + (i === ci ? 'is-current' : '') + '" title="' + esc(fmtRange(w, isoOfUtc(utcOfIso(w) + 5 * DAY))) + '">' + (i + 1) + '</th>';
    }).join('') + '</tr></thead>';
    var body = Object.keys(byCourse).map(Number).sort().filter(function (c) { return gcourse === 'all' || String(c) === gcourse; }).map(function (c) {
      return '<tbody><tr class="g-course"><th scope="rowgroup" colspan="' + (weeks.length + 1) + '">' + (c ? c + ' курс' : 'Курс не указан') + '</th></tr>' +
        byCourse[c].map(function (k) {
          var g = d.groups[k], real = DB.groups.indexOf(k) >= 0 ? k : DB.groups.filter(function (x) { return lettersOf(x) === lettersOf(k) && courseOf(x) === g.course; })[0] ||
            (/\(\d+\)$/.test(k) && DB.groups.indexOf(lettersOf(k) + 'п-' + yearOf(k)) >= 0 ? lettersOf(k) + 'п-' + yearOf(k) : null);
          var name = '<th scope="row" class="g-name">' + (real ? '<button type="button" class="linklike" data-open-group="' + esc(real) + '" title="' + esc(g.name) + '">' + esc(k) + '</button>' : '<span title="' + esc(g.name) + '">' + esc(k) + '</span>') +
            (admin ? '<button type="button" class="pair-edit g-edit" data-gedit="' + esc(k) + '" aria-label="Изменить график группы ' + esc(k) + '">' + PENCIL + '</button>' : '') + '</th>';
          var gm = grafikFor(k);
          return '<tr>' + name + gm.g.days.map(function (days, i) {
            var o = gm.edited[i];
            var label = k + ', неделя ' + (i + 1) + ': ' + weekTypeText(days) + (o ? ' (изменено администратором' + (o.note ? ': ' + o.note : '') + ')' : '');
            return '<td class="' + (i === ci ? 'is-current' : '') + (o ? ' is-edited' : '') + '" title="' + esc(label) + '"' + (admin ? ' data-gk="' + esc(k) + '" data-wi="' + i + '"' : '') + '>' +
              days.split('').map(function (ch) { return '<i class="d-' + ch + '"></i>'; }).join('') + '</td>';
          }).join('') + '</tr>';
        }).join('') + '</tbody>';
    }).join('');
    var legend = Object.keys(WEEK_TYPES).map(function (c) { return '<span><i class="d-' + c + '"></i>' + esc(WEEK_TYPES[c]) + '</span>'; }).join('') +
      (GOV.list.length ? '<span><i class="edited-mark"></i>изменено администратором</span>' : '');
    var pdfs = (d.source || []).map(function (f) {
      var m = /(\d)-kurs/.exec(f);
      return '<a href="data/pdf/' + encodeURIComponent(f) + '" target="_blank" rel="noopener">' + (m ? m[1] + ' курс' : esc(f)) + ' (PDF)</a>';
    }).join('');
    box.innerHTML = '<p class="legend">' + legend + '</p>' +
      '<section class="tablebox"><div class="g-scroll"><table class="grafik">' + head + body + '</table></div></section>' +
      '<p class="hint">Каждая клетка — неделя, пять полосок в ней — дни с понедельника по пятницу. Текущая неделя обведена. ' +
      'Нажмите на группу, чтобы открыть её расписание.' + (admin ? ' Чтобы сдвинуть практику или каникулы, нажмите на клетку недели или на карандаш у группы.' : '') +
      ' Исходные файлы:</p><p class="foot-links">' + pdfs + '</p>';
  }

  // ---------- correcting the PDF schedule (ТЗ, этап 4: «даты практик могут сдвигаться») ----------
  var CODES = ['T', 'A', 'K', 'U', 'P', 'X', 'D', 'G'];
  function daysText(days) {
    return days.split('').map(function (c, j) { return ['пн', 'вт', 'ср', 'чт', 'пт'][j] + ' — ' + (c === '.' ? 'как в PDF' : WEEK_TYPES[c] || c); }).join(', ');
  }
  function openGrafikEditor(key, wi) {
    var d = GRAFIK.data;
    if (!d) return;
    var keys = Object.keys(d.groups).sort(cmpRu);
    if (keys.indexOf(key) < 0) key = keys[0];
    var ci = d.weeks.indexOf(currentWeekIso());
    if (wi == null || wi < 0) wi = ci >= 0 ? ci : 0;
    $('gf-group').innerHTML = keys.map(function (k) { return '<option value="' + esc(k) + '">' + esc(k) + '</option>'; }).join('');
    $('gf-group').value = key;
    var opts = '<option value=".">как в PDF</option>' + CODES.map(function (c) { return '<option value="' + c + '">' + esc(cap(WEEK_TYPES[c])) + '</option>'; }).join('');
    for (var j = 0; j < 5; j++) $('gf-d' + j).innerHTML = opts;
    $('gf-all').innerHTML = '<option value="">Выберите…</option>' + opts;
    $('gf-from').value = d.weeks[wi];
    $('gf-to').value = isoOfUtc(utcOfIso(d.weeks[wi]) + 4 * DAY);
    $('gf-note').value = '';
    $('gf-error').textContent = '';
    gfLoad();
    $('gf-dialog').showModal();
  }
  function gfWeeks() {
    var d = GRAFIK.data, from = $('gf-from').value, to = $('gf-to').value;
    if (!d || !from || !to || to < from) return [];
    var a = isoOfUtc(mondayOfUtc(utcOfIso(from)));
    return d.weeks.filter(function (w) { return w >= a && w <= to; });
  }
  // the day selects show the correction of the first chosen week, if it has one
  function gfLoad() {
    var key = $('gf-group').value, weeks = gfWeeks();
    var ov = weeks.length && GOV.map[key] ? GOV.map[key][weeks[0]] : null;
    var days = ov ? ov.days : '.....';
    for (var j = 0; j < 5; j++) $('gf-d' + j).value = days.charAt(j) || '.';
    if (ov) $('gf-note').value = ov.note || '';
    gfSummary();
  }
  function gfSummary() {
    var key = $('gf-group').value, weeks = gfWeeks(), gm = grafikFor(key), d = GRAFIK.data;
    if (!weeks.length) { $('gf-summary').innerHTML = '<p>Выберите даты внутри учебного года: ' + esc(d.title) + '.</p>'; return; }
    var lines = weeks.slice(0, 8).map(function (w) {
      var i = d.weeks.indexOf(w);
      return fmtRange(w, isoOfUtc(utcOfIso(w) + 4 * DAY)) + ': в PDF — ' + weekTypeText(gm.orig.days[i]) + (gm.edited[i] ? '; сейчас — ' + weekTypeText(gm.g.days[i]) : '');
    });
    $('gf-summary').innerHTML = '<p>Изменится ' + count(weeks.length, 'неделя', 'недели', 'недель') + ':</p><ul>' + lines.map(function (l) { return '<li>' + esc(l) + '</li>'; }).join('') + '</ul>' +
      (weeks.length > 8 ? '<p class="hint">И ещё ' + (weeks.length - 8) + '.</p>' : '');
  }
  function gfSave(reset) {
    var key = $('gf-group').value, weeks = gfWeeks(), days = '';
    for (var j = 0; j < 5; j++) days += $('gf-d' + j).value || '.';
    if (!weeks.length) { $('gf-error').textContent = 'Выберите даты недель внутри учебного года.'; return; }
    if (!reset && days === '.....') { $('gf-error').textContent = 'Выберите, что будет хотя бы в один день, или нажмите «Вернуть как в PDF».'; return; }
    var note = $('gf-note').value.trim();
    var items = weeks.map(function (w) { return { group: key, week: w, days: reset ? '' : days, note: note }; });
    var btn = reset ? $('gf-reset') : $('gf-save'), label = btn.textContent;
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    $('gf-error').textContent = '';
    Api.grafikSave(items).then(function (saved) {
      setGrafik(reset ? items : saved);
      HIST.status = 'idle';
      $('gf-dialog').close();
      recompute();
      render();
      toast(reset ? 'Недели вернулись к графику из PDF.' : 'График группы ' + key + ' изменён: ' + count(weeks.length, 'неделя', 'недели', 'недель') + '.');
    }, function (err) { $('gf-error').textContent = err.message; }).then(function () { btn.disabled = false; btn.textContent = label; });
  }

  // ---------- bulk replacement (ТЗ, этап 4: «преподаватель заболел, заменить все его пары на неделе») ----------
  var bulk = { kind: 'replace' };
  function openBulk(teacher) {
    ensureAll();
    var w = state.week || currentWeekIso();
    $('bulk-from').value = w;
    $('bulk-to').value = isoOfUtc(utcOfIso(w) + 5 * DAY);
    $('bulk-new').value = '';
    $('bulk-room').value = '';
    $('bulk-note').value = '';
    $('bulk-error').textContent = '';
    bulkFillTeachers(teacher || (state.mode === 'teacher' ? state.teacher : ''));
    bulkSetKind('replace');
    $('bulk-dialog').showModal();
  }
  function bulkFillTeachers(want) {
    var cur = want || $('bulk-teacher').value;
    $('bulk-teacher').innerHTML = '<option value="">Выберите преподавателя</option>' + DB.teachers.map(function (t) { return '<option value="' + esc(t) + '">' + esc(t) + '</option>'; }).join('');
    $('bulk-teacher').value = DB.teachers.indexOf(cur) >= 0 ? cur : '';
    datalistFill('dl-teacher', DB.teachers);
  }
  function bulkSetKind(kind) {
    bulk.kind = kind;
    Array.prototype.forEach.call($('bulk-kind').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === kind ? 'true' : 'false');
    });
    $('bulk-new-field').hidden = kind !== 'replace';
    $('bulk-room-field').hidden = kind !== 'replace';
    syncRoving();
    bulkPreview();
  }
  function bulkRows() {
    var t = $('bulk-teacher').value, from = $('bulk-from').value, to = $('bulk-to').value, out = [];
    if (!t || !from || !to) return out;
    Object.keys(DB.sheets).forEach(function (g) {
      DB.sheets[g].rows.forEach(function (r) { if (!r.ghost && r.disc && r.teacher === t && r.day >= from && r.day <= to) out.push(r); });
    });
    return out.sort(function (a, b) { return a.day < b.day ? -1 : a.day > b.day ? 1 : minutesOf(a.start) - minutesOf(b.start) || cmpRu(a.group, b.group); });
  }
  function bulkPreview() {
    var rows = bulkRows(), box = $('bulk-preview'), btn = $('bulk-save');
    if (net.bulk) box.innerHTML = '<p class="hint">Загружаю расписание всех групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…</p>';
    else if (!$('bulk-teacher').value) box.innerHTML = '<p class="hint">Выберите преподавателя и даты.</p>';
    else if (!rows.length) box.innerHTML = '<p class="hint">У этого преподавателя нет пар в выбранные дни.</p>';
    else {
      box.innerHTML = '<p>' + count(rows.length, 'пара', 'пары', 'пар') + ':</p><ul class="bulk-list">' + rows.slice(0, 40).map(function (r) {
        return '<li>' + ['Вс', 'Пн', 'Вт', 'Ср', 'Чт', 'Пт', 'Сб'][isoToDate(r.day).getDay()] + ', ' + esc(fmtDayMonth(r.day)) + ', ' + esc(slotLabel(r.start)) + ': <span class="grp" translate="no">' + esc(r.group) + '</span> ' +
          esc(splitCode(r.disc).name) + (r.room ? ', ' + esc(roomLabel(r.room)) : '') + (r.edit ? ' (уже с правкой)' : '') + '</li>';
      }).join('') + '</ul>' + (rows.length > 40 ? '<p class="hint">И ещё ' + (rows.length - 40) + '.</p>' : '');
    }
    btn.disabled = !rows.length || !!net.bulk;
    btn.textContent = (bulk.kind === 'replace' ? 'Заменить ' : 'Отменить ') + (rows.length ? count(rows.length, 'пару', 'пары', 'пар') : 'пары');
  }
  function bulkSave() {
    var rows = bulkRows(), teacher = $('bulk-teacher').value, newT = $('bulk-new').value.trim(), room = $('bulk-room').value.trim();
    var note = $('bulk-note').value.trim() || (bulk.kind === 'replace' ? 'замена преподавателя ' + teacher : 'пары преподавателя ' + teacher + ' отменены');
    if (bulk.kind === 'replace' && !newT) { $('bulk-error').textContent = 'Укажите, кто будет вести пары.'; $('bulk-new').focus(); return; }
    if (bulk.kind === 'replace' && newT === teacher) { $('bulk-error').textContent = 'Новый преподаватель совпадает с прежним.'; return; }
    var batch = 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), list = [], drop = [];
    rows.forEach(function (r) {
      var e = r.edit;
      var was = JSON.stringify((e ? r.was : [r]).map(function (x) { return { disc: x.disc, teacher: x.teacher, room: x.room, sub: x.sub }; }));
      var slot = { group: r.group, date: r.day, start: r.start, end: r.end, sub: r.sub || 0, note: note };
      if (e && e.kind === 'add') {
        if (bulk.kind === 'cancel') drop.push(e.id);
        else list.push(Object.assign({}, slot, { id: e.id, kind: 'add', disc: e.disc, teacher: newT, room: room || e.room }));
        return;
      }
      list.push(Object.assign({}, slot, { id: e ? e.id : '', was: e ? e.was : was, kind: bulk.kind, disc: r.disc, teacher: bulk.kind === 'replace' ? newT : '', room: bulk.kind === 'replace' ? room || r.room : '' }));
    });
    var btn = $('bulk-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    $('bulk-error').textContent = '';
    var chunks = [];
    for (var i = 0; i < list.length; i += 100) chunks.push(list.slice(i, i + 100));
    chunks.reduce(function (pr, chunk) {
      return pr.then(function () { return Api.save(chunk, { batch: batch, comment: note }).then(mergeEdits); });
    }, Promise.resolve()).then(function () {
      return drop.length ? Api.remove(drop, { comment: note }).then(function (removed) { setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok'); }) : null;
    }).then(function () {
      HIST.status = 'idle';
      $('bulk-dialog').close();
      recompute();
      render();
      toast((bulk.kind === 'replace' ? 'Заменено ' : 'Отменено ') + count(rows.length, 'пара', 'пары', 'пар') + '. Убрать всю замену можно в разделе «Изменения».');
    }, function (err) {
      $('bulk-error').textContent = err.message + (EDITS.list.some(function (x) { return x.batch === batch; }) ? ' Часть пар уже сохранена, её можно убрать в разделе «Изменения».' : '');
      recompute();
      render();
    }).then(function () { bulkPreview(); });
  }

  // ---------- changes view ----------
  var KIND_LABEL = { replace: 'Замена', cancel: 'Отмена', add: 'Новая пара' };
  var ACTION_LABEL = { create: 'Внесена правка', update: 'Правка изменена', delete: 'Правка убрана', revert: 'Откат', import: 'Загружен из документа' };
  function slotLabel(start) {
    var s = DB.slotByStart[normTime(start)];
    return (s ? s.n + ' пара, ' + s.start + '–' + s.end : 'начало в ' + start);
  }
  function snapText(json) {
    if (!json) return '';
    var o;
    try { o = JSON.parse(json); } catch (e) { return ''; }
    if (o.kind === 'cancel') return 'пара отменена' + (o.note ? ' (' + o.note + ')' : '');
    return [o.disc, o.teacher, o.room ? roomLabel(o.room) : '', o.sub ? o.sub + ' подгруппа' : ''].filter(Boolean).join(', ') + (o.note ? ' (' + o.note + ')' : '');
  }
  function baseRowsAt(e) {
    var sheet = DB.sheets[e.group];
    if (!sheet) return null;
    var row = (sheet.byDay[e.date] || []).filter(function (r) { return r.edit && r.edit.id === e.id; });
    var was = [];
    row.forEach(function (r) { (r.was || []).forEach(function (w) { if (was.indexOf(w) < 0) was.push(w); }); });
    return was;
  }
  function editWas(e) {
    var was = baseRowsAt(e);
    if (was && was.length) return wasText(was);
    try { var w = JSON.parse(e.was || '[]'); if (Array.isArray(w) && w.length) return wasText(w.map(function (x) { return { disc: x.disc || '', teacher: x.teacher || '', room: x.room || '' }; })); } catch (err) { /* old format */ }
    return '';
  }
  function renderHistory() {
    var box = $('history-body'), today = todayISO(), admin = isAdmin();
    var hg = $('h-group');
    var groups = uniq(EDITS.list.map(function (e) { return e.group; }).concat(HIST.items.map(function (h) { return h.group; }))).sort(cmpRu);
    var want = HIST.group;
    hg.innerHTML = '<option value="">Все группы</option>' + groups.map(function (g) { return '<option value="' + esc(g) + '"' + (g === want ? ' selected' : '') + '>' + esc(g) + '</option>'; }).join('');
    if (Api.mode === 'off') {
      box.innerHTML = '<div class="empty"><h2>Изменений пока нет</h2><p>Здесь появятся замены и отмены пар, когда администраторы начнут вносить их на сайте.</p>' +
        '<p class="hint">Для администратора: сервер правок ещё не подключён. Как подключить, написано в README.md, раздел «Админка».</p></div>';
      return;
    }
    Array.prototype.forEach.call($('h-type').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === histType ? 'true' : 'false');
    });
    var list = EDITS.list.filter(function (e) { return (!want || e.group === want) && (histType === 'all' || recordType(e) === histType); });
    var upcoming = list.filter(function (e) { return e.date >= today; }).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : minutesOf(a.start) - minutesOf(b.start); });
    var past = list.filter(function (e) { return e.date < today; }).sort(function (a, b) { return a.date > b.date ? -1 : 1; });
    function editCard(e) {
      var was = editWas(e);
      var now = e.kind === 'cancel' ? 'Пары не будет' : [splitCode(e.disc).name, e.teacher, e.room ? roomLabel(e.room) : '', e.sub ? e.sub + ' подгруппа' : ''].filter(Boolean).join(', ');
      var inBatch = e.batch ? EDITS.list.filter(function (x) { return x.batch === e.batch; }).length : 0;
      return '<li class="change k-' + e.kind + '"><div class="change-when"><strong>' + esc(cap(WEEKDAYS[isoToDate(e.date).getDay()])) + ', ' + esc(fmtDayMonth(e.date)) + '</strong><span>' + esc(slotLabel(e.start)) + '</span></div>' +
        '<div class="change-what"><p><span class="grp" translate="no">' + esc(e.group) + '</span> <span class="badge b-' + e.kind + '">' + KIND_LABEL[e.kind] + '</span>' +
        (inBatch > 1 ? ' <span class="badge b-batch">массовая замена, ' + count(inBatch, 'пара', 'пары', 'пар') + '</span>' : '') + '</p>' +
        '<p class="change-now">' + esc(now) + '</p>' + (was ? '<p class="was">Было: ' + esc(was) + '</p>' : '') + (e.note ? '<p class="was">Причина: ' + esc(e.note) + '</p>' : '') +
        sourceHTML({ edit: e, groups: [e.group] }) + '</div>' +
        (admin ? '<div class="change-act"><button type="button" class="btn btn-small" data-open-edit="' + esc(e.id) + '">Изменить</button><button type="button" class="btn btn-quiet btn-small" data-remove-edit="' + esc(e.id) + '">Убрать</button>' +
          (inBatch > 1 ? '<button type="button" class="btn btn-quiet btn-small" data-remove-batch="' + esc(e.batch) + '">Убрать всю замену</button>' : '') + '</div>' : '') + '</li>';
    }
    // ТЗ, этап 4: правка важнее таблицы, пока её не сняли; если таблица тоже изменилась — конфликт
    var conflicts = [];
    if (admin) {
      var seenC = {};
      Object.keys(DB.sheets).forEach(function (g) {
        DB.sheets[g].rows.forEach(function (r) {
          if (r.conflict && r.edit && !r.away && !seenC[r.edit.id] && (!want || r.group === want)) { seenC[r.edit.id] = true; conflicts.push(r); }
        });
      });
    }
    var html = '';
    if (EDITS.status === 'loading') html += '<p class="hint">Загружаю правки…</p>';
    if (EDITS.status === 'error') html += '<p class="hint">Не удалось загрузить правки с сервера.</p>';
    if (EDITS.status === 'cache') html += '<p class="hint">Сервер правок не ответил, показана сохранённая копия от ' + esc(fmtStamp(EDITS.at)) + '.</p>';
    if (conflicts.length) {
      html += '<section class="tablebox conflicts"><div class="table-head"><h2>Конфликты с таблицей</h2><span class="count">' + count(conflicts.length, 'правка', 'правки', 'правок') + '</span></div>' +
        '<ol class="changes">' + conflicts.map(function (r) {
          return '<li class="change"><div class="change-when"><strong>' + esc(cap(WEEKDAYS[isoToDate(r.day).getDay()])) + ', ' + esc(fmtDayMonth(r.day)) + '</strong><span>' + esc(slotLabel(r.start)) + '</span></div>' +
            '<div class="change-what"><p><span class="grp" translate="no">' + esc(r.group) + '</span> Правка: ' + esc(r.edit.kind === 'cancel' ? 'пара отменена' : wasText([r])) + '</p>' +
            '<p class="was">Когда вносили правку, в таблице было: ' + esc(wasText(r.conflict.then) || 'пары не было') + '. Сейчас: ' + esc(wasText(r.conflict.now) || 'пары нет') + '.</p></div>' +
            '<div class="change-act"><button type="button" class="btn btn-small" data-keep-edit="' + esc(r.edit.id) + '">Оставить правку</button><button type="button" class="btn btn-quiet btn-small" data-drop-edit="' + esc(r.edit.id) + '">Убрать правку</button></div></li>';
        }).join('') + '</ol><div class="pad bal-tools"><button type="button" class="btn btn-small" data-drop-many="conflict">Взять из таблицы во всех (' + conflicts.length + ')</button>' +
        '<span class="hint">Здесь группы, чьё расписание открыто на этом компьютере. Чтобы проверить все, нажмите «Обновить и импортировать».</span></div></section>';
    }
    // edits that the table already repeats: the site does not lay them over it, they only clutter the list
    var inTable = admin ? EDITS.list.filter(function (x) { return IN_TABLE[x.id] && (!want || x.group === want); }) : [];
    if (inTable.length) {
      html += '<section class="tablebox"><div class="table-head"><h2>Уже внесено в таблицу</h2><span class="count">' + count(inTable.length, 'правка', 'правки', 'правок') + '</span>' +
        '<button type="button" class="btn btn-small" data-drop-many="intable">Убрать все</button></div>' +
        '<p class="pad hint">В таблице колледжа теперь стоит то же, что в этих правках, поэтому сайт показывает пары прямо из таблицы и не считает их дважды. Сами правки можно убрать.</p>' +
        '<ol class="changes">' + inTable.map(function (x) {
          return '<li class="change"><div class="change-when"><strong>' + esc(cap(WEEKDAYS[isoToDate(x.date).getDay()])) + ', ' + esc(fmtDayMonth(x.date)) + '</strong><span>' + esc(slotLabel(normTime(x.start))) + '</span></div>' +
            '<div class="change-what"><p><span class="grp" translate="no">' + esc(x.group) + '</span> ' + esc(x.kind === 'cancel' ? 'пара отменена' : wasText([x])) + '</p></div>' +
            '<div class="change-act"><button type="button" class="btn btn-quiet btn-small" data-drop-edit="' + esc(x.id) + '">Убрать правку</button></div></li>';
        }).join('') + '</ol></section>';
    }
    html += '<section class="tablebox"><div class="table-head"><h2>Ближайшие замены и отмены</h2><span class="count">' + count(upcoming.length, 'правка', 'правки', 'правок') + '</span></div>' +
      (upcoming.length ? '<ol class="changes">' + upcoming.map(editCard).join('') + '</ol>' : '<p class="pad">На ближайшие дни изменений нет: пары идут по основному расписанию.</p>') + '</section>';
    if (past.length) html += '<details class="tablebox past-edits"><summary>Прошедшие правки (' + past.length + ')</summary><ol class="changes">' + past.map(editCard).join('') + '</ol></details>';
    var items = HIST.items.filter(function (h) { return (!want || h.group === want) && (histType === 'all' || historyType(h) === histType); });
    html += '<section class="tablebox"><div class="table-head"><h2>Журнал правок</h2><span class="count">' +
      (HIST.status === 'loading' ? 'загружаю…' : HIST.status === 'error' ? 'не загрузился' : count(items.length, 'запись', 'записи', 'записей')) + '</span>' +
      (HIST.status === 'error' ? '<button type="button" class="btn btn-small" data-history-retry>Повторить</button>' : '') + '</div>';
    if (items.length) {
      html += '<ol class="log">' + items.map(function (h) {
        var t = new Date(h.time), target = h.target || 'edit', before, after, what, tag;
        if (target === 'grafik') {
          var gb = parseJSON(h.before), ga = parseJSON(h.after);
          before = gb ? daysText(gb.days) : ''; after = ga ? daysText(ga.days) + (ga.note ? ' (' + ga.note + ')' : '') : 'как в PDF';
          tag = '<span class="rt rt-grafik">График</span>';
          what = 'График учебного процесса, <span class="grp" translate="no">' + esc(h.group) + '</span>, неделя с ' + esc(fmtDayMonth(h.date)) + '.';
        } else if (target === 'catalog') {
          var cb = parseJSON(h.before), ca = parseJSON(h.after);
          before = cb ? cb.name : ''; after = ca ? ca.name : '';
          tag = '<span class="rt rt-grafik">Справочник</span>';
          what = (ca || cb || {}).type === 'teacher' ? 'Преподаватель.' : 'Предмет.';
        } else if (target === 'plan') {
          var pb = parseJSON(h.before), pa = parseJSON(h.after);
          before = pb ? pb.s1 + ' / ' + pb.s2 + ' ч.' : 'исходный учебный план';
          after = pa ? pa.s1 + ' / ' + pa.s2 + ' ч.' : 'исходный учебный план';
          tag = '<span class="rt rt-grafik">Учебный план</span>';
          what = '<span translate="no">' + esc(h.group) + '</span>: ' + esc((pa || pb || {}).name || '') + '.';
        } else if (target === 'exams') {
          var xb = parseJSON(h.before) || {}, xa = parseJSON(h.after) || {};
          before = count(+xb.count || 0, 'запись', 'записи', 'записей'); after = count(+xa.count || 0, 'запись', 'записи', 'записей') + (xa.source ? ' из «' + xa.source + '»' : '');
          tag = '<span class="rt rt-grafik">Экзамены</span>';
          what = 'Список экзаменов и зачётов.';
        } else {
          before = snapText(h.before); after = snapText(h.after);
          tag = typeTag(historyType(h));
          what = '<span class="grp" translate="no">' + esc(h.group) + '</span> ' + esc(fmtDayMonth(h.date)) + ', ' + esc(slotLabel(h.start)) + '.';
        }
        return '<li><time datetime="' + esc(h.time) + '">' + esc(isNaN(t) ? h.time : fmtStamp(t.getTime())) + '</time>' +
          '<p>' + tag + ' <strong>' + esc(ACTION_LABEL[h.action] || h.action) + '</strong>' + (h.name ? ', ' + esc(h.name) : '') + '. ' + what + '</p>' +
          (before ? '<p class="was">Было: ' + esc(before) + '</p>' : '') + (after ? '<p class="change-now">Стало: ' + esc(after) + '</p>' : '') +
          (h.comment ? '<p class="was">Комментарий: ' + esc(h.comment) + '</p>' : '') +
          (admin ? '<p class="entry-acts"><button type="button" class="btn btn-quiet btn-small" data-revert="' + esc(h.id) + '">Откатить к «было»</button></p>' : '') + '</li>';
      }).join('') + '</ol>';
      if (!HIST.full) html += '<p class="pad hint">Кто внёс правку, видно после входа администратора.</p>';
    } else if (HIST.status === 'ok') html += '<p class="pad">Записей пока нет.</p>';
    html += '</section>' + tableJournalHTML(want);
    box.innerHTML = html;
  }
  function parseJSON(v) { try { return v ? JSON.parse(v) : null; } catch (e) { return null; } }
  // ТЗ, этап 2: «сравнение с прошлой версией и журнал „что изменилось“» по импортам таблицы
  function tableJournalHTML(want) {
    var items = TJ.items.filter(function (j) { return !want || j.group === want; });
    var head = '<section class="tablebox"><div class="table-head"><h2>Изменения в базе расписания</h2><span class="count">' +
      (TJ.status === 'loading' ? 'загружаю…' : TJ.status === 'error' ? 'не загрузился' : count(items.length, 'запись', 'записи', 'записей')) + '</span></div>';
    if (!items.length) {
      return head + '<p class="pad">' + (TJ.status === 'ok' ? 'Изменений пока нет. Журнал пополняется, когда администратор нажимает «Обновить и импортировать»: сайт сравнивает таблицу с прошлым импортом.' : '') + '</p></section>';
    }
    var byTime = {}, order = [];
    items.forEach(function (j) { if (!byTime[j.time]) { byTime[j.time] = []; order.push(j.time); } byTime[j.time].push(j); });
    return head + order.slice(0, 20).map(function (t) {
      var list = byTime[t], warn = list.filter(function (j) { return j.kind === 'warning'; }), ch = list.filter(function (j) { return j.kind !== 'warning'; });
      var at = Date.parse(t);
      return '<div class="tj"><p class="tj-head"><strong>Импорт ' + esc(isNaN(at) ? t : fmtStamp(at)) + '</strong>' + (list[0].name ? ', ' + esc(list[0].name) : '') + ': ' +
        count(ch.length, 'изменение', 'изменения', 'изменений') + (warn.length ? ', предупреждений: ' + warn.length : '') + '</p>' +
        (warn.length ? '<ul class="tj-warn">' + warn.map(function (j) { return '<li>' + esc(j.note) + '</li>'; }).join('') + '</ul>' : '') +
        (ch.length ? '<ul>' + ch.slice(0, 60).map(function (j) {
          return '<li><span class="grp" translate="no">' + esc(j.group) + '</span> ' + esc(fmtDayMonth(j.date)) + ', ' + esc(slotLabel(j.start)) + (j.sub ? ', ' + j.sub + ' подгруппа' : '') + ': ' + esc(changeText(j)) + '</li>';
        }).join('') + '</ul>' + (ch.length > 60 ? '<p class="hint">И ещё ' + (ch.length - 60) + '.</p>' : '') : '') + '</div>';
    }).join('') + (!TJ.full ? '<p class="pad hint">Предупреждения парсера видны вошедшим администраторам.</p>' : '') + '</section>';
  }

  // ---------- exams view (ТЗ БП 5.7 with the SmartTable rules of БП 1.5) ----------
  function examItems() {
    var out = [];
    EX.records.forEach(function (r) {
      out.push({ rec: r, date: r.date, time: r.time, retake: false });
      if (r.retakeDate) out.push({ rec: r, date: r.retakeDate, time: r.retakeTime, retake: true });
    });
    return out;
  }
  function byWhen(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    var ta = a.time || '99:99', tb = b.time || '99:99';
    if (ta !== tb) return ta < tb ? -1 : 1;
    return cmpRu(a.rec.group, b.rec.group) || cmpRu(a.rec.discipline, b.rec.discipline);
  }
  function statusOf(it, today) { return it.date > today ? 'future' : it.date === today ? 'today' : 'past'; }
  // 1st half of the academic year: September to January (with the winter session); 2nd: February to August
  function halfOf(iso) { var m = isoParts(iso).m; return m >= 9 || m === 1 ? '1' : '2'; }
  function matchesExam(it, today) {
    var r = it.rec;
    if (ex.group.length && ex.group.indexOf(r.group) < 0) return false;
    if (ex.teacher.length && !splitTeachers(r.teacher).some(function (t) { return ex.teacher.indexOf(t) >= 0; })) return false;
    if (ex.disc.length && ex.disc.indexOf(r.discipline) < 0) return false;
    if (ex.type !== 'all' && ex.type !== 'retake' && TYPE_KEY[r.type] !== ex.type) return false;
    if (ex.half !== 'all' && (it.date < half2Of(r.group) ? '1' : '2') !== ex.half) return false;
    var st = statusOf(it, today);
    if (ex.status === 'future' && st === 'past') return false;
    if (ex.status === 'past' && st !== 'past') return false;
    if (ex.from && it.date < ex.from) return false;
    if (ex.to && it.date > ex.to) return false;
    var q = ex.search.trim().toLowerCase().replace(/ё/g, 'е');
    if (q && (r.discipline + ' ' + r.teacher + ' ' + r.room + ' ' + r.group + ' ' + r.type).toLowerCase().replace(/ё/g, 'е').indexOf(q) < 0) return false;
    return true;
  }
  var EX_COLUMNS = [
    { key: 'num', label: '№' },
    { key: 'group', label: 'Группа', sort: function (it) { return it.rec.group; } },
    { key: 'disc', label: 'Дисциплина', sort: function (it) { return it.rec.discipline; } },
    { key: 'teacher', label: 'Преподаватель', sort: function (it) { return it.rec.teacher; } },
    { key: 'date', label: 'Дата', sort: function (it) { return it.date + ' ' + (it.time || '99:99'); } },
    { key: 'time', label: 'Время', minor: true },
    { key: 'room', label: 'Аудитория', minor: true, sort: function (it) { return it.rec.room; } },
    { key: 'type', label: 'Тип', sort: function (it) { return it.retake ? 'Пересдача' : it.rec.type; } },
    { key: 'status', label: 'Статус', sort: function (it) { return it.date; } }
  ];
  // Column visibility and order chosen by the viewer, kept in this browser (ТЗ БП 1.5 «Настроить колонки»)
  var colPrefs = null;
  try {
    var cp = JSON.parse(localStorage.getItem(COLS_KEY) || 'null');
    if (cp && Array.isArray(cp.order) && Array.isArray(cp.hidden)) colPrefs = cp;
  } catch (e) { /* storage unavailable */ }
  function colByKey(k) { return EX_COLUMNS.filter(function (c) { return c.key === k; })[0]; }
  function columnOrder() {
    var keys = EX_COLUMNS.map(function (c) { return c.key; });
    if (!colPrefs) return keys;
    return colPrefs.order.filter(function (k) { return keys.indexOf(k) >= 0; })
      .concat(keys.filter(function (k) { return colPrefs.order.indexOf(k) < 0; }));
  }
  function isHidden(k) { return !!colPrefs && colPrefs.hidden.indexOf(k) >= 0 && k !== 'disc'; }
  function columnsInUse() {
    return columnOrder().filter(function (k) { return !isHidden(k) && !(k === 'group' && ex.group.length === 1); }).map(colByKey);
  }
  function saveColPrefs() {
    try {
      if (colPrefs) localStorage.setItem(COLS_KEY, JSON.stringify(colPrefs));
      else localStorage.removeItem(COLS_KEY);
    } catch (e) { /* storage unavailable */ }
  }
  function sortItems(items) {
    var o = ex.ordering, desc = o.charAt(0) === '-', key = desc ? o.slice(1) : o;
    var col = EX_COLUMNS.filter(function (c) { return c.key === key && c.sort; })[0];
    if (!col) return items.sort(byWhen);
    return items.sort(function (a, b) { return (desc ? -1 : 1) * (cmpRu(col.sort(a), col.sort(b)) || byWhen(a, b)); });
  }
  function examCell(col, it, i, today) {
    var r = it.rec, v;
    switch (col.key) {
      case 'num': return '<td class="num" data-label="№">' + i + '</td>';
      case 'group': return '<td data-label="Группа"><span class="grp" translate="no">' + esc(r.group) + '</span></td>';
      case 'disc': return '<td class="title" data-label="Дисциплина">' + (isAdmin() ? '<button type="button" class="btn btn-small exam-edit" data-exam-edit="' + EX.records.indexOf(r) + '" aria-label="Изменить: ' + esc(r.discipline) + '">' + PENCIL + 'Изменить</button> ' : '') + discHTML(r.discipline) + (it.retake ? ' (Пересдача)' : '') + '</td>';
      case 'teacher': return '<td data-label="Преподаватель">' + esc(r.teacher || '—') + '</td>';
      case 'date': return '<td class="mono" data-label="Дата">' + fmtShort(it.date) + '</td>';
      case 'time': return '<td class="mono col-minor" data-label="Время">' + esc(it.time || '—') + '</td>';
      case 'room': return '<td class="col-minor" data-label="Аудитория">' + esc(r.room || '—') + '</td>';
      case 'type':
        v = it.retake ? '<span class="badge b-retake">Пересдача</span>' : '<span class="badge b-' + (TYPE_KEY[r.type] || 'credit') + '">' + esc(r.type) + '</span>';
        return '<td data-label="Тип">' + v + '</td>';
      case 'status':
        var st = statusOf(it, today);
        return '<td data-label="Статус"><span class="badge s-' + st + '">' + (st === 'future' ? 'Будущий' : st === 'today' ? 'Сегодня' : 'Прошедший') + '</span></td>';
    }
    return '<td></td>';
  }
  function renderTable(tableEl, pagerEl, items, pageKey, today) {
    var cols = columnsInUse();
    var size = ex.pageSize, pages = Math.max(1, Math.ceil(items.length / size));
    if (ex[pageKey] > pages) ex[pageKey] = pages;
    var page = ex[pageKey], from = (page - 1) * size, shown = items.slice(from, from + size);
    var o = ex.ordering, desc = o.charAt(0) === '-', okey = desc ? o.slice(1) : o;
    var head = '<thead><tr>' + cols.map(function (c) {
      var cls = c.minor ? ' class="col-minor"' : '';
      if (!c.sort) return '<th scope="col"' + cls + '>' + c.label + '</th>';
      var on = okey === c.key, arrow = on ? (desc ? '▼' : '▲') : '';
      var aria = on ? (desc ? 'descending' : 'ascending') : 'none';
      return '<th scope="col"' + cls + ' aria-sort="' + aria + '"><button type="button" data-sort="' + c.key + '">' + c.label +
        '<span class="arrow" aria-hidden="true">' + arrow + '</span></button></th>';
    }).join('') + '</tr></thead>';
    var body = '<tbody>' + shown.map(function (it, i) {
      return '<tr class="' + (statusOf(it, today) === 'past' ? 'is-past' : '') + '">' + cols.map(function (c) { return examCell(c, it, from + i + 1, today); }).join('') + '</tr>';
    }).join('') + '</tbody>';
    tableEl.innerHTML = head + body;
    pagerEl.innerHTML = '<span class="info">Показано ' + (items.length ? (from + 1) + '–' + (from + shown.length) : 0) + ' из ' + items.length + '</span>' +
      '<label>На странице <select data-size>' + [10, 20, 50, 100].map(function (n) {
        return '<option value="' + n + '"' + (n === size ? ' selected' : '') + '>' + n + '</option>';
      }).join('') + '</select></label>' +
      '<button type="button" class="btn btn-small" data-page="' + pageKey + ':' + (page - 1) + '"' + (page <= 1 ? ' disabled' : '') + '>← Предыдущие ' + size + '</button>' +
      '<button type="button" class="btn btn-small" data-page="' + pageKey + ':' + (page + 1) + '"' + (page >= pages ? ' disabled' : '') + '>Следующие ' + size + ' →</button>';
  }
  // ---------- multi-select dropdowns for group, teacher and discipline ----------
  var MS = {};
  function msInit(id) {
    var el = $(id), key = el.getAttribute('data-key');
    el.innerHTML = '<button type="button" class="ms-btn" aria-haspopup="true" aria-expanded="false" aria-labelledby="' + id + '-label ' + id + '-value"><span id="' + id + '-value"></span></button>' +
      '<div class="ms-pop" hidden><input type="search" name="find_' + key + '" placeholder="Найти в списке…" aria-label="Найти в списке" autocomplete="off">' +
      '<ul class="ms-list"></ul><div class="ms-foot"><button type="button" class="btn btn-quiet btn-small" data-ms-clear>Сбросить</button>' +
      '<button type="button" class="btn btn-small" data-ms-close>Готово</button></div></div>';
    var btn = el.querySelector('.ms-btn'), pop = el.querySelector('.ms-pop'), find = pop.querySelector('input');
    MS[key] = { el: el, btn: btn, pop: pop, list: pop.querySelector('ul'), options: [] };
    btn.addEventListener('click', function () { msOpen(key, pop.hidden); });
    find.addEventListener('input', function () {
      var q = find.value.trim().toLowerCase();
      Array.prototype.forEach.call(MS[key].list.children, function (li) { li.hidden = q && li.textContent.toLowerCase().indexOf(q) < 0; });
    });
    MS[key].list.addEventListener('change', function () {
      ex[key] = Array.prototype.filter.call(MS[key].list.querySelectorAll('input'), function (cb) { return cb.checked; }).map(function (cb) { return cb.value; });
      afterMsChange(key);
    });
    pop.addEventListener('click', function (e) {
      if (e.target.closest('[data-ms-clear]')) { ex[key] = []; afterMsChange(key); }
      if (e.target.closest('[data-ms-close]')) { msOpen(key, false); btn.focus(); }
    });
  }
  function afterMsChange(key) {
    ex.page = 1; ex.rpage = 1;
    if (key === 'group') {
      ex.disc = ex.disc.filter(function (d) { return discOptions().indexOf(d) >= 0; });
      msFill('disc', discOptions());
    }
    msSync(key);
    render();
  }
  function msOpen(key, open) {
    Object.keys(MS).forEach(function (k) {
      var on = k === key && open;
      MS[k].pop.hidden = !on;
      MS[k].btn.setAttribute('aria-expanded', on ? 'true' : 'false');
    });
    if (open) MS[key].pop.querySelector('input').focus();
  }
  function msFill(key, options) {
    var m = MS[key], list = options.slice();
    ex[key].forEach(function (v) { if (list.indexOf(v) < 0) list.unshift(v); });
    m.options = list;
    m.list.innerHTML = list.map(function (v) {
      return '<li><label><input type="checkbox" value="' + esc(v) + '"><span>' + esc(v) + '</span></label></li>';
    }).join('') || '<li class="hint">Список пуст</li>';
    msSync(key);
  }
  function msSync(key) {
    var m = MS[key], sel = ex[key];
    Array.prototype.forEach.call(m.list.querySelectorAll('input'), function (cb) { cb.checked = sel.indexOf(cb.value) >= 0; });
    var text = !sel.length ? m.el.getAttribute('data-all') : sel.length <= 2 ? sel.join(', ') : 'Выбрано ' + sel.length + ' ' + m.el.getAttribute('data-many');
    $(m.el.id + '-value').textContent = text;
    m.btn.title = sel.join(', ');
  }
  document.addEventListener('click', function (e) {
    if (!e.target.closest('.ms')) Object.keys(MS).forEach(function (k) { if (!MS[k].pop.hidden) msOpen(k, false); });
  });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    Object.keys(MS).forEach(function (k) { if (!MS[k].pop.hidden) { msOpen(k, false); MS[k].btn.focus(); } });
    if ($('e-sheet').classList.contains('open')) toggleSheet(false);
  });
  function discOptions() {
    return uniq(EX.records.filter(function (r) { return !ex.group.length || ex.group.indexOf(r.group) >= 0; }).map(function (r) { return r.discipline; })).sort(cmpRu);
  }
  function fillExamOptions() {
    var recs = EX.records;
    msFill('group', uniq(recs.map(function (r) { return r.group; })).sort(cmpRu));
    msFill('teacher', uniq([].concat.apply([], recs.map(function (r) { return splitTeachers(r.teacher); }))).sort(cmpRu));
    msFill('disc', discOptions());
  }
  function activeFilterCount() {
    var d = exDefaults();
    return EX_FILTER_KEYS.filter(function (k) {
      if (k === 'search') return false;
      return Array.isArray(ex[k]) ? ex[k].length > 0 : ex[k] !== d[k];
    }).length;
  }

  // ---------- favourite filter sets (ТЗ БП 1.5 «Избранные фильтры») ----------
  function loadFavs() {
    try { var v = JSON.parse(localStorage.getItem(FAVS_KEY) || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; }
  }
  function saveFavs(list) { try { localStorage.setItem(FAVS_KEY, JSON.stringify(list)); } catch (e) { /* storage unavailable */ } }
  function renderFavs() {
    var favs = loadFavs();
    $('e-favs').innerHTML = favs.length ? favs.map(function (f, i) {
      return '<span class="fav"><button type="button" data-fav="' + i + '" title="Применить набор">' + esc(f.name) + '</button>' +
        '<button type="button" data-fav-del="' + i + '" aria-label="Удалить набор «' + esc(f.name) + '»">×</button></span>';
    }).join('') : '<span class="hint">пока нет</span>';
  }

  function renderExams() {
    var today = todayISO();
    Array.prototype.forEach.call($('ex-page').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === state.exPage ? 'true' : 'false');
    });
    $('ex-list').hidden = state.exPage !== 'list';
    $('ex-book').hidden = state.exPage !== 'book';
    $('ex-teacher').hidden = state.exPage !== 'teacher';
    $('e-import').hidden = !isAdmin();
    $('e-xreset').hidden = !isAdmin();
    if (state.exPage === 'book') { renderBook(); return; }
    if (state.exPage === 'teacher') { renderTeacherExams(); return; }
    Object.keys(MS).forEach(msSync);
    $('e-from').value = ex.from;
    $('e-to').value = ex.to;
    if (document.activeElement !== $('e-search')) $('e-search').value = ex.search;
    [['e-type', ex.type], ['e-status', ex.status], ['e-half', ex.half]].forEach(function (pair) {
      Array.prototype.forEach.call($(pair[0]).querySelectorAll('button'), function (b) {
        b.setAttribute('aria-checked', b.getAttribute('data-value') === pair[1] ? 'true' : 'false');
      });
    });
    var nActive = activeFilterCount();
    $('e-reset').hidden = !nActive;
    $('e-fcount').textContent = nActive ? String(nActive) : '';

    var emptyEl = $('e-empty');
    $('e-main').hidden = true;
    $('e-retakes').hidden = true;
    emptyEl.hidden = true;
    if (EX.status === 'loading') { emptyEl.innerHTML = '<p>Загружаю расписание экзаменов из базы расписания…</p>'; emptyEl.hidden = false; return; }
    if (EX.status === 'error') { emptyEl.innerHTML = '<p>Не удалось загрузить расписание экзаменов. Проверьте подключение к интернету.</p><button type="button" class="btn" data-retry>Повторить</button>'; emptyEl.hidden = false; return; }
    if (EX.status === 'none') {
      emptyEl.innerHTML = '<h2>График экзаменов пока не опубликован</h2>' +
        '<p>Когда в Google-таблице появится лист «' + esc(CFG.examSheet || 'Экзамены') + '», экзамены, зачёты и пересдачи будут показаны здесь автоматически.</p>' +
        '<ol><li>Добавьте в таблицу лист с названием <code>' + esc(CFG.examSheet || 'Экзамены') + '</code>.</li>' +
        '<li>В первой строке подпишите колонки: №, Группа, Дисциплина, Дата и время, Аудитория, Преподаватель, Тип аттестации, Дата и время пересдачи.</li>' +
        '<li>Дата — «23.06.2027 10:00» или «23 июня 2027 г 10 часов 00 минут». Тип — Экзамен, Зачёт или Диф. зачёт.</li></ol>' +
        '<button type="button" class="btn" data-template>Скачать шаблон листа</button>';
      emptyEl.hidden = false;
      return;
    }
    var all = examItems().filter(function (it) { return matchesExam(it, today); });
    // the general schedule lists exams; credits, diff. credits and retakes open only by their own buttons
    var main = ex.type === 'retake' ? [] : sortItems(all.filter(function (it) { return !it.retake; }));
    var ret = ex.type === 'retake' ? sortItems(all.filter(function (it) { return it.retake; })) : [];
    if (!main.length && !ret.length) {
      emptyEl.innerHTML = '<p>' + (EX.records.length ? 'По этим фильтрам ничего не нашлось.' : 'В базе пока нет экзаменов и зачётов. Администратор может добавить их кнопкой выше.') + '</p>' +
        (EX.records.length ? '<button type="button" class="btn" data-reset>Сбросить фильтры</button>' : '');
      emptyEl.hidden = false;
      return;
    }
    if (main.length || ex.type !== 'retake') {
      $('e-main').hidden = !main.length;
      $('e-main-title').textContent = ex.type === 'exam' ? 'Экзамены' : ex.type === 'credit' ? 'Зачёты' : ex.type === 'diff' ? 'Дифференцированные зачёты' : 'Экзамены и зачёты';
      $('e-count').textContent = count(main.length, 'запись', 'записи', 'записей') + (EX.skipped ? ' · строк не прочитано: ' + EX.skipped : '');
      if (main.length) renderTable($('e-table'), $('e-pager'), main, 'page', today);
    }
    $('e-retakes').hidden = !ret.length;
    $('e-rcount').textContent = count(ret.length, 'запись', 'записи', 'записей');
    if (ret.length) renderTable($('e-rtable'), $('e-rpager'), ret, 'rpage', today);
  }

  // ---------- the record book and a teacher's exams (ТЗ, этап 8) ----------
  // An exam of the schedule matched to a plan subject of the group, in the given half of the year.
  // 1 August that opens the current academic year: older sessions stay out of this year's record book
  function acadStart() {
    if (GRAFIK.data && GRAFIK.data.weeks.length) return GRAFIK.data.weeks[0].slice(0, 4) + '-08-01';
    var t = new Date(); return ymd(t.getMonth() >= 7 ? t.getFullYear() : t.getFullYear() - 1, 8, 1);
  }
  function findExam(group, name, half) {
    var h2 = half2Of(group), from = acadStart(), best = null, bestScore = 59;
    EX.records.forEach(function (r) {
      if (r.date < from) return;
      if (r.group !== group && resolveKey(r.group, [group], null) !== group) return;
      if (half && (half === 1 ? r.date >= h2 : r.date < h2)) return;
      var sc = Hours.matchScore(name, r.discipline);
      if (sc > bestScore) { best = r; bestScore = sc; }
    });
    return best;
  }
  function examManageHTML(r, seed) {
    if (!isAdmin()) return '';
    return r ? '<button type="button" class="btn btn-small" data-exam-edit="' + EX.records.indexOf(r) + '">Изменить</button> <button type="button" class="btn btn-quiet btn-small" data-retake-edit="' + EX.records.indexOf(r) + '">Пересдача</button>' : '<button type="button" class="btn btn-small" data-exam-seed="' + esc(JSON.stringify(seed)) + '">Назначить дату</button>';
  }
  function examWhen(r) {
    if (!r) return '<span class="hint">не назначена</span>';
    return esc(fmtShort(r.date)) + (r.time ? '<br>' + esc(r.time) : '') + (r.room ? '<br><span class="hint">' + esc(roomLabel(r.room)) + '</span>' : '');
  }
  function groupSelectHTML(id, value) {
    return '<select id="' + id + '" name="' + id + '"><option value="">Выберите группу</option>' + groupsByCourse().map(function (c) {
      return '<optgroup label="' + esc(c.label) + '">' + c.groups.map(function (g) { return '<option value="' + esc(g) + '"' + (g === value ? ' selected' : '') + '>' + esc(g) + '</option>'; }).join('') + '</optgroup>';
    }).join('') + '</select>';
  }
  function renderBook() {
    var box = $('ex-book'), g = state.bookGroup || state.group;
    var half = state.bookHalf || (g ? currentHalf(g) : 1), plan = g ? planFor(g) : null;
    var controls = '<section class="filters"><div class="filter-row"><div class="field f-entity"><label for="book-group">Группа</label>' + groupSelectHTML('book-group', g) + '</div>' +
      '<div class="field"><span class="label" id="book-half-label">Семестр</span><div class="seg" id="book-half" role="radiogroup" aria-labelledby="book-half-label">' +
      [1, 2].map(function (h) { return '<button type="button" role="radio" data-value="' + h + '" aria-checked="' + (h === half ? 'true' : 'false') + '">' + h + ' семестр</button>'; }).join('') + '</div></div></div></section>';
    if (!g) { box.innerHTML = controls + '<p class="empty">Выберите группу, чтобы открыть разворот зачётной книжки.</p>'; return; }
    if (PLAN.status === 'loading') { box.innerHTML = controls + '<p class="empty">Загружаю учебный план…</p>'; return; }
    if (!plan) { box.innerHTML = controls + '<p class="empty">Для группы ' + esc(g) + ' на листе «' + esc(CFG.planSheet) + '» нет учебного плана, поэтому список экзаменов и зачётов взять неоткуда.</p>'; return; }
    var exams = [], diffs = [], credits = [];
    plan.subjects.forEach(function (sj) {
      var f = sj.form[half - 1];
      if (!f) return;
      var item = { s: sj, f: f, hours: half === 1 ? sj.s1 : sj.s2, rec: findExam(g, sj.name, half) };
      if (f === 'Э') exams.push(item); else if (f === 'ДЗ' || f === 'ЗаО') diffs.push(item); else if (f === 'З') credits.push(item);
    });
    // no grades in the record book: a «З» or «ДЗ» next to a subject reads like a mark, so the kind of credit is a subheading
    function page(title, parts) {
      parts = parts.filter(function (p) { return p.list.length; });
      return '<section class="book-page"><h3>' + esc(title) + '</h3>' + (parts.length ? parts.map(function (p) { return (p.sub ? '<h4 class="book-sub">' + esc(p.sub) + '</h4>' : '') + bookTable(p.list); }).join('') : '<p class="hint">В этом семестре по плану нет.</p>') + '</section>';
    }
    function bookTable(list) {
      return '<table class="book-table"><thead><tr><th scope="col">№</th><th scope="col">Дисциплина</th><th scope="col" class="n">Часов</th>' +
        '<th scope="col">Дата</th><th scope="col">Преподаватель</th></tr></thead><tbody>' + list.map(function (it, i) {
          return '<tr><td class="n">' + (i + 1) + '</td><td>' + discHTML(it.s.name) + '</td>' +
            '<td class="n">' + (it.hours || '—') + '</td><td>' + examWhen(it.rec) + examManageHTML(it.rec, {group:g,discipline:it.s.name,teacher:it.s.teacher,type:it.f === 'Э' ? 'Экзамен' : it.f === 'З' ? 'Зачёт' : 'Диф. зачёт'}) + '</td><td>' + esc(it.s.teacher || '—') + '</td></tr>';
        }).join('') + '</tbody></table>';
    }
    box.innerHTML = controls + '<div class="book" aria-label="Разворот зачётной книжки">' +
      '<div class="book-cap"><strong>Зачётная книжка</strong><span>Группа <span translate="no">' + esc(g) + '</span>' + (plan.title ? ', ' + esc(plan.title) : '') + ', ' + half + ' семестр ' + esc(GRAFIK.data ? GRAFIK.data.title.replace(/^.*?(\d{4}\/\d{4}).*$/, '$1') : '') + '</span></div>' +
      page('Экзамены', [{ list: exams }]) + page('Зачёты и дифференцированные зачёты', [{ sub: 'Дифференцированные зачёты', list: diffs }, { sub: 'Зачёты', list: credits }]) + '</div>' +
      '<p class="hint">Список и часы — по учебному плану группы (лист «' + esc(CFG.planSheet) + '»), даты — по расписанию экзаменов. Контрольные работы в зачётку не вносятся.</p>';
  }
  function renderTeacherExams() {
    var box = $('ex-teacher'), t = state.exTeacher || state.teacher;
    var plannedTeachers = [];
    DB.groups.forEach(function (g) { var p = planFor(g); if (p) p.subjects.forEach(function (s) { plannedTeachers = plannedTeachers.concat(splitTeachers(s.teacher)); }); });
    var names = uniq(DB.teachers.concat(plannedTeachers, EX.records.map(function (r) { return splitTeachers(r.teacher)[0] || ''; }))).sort(cmpRu);
    var controls = '<section class="filters"><div class="filter-row"><div class="field f-entity"><label for="ext-teacher">Преподаватель</label><select id="ext-teacher" name="ext-teacher">' +
      '<option value="">Выберите преподавателя</option>' + names.map(function (n) { return '<option value="' + esc(n) + '"' + (n === t ? ' selected' : '') + '>' + esc(n) + '</option>'; }).join('') + '</select></div></div></section>';
    if (!t) { box.innerHTML = controls + '<p class="empty">Выберите преподавателя, чтобы увидеть его экзамены и зачёты.</p>'; return; }
    var surname = t.split(/\s+/)[0].toLowerCase().replace(/ё/g, 'е'), items = [], used = [];
    function hasT(str) { return String(str || '').toLowerCase().replace(/ё/g, 'е').indexOf(surname) >= 0; }
    (DB.groups.length ? DB.groups : PLAN.order).forEach(function (g) {
      var p = planFor(g);
      if (!p) return;
      p.subjects.forEach(function (sj) {
        if (!hasT(sj.teacher)) return;
        [1, 2].forEach(function (h) {
          var f = sj.form[h - 1];
          if (!f || f === 'К/Р') return;
          var rec = findExam(g, sj.name, h);
          if (rec) used.push(rec);
          items.push({ group: g, name: sj.name, form: f, half: h, rec: rec });
        });
      });
    });
    EX.records.forEach(function (r) {
      if (hasT(r.teacher) && used.indexOf(r) < 0) items.push({ group: r.group, name: r.discipline, form: TYPE_KEY[r.type] === 'exam' ? 'Э' : TYPE_KEY[r.type] === 'diff' ? 'ДЗ' : 'З', half: r.date < acadStart() ? +halfOf(r.date) : r.date < half2Of(r.group) ? 1 : 2, rec: r });
    });
    items.sort(function (a, b) { return a.half - b.half || ((a.rec ? a.rec.date : '9') < (b.rec ? b.rec.date : '9') ? -1 : (a.rec ? a.rec.date : '9') > (b.rec ? b.rec.date : '9') ? 1 : cmpRu(a.group, b.group)); });
    if (!items.length) { box.innerHTML = controls + '<p class="empty">У преподавателя ' + esc(t) + ' по учебным планам и расписанию экзаменов нет ни экзаменов, ни зачётов.</p>'; return; }
    box.innerHTML = controls + '<section class="tablebox"><div class="table-head"><h2>Экзамены и зачёты: ' + esc(t) + '</h2><span class="count">' + count(items.length, 'запись', 'записи', 'записей') + '</span></div>' +
      '<div class="table-wrap"><table class="smart"><thead><tr><th scope="col">Семестр</th><th scope="col">Группа</th><th scope="col">Дисциплина</th><th scope="col">Форма</th><th scope="col">Дата</th></tr></thead><tbody>' +
      items.map(function (it) {
        return '<tr><td data-label="Семестр">' + it.half + '</td><td data-label="Группа"><span class="grp" translate="no">' + esc(it.group) + '</span></td><td class="title" data-label="Дисциплина">' + discHTML(it.name) + '</td>' +
          '<td data-label="Форма">' + formHTML(it.form) + '</td><td data-label="Дата">' + examWhen(it.rec) + examManageHTML(it.rec, {group:it.group,discipline:it.name,teacher:t,type:it.form === 'Э' ? 'Экзамен' : it.form === 'З' ? 'Зачёт' : 'Диф. зачёт'}) + '</td></tr>';
      }).join('') + '</tbody></table></div></section>' +
      '<p class="hint">Из учебных планов групп (где преподаватель указан у предмета) и расписания экзаменов.</p>';
  }

  // ---------- exams loaded from a document (ТЗ, этап 8: «экзамены в документе, импорт») ----------
  var MAMMOTH_URL = 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js';
  var mammothPromise = null, xi = { records: [], skipped: 0, bad: [], source: '', review: null };
  function loadMammoth() {
    if (window.mammoth) return Promise.resolve(window.mammoth);
    if (!mammothPromise) {
      mammothPromise = new Promise(function (resolve, reject) {
        var sc = document.createElement('script');
        sc.src = MAMMOTH_URL;
        sc.onload = function () { window.mammoth ? resolve(window.mammoth) : reject(new Error('mammoth')); };
        sc.onerror = function () { mammothPromise = null; reject(new Error('Не удалось загрузить модуль для .docx')); };
        document.head.appendChild(sc);
      });
    }
    return mammothPromise;
  }
  // the last group code in the text, the one nearest to the table. «группы № ПСОп-23», «ОДЛу -24»;
  // a Cyrillic З or О in the year stands for 3 or 0 («ПК-2З»)
  var GROUP_IN_TEXT = /([А-ЯЁ][А-ЯЁа-яё]*\d?)\s*-\s*([\dЗО]{2})([а-яё]?)(?![\dА-ЯЁа-яё])/g;
  function groupFromText(text) {
    var m, last = '';
    GROUP_IN_TEXT.lastIndex = 0;
    while ((m = GROUP_IN_TEXT.exec(text || ''))) last = m[1] + '-' + m[2].replace(/З/g, '3').replace(/О/g, '0') + m[3];
    return last;
  }
  function nodeText(td) {
    var ps = td.querySelectorAll('p');
    return ps.length ? Array.prototype.map.call(ps, function (q) { return q.textContent; }).join('\n') : td.textContent;
  }
  // every table of an HTML document in reading order, with the text between it and the table before
  // (it names the group) and the kind of the document so far («экзаменационной сессии», «зачетной недели»)
  function tablesFromHtml(html) {
    var doc = new DOMParser().parseFromString(html, 'text/html'), out = [], ctx = '', kind = '';
    (function walk(el) {
      Array.prototype.forEach.call(el.children, function (ch) {
        if (ch.tagName === 'TABLE') {
          var rows = Array.prototype.filter.call(ch.querySelectorAll('tr'), function (tr) { return tr.closest('table') === ch; });
          out.push({ rows: rows.map(function (tr) { return Array.prototype.map.call(tr.children, nodeText); }), context: ctx, kind: kind });
          ctx = '';
          return;
        }
        if (ch.querySelector('table')) { walk(ch); return; }
        var t = ch.textContent || '', k = examKindOf(t);
        if (k) kind = k;
        ctx += ' ' + t;
      });
    })(doc.body);
    return out;
  }
  function examsFromTables(tables) {
    var records = [], skipped = 0, bad = [], group = '', map = null, cols = 0, header = false;
    tables.forEach(function (t) {
      // in a document a table without its own group line goes on with the group above it (a table
      // broken over a page); every sheet of a workbook stands on its own
      var own = groupFromText(t.context);
      group = t.kind == null ? own : own || group;
      var width = Math.max.apply(null, [0].concat(t.rows.map(function (r) { return r.length; })));
      var res = parseExamSheet(t.rows, { group: group, kind: t.kind || '', map: t.kind != null && width === cols ? map : null });
      if (res) { records = records.concat(res.records); skipped += res.skipped; bad = bad.concat(res.bad); map = res.map; cols = width; header = header || res.header; }
    });
    return { records: records, skipped: skipped, bad: bad, header: header };
  }
  function readExamFile(file) { return readDocTables(file).then(examsFromTables); }

  // ---------- the uploaded table against everything the site knows ----------
  // Every row is checked against the group list, the teachers of the timetable, the study plan
  // (is it an exam or a credit this term), the PDF schedule (attestation days), holidays, the other
  // rows of the file and the lessons that stand at that time. Then the list of exams is changed to
  // match the table for the groups it covers, and the group's pairs in the exam's hours are cancelled.
  function bare(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/[^a-zа-я0-9]/g, ''); }
  function canonGroup(g) {
    var fixed = cellText(g).replace(/(-\s*)([\dЗО]{2})(?![\dА-ЯЁа-яё])/, function (m, a, y) { return a + y.replace(/З/g, '3').replace(/О/g, '0'); });
    var flat = bare(fixed);
    var hit = DB.groups.filter(function (x) { return bare(x) === flat; })[0];
    return hit || (DB.groups.length ? resolveKey(fixed.replace(/\s+/g, ''), DB.groups, courseOf) : '') || fixed.replace(/\s*-\s*/, '-');
  }
  // «Иванова А. С.» → «Иванова А.С.», and the spelling used in the timetable when there is one
  function canonTeacher(t) {
    return splitTeachers(String(t || '').replace(/\n/g, ',')).map(function (x) {
      var n = cellText(x).replace(/\.\s+(?=[А-ЯЁ]\.)/g, '.');
      var key = bare(n), hit = DB.teachers.filter(function (y) { return bare(y) === key; })[0];
      return hit || n;
    }).filter(Boolean).join(', ');
  }
  function sameDisc(a, b) { return bare(a) === bare(b) || Hours.matchScore(a, b) >= 90; }
  function shiftIso(iso, days) { return isoOfUtc(utcOfIso(iso) + days * DAY); }
  var PLAN_FORM_TYPE = { 'Э': 'Экзамен', 'ДЗ': 'Диф. зачёт', 'ЗаО': 'Диф. зачёт', 'З': 'Зачёт' };
  // the hours an exam takes: [start, end) in minutes, null for the whole day
  function examSpan(r) {
    if (!r.time) return null;
    var st = minutesOf(r.time);
    return [st, st + (r.type === 'Экзамен' ? (CFG.examMinutes || 240) : 90)];
  }
  function dayLetter(group, iso) {
    var gf = grafikFor(group);
    if (!gf) return '';
    var i = GRAFIK.data.weeks.indexOf(isoOfUtc(mondayOfUtc(utcOfIso(iso)))), wd = isoToDate(iso).getDay();
    return i < 0 || wd === 0 ? '' : gf.g.days[i].charAt(Math.min(wd, 5) - 1);
  }
  function reviewExams(input) {
    var today = todayISO(), records = input.map(function (r) {
      return Object.assign({}, r, { group: canonGroup(r.group), teacher: canonTeacher(r.teacher), rawGroup: r.group });
    });
    var win = {};
    records.forEach(function (r) {
      var w = win[r.group] || (win[r.group] = { min: r.date, max: r.date });
      if (r.date < w.min) w.min = r.date;
      if (r.date > w.max) w.max = r.date;
    });
    // what the site shows now for these groups in this session (±30 days around the dates of the file)
    function inWin(r) { var w = win[r.group]; return !!w && r.date >= shiftIso(w.min, -30) && r.date <= shiftIso(w.max, 30); }
    var current = EX.records.filter(inWin), used = [];
    var items = records.map(function (r, idx) {
      var notes = [], cancels = [];
      function note(level, text) { notes.push({ level: level, text: text }); }
      // the record it replaces
      var old = null;
      current.forEach(function (c, ci) { if (!old && used.indexOf(ci) < 0 && c.group === r.group && sameDisc(c.discipline, r.discipline) && c.type === r.type) { old = c; used.push(ci); } });
      if (!old) current.forEach(function (c, ci) { if (!old && used.indexOf(ci) < 0 && c.group === r.group && sameDisc(c.discipline, r.discipline)) { old = c; used.push(ci); } });
      var changes = [];
      if (old) {
        if (old.date !== r.date || (old.time || '') !== (r.time || '')) changes.push('было ' + fmtShort(old.date) + (old.time ? ' ' + old.time : ''));
        if ((old.room || '') !== (r.room || '')) changes.push('аудитория была ' + (old.room || '—'));
        if (bare(old.teacher) !== bare(r.teacher)) changes.push('преподаватель был ' + (old.teacher || '—'));
        if (old.type !== r.type) changes.push('было: ' + old.type.toLowerCase());
        if ((old.retakeDate || '') !== (r.retakeDate || '') || (old.retakeTime || '') !== (r.retakeTime || '')) changes.push('пересдача была ' + (old.retakeDate ? fmtShort(old.retakeDate) + (old.retakeTime ? ' ' + old.retakeTime : '') : '—'));
      }
      if (r.groupGuess) note('info', 'Группа в таблице не указана — взята выбранная на странице: ' + r.group + '.');
      if (r.typeGuess) note('info', 'Тип аттестации не указан — записано как экзамен.');
      // the group, the teacher
      if (DB.groups.length && DB.groups.indexOf(r.group) < 0) note('warn', 'Группы «' + r.rawGroup + '» нет в списке групп.');
      else if (bare(r.rawGroup) !== bare(r.group)) note('info', 'Группа «' + r.rawGroup + '» записана как ' + r.group + '.');
      splitTeachers(r.teacher).forEach(function (t) { if (DB.teachers.length && DB.teachers.indexOf(t) < 0 && !/^вак/i.test(t)) note('info', 'Преподавателя ' + t + ' нет в расписании занятий: проверьте, как он записан.'); });
      // the study plan: is it on the plan, and is it an exam or a credit this term
      var plan = planFor(r.group);
      if (plan) {
        var best = null, score = 59;
        plan.subjects.forEach(function (sb) { var sc = Hours.matchScore(sb.name, r.discipline); if (sc > score) { best = sb; score = sc; } });
        if (!best) note('warn', 'Такого предмета нет в учебном плане группы.');
        else {
          var half = r.date < half2Of(r.group) ? 1 : 2, form = best.form[half - 1], want = PLAN_FORM_TYPE[form];
          if (want && want !== r.type) note('warn', 'По учебному плану в ' + half + '-м семестре — ' + want.toLowerCase() + ', а в таблице — ' + r.type.toLowerCase() + '.');
          else if (!form) {
            var other = best.form[2 - half];
            note('info', other ? 'По учебному плану аттестация по предмету — в ' + (3 - half) + '-м семестре (' + (PLAN_FORM_TYPE[other] || other).toLowerCase() + ').' : 'В учебном плане у предмета в ' + half + '-м семестре нет формы аттестации.');
          }
        }
      }
      // the date
      var wd = isoToDate(r.date).getDay(), letter = dayLetter(r.group, r.date);
      if (r.date < today) note('info', 'Дата уже прошла.');
      if (wd === 0) note('warn', 'Это воскресенье.');
      if (isHoliday(r.date)) note('warn', 'Это праздничный день.');
      if (letter && letter !== 'A' && letter !== 'G') note('info', 'По графику учебного процесса в этот день: ' + (WEEK_TYPES[letter] || letter) + '.');
      if (r.retakeDate && isHoliday(r.retakeDate)) note('warn', 'Пересдача выпадает на праздник.');
      // the other rows of the file
      var span = examSpan(r);
      records.forEach(function (o, j) {
        if (j === idx || o.date !== r.date) return;
        if (o.group === r.group && j < idx && !sameDisc(o.discipline, r.discipline)) note('warn', 'В этот же день у группы ещё: ' + splitCode(o.discipline).name + '.');
        var os = examSpan(o), overlap = span && os && os[0] < span[1] && os[1] > span[0];
        if (!overlap || o.group === r.group || sameDisc(o.discipline, r.discipline)) return;
        var shared = splitTeachers(o.teacher).filter(function (t) { return splitTeachers(r.teacher).indexOf(t) >= 0; });
        if (shared.length) note('warn', shared.join(', ') + ' в это время принимает ещё ' + splitCode(o.discipline).name + ' у группы ' + o.group + '.');
        if (r.room && o.room && bare(r.room) === bare(o.room)) note('warn', 'Аудитория ' + r.room + ' в это время занята: ' + o.group + ', ' + splitCode(o.discipline).name + '.');
      });
      // lessons at that time: the group's pairs give way to an exam; the teacher's pairs elsewhere are a clash
      var sheet = DB.sheets[r.group];
      if (sheet) {
        (sheet.byDay[r.date] || []).forEach(function (row) {
          if (row.ghost || !row.disc) return;
          var s0 = minutesOf(row.start), s1 = minutesOf(row.end);
          if (span ? !(s0 < span[1] && s1 > span[0]) : r.type !== 'Экзамен') return;
          cancels.push({ row: row, on: r.type === 'Экзамен' && !(row.edit && row.edit.kind === 'add') });
        });
      } else if (DB.groups.indexOf(r.group) >= 0) note('info', 'Расписание группы ещё не загружено, пары в этот день не проверены.');
      if (span) splitTeachers(r.teacher).forEach(function (t) {
        Object.keys(DB.sheets).forEach(function (g) {
          if (g === r.group) return;
          (DB.sheets[g].byDay[r.date] || []).forEach(function (row) {
            if (!row.ghost && splitTeachers(row.teacher).indexOf(t) >= 0 && minutesOf(row.start) < span[1] && minutesOf(row.end) > span[0]) note('warn', t + ' в это время ведёт пару в группе ' + g + ' (' + slotLabel(row.start) + ').');
          });
        });
      });
      return { rec: r, status: !old ? 'new' : changes.length ? 'changed' : 'same', changes: changes, old: old, notes: notes, cancels: cancels };
    });
    var removed = current.filter(function (c, ci) { return used.indexOf(ci) < 0; });
    return { items: items, removed: removed, groups: Object.keys(win).sort(cmpRu), win: win };
  }
  var xiGroup = '';
  function openExamImport() {
    xiGroup = ex.group.length === 1 ? ex.group[0] : state.view === 'exams' && state.exPage === 'book' && state.bookGroup ? state.bookGroup : '';
    xi = { records: [], skipped: 0, bad: [], source: '', review: null };
    $('xi-file').value = '';
    $('xi-url').value = '';
    $('xi-status').textContent = '';
    $('xi-preview').innerHTML = '';
    $('xi-save').disabled = true;
    $('xi-save').textContent = 'Применить';
    $('xi-dialog').showModal();
    // the teachers' pairs in every group are needed to see clashes
    ensureAll();
  }
  var XI_STATUS = { 'new': ['Новая запись', 'b-add'], changed: ['Изменится', 'b-replace'], same: ['Без изменений', 's-past'] };
  function showExamPreview(res, source) {
    xi = { records: res.records, skipped: res.skipped, bad: res.bad || [], source: source, review: null };
    if (!res.records.length && res.header && !xi.bad.length) {
      $('xi-save').disabled = true;
      $('xi-status').textContent = 'Таблица прочитана, но в ней нет заполненных строк. Заполните хотя бы одну строку: дисциплина и дата (группа, преподаватель, аудитория и тип — по желанию), сохраните файл и загрузите его снова. Пример строки — на листе «Как заполнять» в Excel-шаблоне, готовые примеры — по ссылкам «Примеры» выше.';
      $('xi-preview').innerHTML = '';
      return;
    }
    if (!res.records.length) {
      $('xi-save').disabled = true;
      $('xi-status').textContent = 'В документе не нашлось таблицы с экзаменами. Нужны колонки «Дисциплина» и «Дата»; группа — колонкой или строкой над таблицей («группы № ИС2-25»); экзамен это или зачёт — колонкой «Тип аттестации» или заголовком документа («…экзаменационной сессии», «…зачётной недели»).' +
        (xi.bad.length ? ' Не прочитано строк: ' + xi.bad.length + '.' : '');
      $('xi-preview').innerHTML = xi.bad.length ? badRowsHTML(xi.bad) : '';
      return;
    }
    renderExamReview();
  }
  function badRowsHTML(bad) {
    return '<details class="xi-bad"><summary>Строки, которые не прочитались: ' + bad.length + '</summary><ul>' +
      bad.slice(0, 60).map(function (b) { return '<li>' + esc(b.text || '(пустая строка)') + ' — ' + esc(b.why) + '</li>'; }).join('') + '</ul></details>';
  }
  function renderExamReview() {
    var keep = {};
    if (xi.review) xi.review.items.forEach(function (it, i) { it.cancels.forEach(function (c, j) { keep[i + ':' + j] = c.on; }); });
    var rv = reviewExams(xi.records);
    rv.items.forEach(function (it, i) { it.cancels.forEach(function (c, j) { if (keep[i + ':' + j] != null) c.on = keep[i + ':' + j]; }); });
    xi.review = rv;
    var n = { 'new': 0, changed: 0, same: 0 }, warn = 0, pairs = 0;
    rv.items.forEach(function (it) { n[it.status]++; if (it.notes.some(function (x) { return x.level === 'warn'; })) warn++; it.cancels.forEach(function (c) { if (c.on) pairs++; }); });
    var loading = net.bulk ? ' Загружаю расписание групп для проверки: ' + net.bulk.done + ' из ' + net.bulk.total + '…' : '';
    $('xi-status').textContent = 'Прочитано ' + count(rv.items.length, 'запись', 'записи', 'записей') + ' для ' + count(rv.groups.length, 'группы', 'групп', 'групп') + '. ' +
      'Новых: ' + n['new'] + ', изменится: ' + n.changed + ', без изменений: ' + n.same + (rv.removed.length ? ', уберётся: ' + rv.removed.length : '') + '. ' +
      (warn ? 'С замечаниями: ' + warn + '. ' : 'Замечаний нет. ') + (xi.skipped ? 'Не прочитано строк: ' + xi.skipped + '.' : '') + loading;
    var order = rv.items.map(function (it, i) { return i; }).sort(function (a, b) {
      var A = rv.items[a], B = rv.items[b], w = function (it) { return it.notes.some(function (x) { return x.level === 'warn'; }) ? 0 : it.status === 'changed' ? 1 : it.status === 'new' ? 2 : 3; };
      return w(A) - w(B) || cmpRu(A.rec.group, B.rec.group) || (A.rec.date < B.rec.date ? -1 : A.rec.date > B.rec.date ? 1 : 0);
    });
    var body = order.map(function (i) {
      var it = rv.items[i], r = it.rec, st = XI_STATUS[it.status];
      var notes = it.notes.map(function (x) { return '<li class="xi-' + x.level + '">' + esc(x.text) + '</li>'; }).join('');
      var cancels = it.cancels.map(function (c, j) {
        var row = c.row;
        return '<li><label><input type="checkbox" data-xi-cancel="' + i + ':' + j + '"' + (c.on ? ' checked' : '') + '> Отменить ' + esc(slotLabel(row.start)) + ': ' + esc(splitCode(row.disc).name) + (row.teacher ? ', ' + esc(row.teacher) : '') + (row.sub ? ', ' + row.sub + ' подгруппа' : '') + '</label></li>';
      }).join('');
      return '<tr><td data-label="Группа"><span class="grp" translate="no">' + esc(r.group) + '</span></td>' +
        '<td class="title" data-label="Дисциплина">' + discHTML(r.discipline) + '<span class="aka">' + esc(r.type) + (r.teacher ? ' · ' + esc(r.teacher) : '') + (r.room ? ' · ауд. ' + esc(r.room) : '') + '</span></td>' +
        '<td data-label="Когда">' + esc(fmtShort(r.date)) + (r.time ? ', ' + esc(r.time) : '') + (r.retakeDate ? '<span class="aka">пересдача ' + esc(fmtShort(r.retakeDate)) + (r.retakeTime ? ', ' + esc(r.retakeTime) : '') + '</span>' : '') + '</td>' +
        '<td data-label="Что будет"><span class="badge ' + st[1] + '">' + st[0] + '</span>' + (it.changes.length ? '<span class="aka">' + esc(it.changes.join('; ')) + '</span>' : '') + '</td>' +
        '<td data-label="Проверка">' + (notes || cancels ? '<ul class="xi-notes">' + notes + cancels + '</ul>' : '<span class="xi-ok">✓ всё сходится</span>') + '</td></tr>';
    }).join('');
    var removed = rv.removed.length ? '<details class="xi-bad" open><summary>Уберутся с сайта, потому что их нет в новой таблице: ' + rv.removed.length + '</summary><ul>' +
      rv.removed.map(function (c) { return '<li><span class="grp" translate="no">' + esc(c.group) + '</span> ' + esc(splitCode(c.discipline).name) + ', ' + esc(c.type.toLowerCase()) + ', ' + esc(fmtShort(c.date)) + (c.time ? ' ' + esc(c.time) : '') + '</li>'; }).join('') + '</ul></details>' : '';
    $('xi-preview').innerHTML = '<p class="hint">Таблица заменит экзамены и зачёты только у этих групп: ' + esc(rv.groups.join(', ')) + '. У остальных групп ничего не изменится. Отмеченные пары в часы экзамена будут отменены правкой администратора.</p>' +
      '<div class="table-wrap"><table class="smart xi-table"><thead><tr><th scope="col">Группа</th><th scope="col">Дисциплина</th><th scope="col">Когда</th><th scope="col">Что будет</th><th scope="col">Проверка</th></tr></thead><tbody>' + body + '</tbody></table></div>' +
      removed + (xi.bad.length ? badRowsHTML(xi.bad) : '');
    $('xi-save').disabled = false;
    $('xi-save').textContent = 'Применить: ' + count(rv.items.length, 'запись', 'записи', 'записей') + (pairs ? ', отменить ' + count(pairs, 'пару', 'пары', 'пар') : '');
  }
  function saveExamImport() {
    var rv = xi.review;
    if (!rv) return;
    var btn = $('xi-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    // exams uploaded earlier stay for the other groups and for other sessions of these groups
    var keep = EXI.raw.filter(function (r) {
      var w = rv.win[r.group];
      return !(w && r.date >= shiftIso(w.min, -30) && r.date <= shiftIso(w.max, 30));
    }).map(function (r) {
      return { group: r.group, discipline: r.discipline, date: r.date, time: r.time, room: r.room, teacher: r.teacher, type: r.type, retakeDate: r.retakeDate, retakeTime: r.retakeTime };
    });
    var fresh = rv.items.map(function (it) {
      var r = it.rec;
      return { group: r.group, discipline: r.discipline, date: r.date, time: r.time, room: r.room, teacher: r.teacher, type: r.type, retakeDate: r.retakeDate, retakeTime: r.retakeTime };
    });
    var batch = 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), cancels = [], seen = {};
    rv.items.forEach(function (it) {
      it.cancels.forEach(function (c) {
        var row = c.row, k = row.group + '|' + row.day + '|' + row.start + '|' + (row.sub || 0);
        if (!c.on || seen[k] || (row.edit && row.edit.kind === 'add')) return;
        seen[k] = true;
        var was = JSON.stringify(slotRows(row.group, row.day, row.start).filter(function (x) { return !x.ghost && !x.edit; }).map(function (x) { return { disc: x.disc, teacher: x.teacher, room: x.room, sub: x.sub }; }));
        cancels.push({ id: row.edit ? row.edit.id : '', kind: 'cancel', group: row.group, date: row.day, start: row.start, end: row.end, sub: +row.sub || 0,
          note: (it.rec.type === 'Экзамен' ? 'экзамен: ' : it.rec.type.toLowerCase() + ': ') + splitCode(it.rec.discipline).name, was: row.edit ? row.edit.was : was });
      });
    });
    var note = 'экзамены из «' + (xi.source || 'документа') + '»';
    Api.examsImport(keep.concat(fresh), xi.source).then(function () {
      var chunks = [];
      for (var i = 0; i < cancels.length; i += 100) chunks.push(cancels.slice(i, i + 100));
      return chunks.reduce(function (pr, ch) { return pr.then(function () { return Api.save(ch, { batch: batch, comment: note }).then(mergeEdits); }); }, Promise.resolve());
    }).then(function () {
      HIST.status = 'idle';
      $('xi-dialog').close();
      return loadEdits().then(function () {
        toast('Экзамены обновлены: ' + count(fresh.length, 'запись', 'записи', 'записей') + ' для ' + count(rv.groups.length, 'группы', 'групп', 'групп') +
          (cancels.length ? ', отменено ' + count(cancels.length, 'пара', 'пары', 'пар') : '') + '. Вернуть прошлый список можно в журнале правок («Откатить»), отмены пар — в разделе «Изменения».');
      });
    }, function (err) {
      $('xi-status').textContent = err.message + (EDITS.list.some(function (x) { return x.batch === batch; }) ? ' Экзамены сохранены, часть отмен пар тоже — их видно в разделе «Изменения».' : '');
      btn.disabled = false;
      btn.textContent = 'Применить';
    });
  }

  // ---------- practice from a table or a document ----------
  // Two kinds of files are understood:
  // 1) periods: a table with the group and the dates of УП / ПП / ПДП / exams / holidays — the college's
  //    «Учебный график 26-27» (rows «Дата начала» / «Дата окончания» under each group), or a plain list
  //    «Группа · Вид · Начало · Окончание», or «с 24 ноября 2026 по 7 декабря 2026» in a cell.
  //    The weeks of the PDF schedule are changed for those days, and the group's ordinary pairs on them go.
  // 2) a practice timetable by pairs («Расписание учебной практики»: дата, пара, наименование, руководитель):
  //    on each of its days the group's pairs are replaced by the pairs of the file.
  var pi = { kind: '', items: [], source: '', notes: [] };
  var PI_NAMES = { U: 'учебная практика', P: 'производственная практика', X: 'преддипломная практика', A: 'экзамены', K: 'каникулы' };
  function piKindOfHeader(h) {
    h = cellText(h).toLowerCase().replace(/ё/g, 'е');
    if (!h) return '';
    if (/преддипл|(^|[^а-я])пдп([^а-я]|$)/.test(h)) return 'X';
    if (/производств|(^|[^а-я])пп([^а-я.]|$)/.test(h)) return 'P';
    if (/учебн\S*\s+практ|(^|[^а-я])уп([^а-я.]|$)/.test(h)) return 'U';
    if (/каникул/.test(h)) return 'K';
    if (/экзамен|аттестац|сесси/.test(h)) return 'A';
    return '';
  }
  // every date in a text: «24.11.2026», «24 ноября 2026», «24 ноября 2026г.»
  function datesIn(text) {
    // a date cell of a workbook may come as its day number (46350 = 24.11.2026)
    if (/^\s*\d{5}(\.\d+)?\s*$/.test(String(text || ''))) { var sd = parseExamDate(text); return sd && !sd.error && sd.date > '2000' ? [sd.date] : []; }
    var out = [], re = /(\d{1,2})\.(\d{1,2})\.(\d{2,4})|(\d{1,2})\s+([А-ЯЁа-яё]+)\s+(\d{4})/g, m;
    while ((m = re.exec(cellText(text)))) {
      var d = parseExamDate(m[0]);
      if (d && !d.error) out.push(d.date);
    }
    return out;
  }
  // a range typed with the year of its start at the end («29.12.2026 – 11.01.2026») runs into the next year
  function piRange(from, to) {
    if (!from || !to) return null;
    if (to < from) { var t2 = (+to.slice(0, 4) + 1) + to.slice(4); if (t2 >= from && daysBetween(from, t2) < 200) to = t2; else return null; }
    return daysBetween(from, to) > 200 ? null : [from, to];
  }
  function piPeriods(tables) {
    var out = [];
    tables.forEach(function (t) {
      var rows = t.rows, hi = -1, gcol = -1, kinds = {}, list = null;
      for (var i = 0; i < Math.min(rows.length, 15) && hi < 0; i++) {
        var r = rows[i] || [], g = -1, ks = {}, lv = {};
        r.forEach(function (c, j) {
          var h = cellText(c).toLowerCase().replace(/ё/g, 'е');
          if (/^групп/.test(h) && g < 0) g = j;
          else if (/^(вид|тип)/.test(h)) lv.kind = j;
          else if (/начал|^с$|^дата с/.test(h)) lv.from = j;
          else if (/оконч|конец|^по$|^дата по/.test(h)) lv.to = j;
          else if (/период|сроки|даты/.test(h)) lv.span = j;
          else { var k = piKindOfHeader(h); if (k) ks[j] = k; }
        });
        if (g < 0) continue;
        if (lv.kind != null && (lv.from != null || lv.span != null)) { hi = i; gcol = g; list = lv; }
        else if (Object.keys(ks).length) { hi = i; gcol = g; kinds = ks; }
      }
      if (hi < 0) return;
      var pending = {};
      for (var r2 = hi + 1; r2 < rows.length; r2++) {
        var row = rows[r2] || [], gname = cellText(row[gcol]);
        if (!gname || !groupFromText(gname.replace(/\s+/g, '')) && !/^[А-ЯЁ][А-ЯЁа-яё]*\d?-?\d{2}/.test(gname)) continue;
        var group = canonGroup(gname);
        if (list) {
          var kind = piKindOfHeader(row[list.kind]) || piKindOfHeader('(' + cellText(row[list.kind]) + ')');
          var ds = list.span != null ? datesIn(row[list.span]) : datesIn(row[list.from]).concat(list.to != null ? datesIn(row[list.to]) : []);
          var rg = kind && piRange(ds[0], ds[1] || ds[0]);
          if (rg) out.push({ group: group, raw: gname, kind: kind, label: cellText(row[list.kind]), from: rg[0], to: rg[1] });
          continue;
        }
        var label = row.filter(function (c, j) { return j !== gcol && !kinds[j]; }).join(' ').toLowerCase();
        var isStart = /начал/.test(label), isEnd = /оконч|конц/.test(label);
        Object.keys(kinds).forEach(function (j) {
          var ds2 = datesIn(row[j]), key = group + '|' + j;
          if (isStart) { if (ds2[0]) pending[key] = ds2[0]; return; }
          var from = isEnd ? pending[key] : ds2[0], to = isEnd ? ds2[0] : ds2[1] || ds2[0];
          var rg2 = piRange(from, to);
          if (rg2) out.push({ group: group, raw: gname, kind: kinds[j], label: cellText(rows[hi][j]), from: rg2[0], to: rg2[1] });
        });
      }
    });
    // the same period written twice (in two semester blocks or on two sheets) is one
    var seen = {};
    return out.filter(function (p) { var k = [p.group, p.kind, p.from, p.to].join('|'); if (seen[k]) return false; seen[k] = true; return true; });
  }
  // «2 пара», «2», «II», «10:10-11:40», «10:10» → the start of the pair
  function piSlot(v) {
    var s = cellText(v), m = /(\d{1,2})[:.](\d{2})/.exec(s);
    var slots = (CFG.slots || []).map(function (x) { return normTime(x[0]); });
    if (m) { var t = normTime(m[1] + ':' + m[2]); return slots.indexOf(t) >= 0 ? t : slots.filter(function (x) { return Math.abs(minutesOf(x) - minutesOf(t)) <= 20; })[0] || t; }
    var roman = { 'I': 1, 'II': 2, 'III': 3, 'IV': 4, 'V': 5, 'VI': 6 };
    m = /^(\d)\b|^(I{1,3}|IV|V|VI)\b/.exec(s.toUpperCase());
    var n = m ? (m[1] ? +m[1] : roman[m[2]]) : 0;
    return n && slots[n - 1] ? slots[n - 1] : '';
  }
  function piPairs(tables, fileName) {
    var out = [];
    tables.forEach(function (t) {
      var rows = t.rows, map = null, hi = -1;
      for (var i = 0; i < Math.min(rows.length, 15) && !map; i++) {
        var mp = {};
        (rows[i] || []).forEach(function (c, j) {
          var h = cellText(c).toLowerCase().replace(/ё/g, 'е');
          if (!h) return;
          if (/^групп/.test(h)) { if (mp.group == null) mp.group = j; }
          else if (/дата|день/.test(h)) { if (mp.date == null) mp.date = j; }
          else if (/пара|время|урок/.test(h)) { if (mp.slot == null) mp.slot = j; }
          else if (/руковод|препод|мастер|ответств/.test(h)) { if (mp.teacher == null) mp.teacher = j; }
          else if (/ауд|кабин|место/.test(h)) { if (mp.room == null) mp.room = j; }
          else if (/наименован|тема|вид|содерж|дисципл|модул|работ/.test(h)) { if (mp.name == null) mp.name = j; }
        });
        if (mp.date != null && mp.slot != null && mp.name != null) { map = mp; hi = i; }
      }
      if (!map) return;
      var ctx = t.context + ' ' + fileName + ' ' + rows.slice(0, hi).map(function (r) { return (r || []).join(' '); }).join(' ');
      var dGroup = groupFromText(ctx) || groupFromText(ctx.replace(/([А-ЯЁ][А-ЯЁа-яё]*\d?)(\d{2})(?=[^\d]|$)/g, '$1-$2'));
      var low = ctx.toLowerCase(), kind = /производствен/.test(low) ? 'P' : /преддипл/.test(low) ? 'X' : 'U';
      var lastDate = '';
      for (var r = hi + 1; r < rows.length; r++) {
        var row = rows[r] || [], ds = datesIn(row[map.date]);
        if (ds[0]) lastDate = ds[0];
        var name = cellText(row[map.name]), start = piSlot(row[map.slot]);
        if (!name || !start || !lastDate) continue;
        var group = canonGroup(cellText(row[map.group]) || dGroup);
        if (!group) continue;
        out.push({ group: group, date: lastDate, start: start, name: name, teacher: canonTeacher(row[map.teacher]), room: cellText(row[map.room]), kind: kind });
      }
    });
    return out;
  }
  function tablesFromWorkbook(X, buf) {
    var wb = X.read(buf, { type: 'array' });
    return wb.SheetNames.map(function (sn) {
      var rows = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: '' }).map(function (r) { return r.map(function (v) { return v == null ? '' : String(v); }); });
      return { rows: rows, context: sn, kind: null };
    });
  }
  // a file's tables by its extension; a file without one is told by its bytes (a .docx and an .xlsx are both zips)
  function tablesFromBuffer(buf, name) {
    var ext = String(name || '').toLowerCase().split('.').pop();
    if (!/^(docx|xlsx|xls|csv|html?|ods)$/.test(ext)) {
      var head = new TextDecoder('latin1').decode(new Uint8Array(buf.slice(0, Math.min(buf.byteLength, 4000))));
      var tail = new TextDecoder('latin1').decode(new Uint8Array(buf.slice(Math.max(0, buf.byteLength - 20000))));
      ext = /^PK/.test(head) ? (/word\//.test(head + tail) ? 'docx' : 'xlsx') : /<html|<table/i.test(head) ? 'html' : /^\s*[^\x00]*[,;]/.test(head) && !/[\x00-\x08]/.test(head) ? 'csv' : 'xls';
    }
    if (ext === 'docx') return loadMammoth().then(function (m) { return m.convertToHtml({ arrayBuffer: buf }); }).then(function (r) { return tablesFromHtml(r.value); });
    if (ext === 'html' || ext === 'htm') return Promise.resolve(tablesFromHtml(new TextDecoder().decode(buf)));
    if (ext === 'csv') return Promise.resolve([{ rows: parseCSV(new TextDecoder().decode(buf).replace(/^﻿/, '')), context: name, kind: null }]);
    return loadXlsx().then(function (X) { return tablesFromWorkbook(X, buf); });
  }
  function readDocTables(file) {
    var name = file.name || '', ext = name.toLowerCase().split('.').pop();
    if (ext === 'doc' || ext === 'rtf' || ext === 'odt') {
      return Promise.reject(new Error('файлы .' + ext + ' браузер не читает. Откройте документ в Word и сохраните как .docx (Файл → Сохранить как → Документ Word), потом загрузите его'));
    }
    return file.arrayBuffer().then(function (buf) { return tablesFromBuffer(buf, name); });
  }
  // ---------- a Google link: Google Sheets, Google Docs or a file on Google Drive ----------
  // The browser asks Google itself first (no server needed, works in the demo too); if Google does not
  // give the file to the page, the admin server (apps-script/Code.gs) fetches it, when it is connected.
  function googleLink(u) {
    var m, s = String(u || '').trim();
    if ((m = /docs\.google\.com\/spreadsheets\/d\/([\w-]{20,})/.exec(s))) return { kind: 'sheet', id: m[1], gid: (/[#&?]gid=(\d+)/.exec(s) || [])[1] || '' };
    if ((m = /docs\.google\.com\/document\/d\/([\w-]{20,})/.exec(s))) return { kind: 'doc', id: m[1] };
    if ((m = /drive\.google\.com\/(?:file\/d\/|open\?(?:[^#]*&)?id=|uc\?(?:[^#]*&)?id=)([\w-]{20,})/.exec(s)) || (m = /docs\.google\.com\/(?:presentation|forms)\/d\/([\w-]{20,})/.exec(s))) return { kind: 'file', id: m[1] };
    return null;
  }
  function fetchBuffer(url) {
    return fetch(url, { cache: 'no-store', redirect: 'follow' }).then(function (r) {
      var type = r.headers.get('content-type') || '';
      if (!r.ok || /text\/html/.test(type) && !/export\?format=html/.test(url)) throw new Error('Google ответил ' + r.status);
      return r.arrayBuffer();
    });
  }
  function readGoogleLink(u) {
    var g = googleLink(u);
    if (!g) return Promise.reject(new Error('Это не ссылка на Google-таблицу, Google-документ или файл на Google Диске'));
    var ask;
    if (g.kind === 'sheet') {
      // the whole workbook as .xlsx, else the sheet the link points to through the same channel the timetable uses
      ask = fetchBuffer('https://docs.google.com/spreadsheets/d/' + g.id + '/export?format=xlsx').then(function (buf) {
        return tablesFromBuffer(buf, 'table.xlsx').then(function (t) { return { tables: t, name: 'Google-таблица' }; });
      }).catch(function () {
        return fetch('https://docs.google.com/spreadsheets/d/' + g.id + '/gviz/tq?tqx=out:csv&headers=0' + (g.gid ? '&gid=' + g.gid : ''), { cache: 'no-store' }).then(function (r) {
          if (!r.ok || (r.headers.get('content-type') || '').indexOf('csv') < 0) throw new Error('Google ответил ' + r.status);
          return r.text();
        }).then(function (csv) { return { tables: [{ rows: parseCSV(csv), context: '', kind: null }], name: 'Google-таблица', oneSheet: true }; });
      });
    } else if (g.kind === 'doc') {
      ask = fetchBuffer('https://docs.google.com/document/d/' + g.id + '/export?format=html').then(function (buf) {
        return { tables: tablesFromHtml(new TextDecoder().decode(buf)), name: 'Google-документ' };
      });
    } else {
      ask = fetchBuffer('https://drive.google.com/uc?export=download&id=' + g.id).then(function (buf) {
        return tablesFromBuffer(buf, '').then(function (t) { return { tables: t, name: 'файл с Google Диска' }; });
      });
    }
    return ask.catch(function () {
      if (Api.mode !== 'live') throw new Error('Google не отдал файл странице, а сервер правок не подключён (демо-режим). Скачайте файл на телефон или компьютер (в Google: Файл → Скачать → .xlsx или .docx) и загрузите его кнопкой «Перетащите файл сюда».');
      return Api.fetchLink(u).then(function (r) {
        if (r.html) return { tables: tablesFromHtml(r.html), name: r.name || 'Google-документ' };
        var bin = atob(r.data || ''), bytes = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
        return tablesFromBuffer(bytes.buffer, r.name || '').then(function (t) { return { tables: t, name: r.name || 'файл с Google Диска' }; });
      }, function (err) { throw new Error(err.message + ' Если не выходит, скачайте файл и загрузите его кнопкой «Перетащите файл сюда».'); });
    });
  }
  function grafikLetter(group, iso) { return dayLetter(group, iso); }
  // the days of a period that the PDF schedule knows, Mon–Fri
  function piDays(from, to) {
    var out = [];
    for (var t = utcOfIso(from); t <= utcOfIso(to); t += DAY) { var wd = new Date(t).getUTCDay(); if (wd > 0 && wd < 6) out.push(isoOfUtc(t)); }
    return out;
  }
  // pairs that give way on a day: on practice days the practice pairs of the timetable stay
  function piGoneRows(group, iso, letter) {
    var sh = DB.sheets[group];
    return ((sh && sh.byDay[iso]) || []).filter(function (r) {
      return !r.ghost && r.disc && !(r.edit && r.edit.kind === 'add' && /практик/i.test(r.disc)) && !(/[UPX]/.test(letter) && Hours.isPractice(r.disc));
    });
  }
  function piReview() {
    var today = todayISO(), items = [];
    if (pi.kind === 'periods') {
      pi.periods.forEach(function (p, idx) {
        var notes = [], gf = grafikFor(p.group), changed = 0, was = {};
        if (DB.groups.length && DB.groups.indexOf(p.group) < 0) notes.push({ level: 'warn', text: 'Группы «' + p.raw + '» нет в списке групп.' });
        if (!gf) notes.push({ level: 'warn', text: 'Группы нет в графике учебного процесса: недели не изменятся, только пары.' });
        var days = piDays(p.from, p.to), gone = [];
        days.forEach(function (d) {
          var cur = grafikLetter(p.group, d);
          if (gf && cur !== p.kind) { changed++; if (cur) was[cur] = (was[cur] || 0) + 1; }
          if (d >= today) piGoneRows(p.group, d, p.kind).forEach(function (r) { gone.push(r); });
        });
        if (!DB.sheets[p.group] && DB.groups.indexOf(p.group) >= 0) notes.push({ level: 'info', text: 'Расписание группы ещё не загружено, пары не проверены.' });
        if (p.to < today) notes.push({ level: 'info', text: 'Период уже прошёл: поменяется только график.' });
        var prev = pi.items[idx];
        items.push({ p: p, days: days, changed: changed, was: was, gone: gone, notes: notes,
          apply: prev ? prev.apply : true, cancel: prev ? prev.cancel : (p.kind !== 'A' && gone.length > 0) });
      });
    } else {
      var byDay = {};
      pi.pairs.forEach(function (x) { var k = x.group + '|' + x.date; (byDay[k] || (byDay[k] = [])).push(x); });
      Object.keys(byDay).sort(function (a, b) { var A = a.split('|'), B = b.split('|'); return A[1] < B[1] ? -1 : A[1] > B[1] ? 1 : cmpRu(A[0], B[0]); }).forEach(function (k, idx) {
        var list = byDay[k], group = list[0].group, date = list[0].date, notes = [];
        if (DB.groups.length && DB.groups.indexOf(group) < 0) notes.push({ level: 'warn', text: 'Группы «' + group + '» нет в списке групп.' });
        if (date < today) notes.push({ level: 'info', text: 'Дата уже прошла.' });
        if (isHoliday(date)) notes.push({ level: 'warn', text: 'Это праздничный день.' });
        var sh = DB.sheets[group], now = ((sh && sh.byDay[date]) || []).filter(function (r) { return !r.ghost && r.disc; });
        var starts = list.map(function (x) { return x.start; });
        var gone = now.filter(function (r) { return starts.indexOf(r.start) < 0; }), replaced = now.filter(function (r) { return starts.indexOf(r.start) >= 0; });
        list.forEach(function (x) {
          splitTeachers(x.teacher).forEach(function (t) {
            Object.keys(DB.sheets).forEach(function (g) {
              if (g === group) return;
              (DB.sheets[g].byDay[date] || []).forEach(function (r) { if (!r.ghost && r.start === x.start && splitTeachers(r.teacher).indexOf(t) >= 0) notes.push({ level: 'warn', text: t + ' в это же время (' + slotLabel(x.start) + ') ведёт пару в группе ' + g + '.' }); });
            });
          });
        });
        var prev = pi.items[idx];
        items.push({ group: group, date: date, list: list, now: now, gone: gone, replaced: replaced, notes: notes, apply: prev ? prev.apply : true, cancel: prev ? prev.cancel : true, letter: list[0].kind });
      });
    }
    pi.items = items;
  }
  function openPracticeImport() {
    pi = { kind: '', items: [], source: '', periods: [], pairs: [] };
    $('pi-file').value = '';
    $('pi-url').value = '';
    $('pi-status').textContent = '';
    $('pi-preview').innerHTML = '';
    $('pi-save').disabled = true;
    $('pi-save').textContent = 'Применить';
    $('pi-dialog').showModal();
    ensureAll();
  }
  function takePracticeTables(tables, source) {
    var periods = piPeriods(tables), pairs = periods.length ? [] : piPairs(tables, source);
    pi = { kind: periods.length ? 'periods' : pairs.length ? 'pairs' : '', items: [], source: source, periods: periods, pairs: pairs };
    var emptyTpl = !pi.kind && tables.some(function (t) { return (t.rows[0] || []).some(function (c) { return /^групп/i.test(cellText(c)); }) && t.rows.slice(1).every(function (r) { return !(r || []).some(function (c) { return cellText(c); }); }); });
    if (emptyTpl) {
      $('pi-status').textContent = 'Таблица прочитана, но в ней нет заполненных строк. Заполните хотя бы одну строку (пример — на листе «Как заполнять» в шаблоне), сохраните файл и загрузите его снова.';
      $('pi-preview').innerHTML = '';
      $('pi-save').disabled = true;
      return;
    }
    if (!pi.kind) {
      $('pi-status').textContent = 'В файле не нашлось ни сроков практики, ни расписания практики по парам. Сроки: колонка «Группа» и колонки «Учебная практика (УП)», «Производственная практика (ПП)» с датами начала и окончания (или «с … по …»), либо колонки «Группа», «Вид», «Начало», «Окончание». Расписание по парам: колонки «Дата», «Пара», «Наименование» (и «Руководитель»), группа — колонкой или в заголовке.';
      $('pi-preview').innerHTML = '';
      $('pi-save').disabled = true;
      return;
    }
    renderPracticeReview();
  }
  function takePracticeFile(f) {
    if (!f) return;
    $('pi-status').textContent = 'Читаю «' + f.name + '»…';
    $('pi-preview').innerHTML = '';
    $('pi-save').disabled = true;
    readDocTables(f).then(function (t) { takePracticeTables(t, f.name); }, function (err) { $('pi-status').textContent = 'Не удалось прочитать файл: ' + err.message + '.'; });
  }
  function letterText(c) { return WEEK_TYPES[c] || PI_NAMES[c] || c; }
  function renderPracticeReview() {
    piReview();
    var it = pi.items, loading = net.bulk ? ' Загружаю расписание групп для проверки: ' + net.bulk.done + ' из ' + net.bulk.total + '…' : '';
    var nDays = 0, nPairs = 0, nGroups = uniq(it.map(function (x) { return x.p ? x.p.group : x.group; })).length;
    var rows;
    if (pi.kind === 'periods') {
      it.forEach(function (x) { if (!x.apply) return; nDays += x.changed; if (x.cancel) nPairs += x.gone.length; });
      $('pi-status').textContent = 'Найдено сроков: ' + it.length + ' для ' + count(nGroups, 'группы', 'групп', 'групп') + '. В графике изменится ' + count(nDays, 'день', 'дня', 'дней') + ', отменится ' + count(nPairs, 'пара', 'пары', 'пар') + '.' + loading;
      rows = it.map(function (x, i) {
        var p = x.p, wasTxt = Object.keys(x.was).map(function (c) { return letterText(c) + ' — ' + x.was[c]; }).join(', ');
        var notes = x.notes.map(function (n) { return '<li class="xi-' + n.level + '">' + esc(n.text) + '</li>'; }).join('');
        var gone = x.gone.length ? '<li><label><input type="checkbox" data-pi-cancel="' + i + '"' + (x.cancel ? ' checked' : '') + '> Отменить ' + count(x.gone.length, 'пару', 'пары', 'пар') + ' группы в эти дни</label>' +
          '<span class="aka">' + esc(uniq(x.gone.map(function (r) { return splitCode(r.disc).name; })).slice(0, 4).join(', ')) + (uniq(x.gone.map(function (r) { return r.disc; })).length > 4 ? '…' : '') + '</span></li>' : '';
        return '<tr' + (x.apply ? '' : ' class="is-off"') + '><td data-label="Применить"><input type="checkbox" data-pi-apply="' + i + '"' + (x.apply ? ' checked' : '') + ' aria-label="Применить ' + esc(p.group) + '"></td>' +
          '<td data-label="Группа"><span class="grp" translate="no">' + esc(p.group) + '</span></td>' +
          '<td class="title" data-label="Что">' + esc(cap(letterText(p.kind))) + (p.label && piKindOfHeader(p.label) !== p.kind ? '' : '') + '</td>' +
          '<td data-label="Когда">' + esc(fmtShort(p.from)) + ' — ' + esc(fmtShort(p.to)) + '</td>' +
          '<td data-label="График">' + (x.changed ? '<span class="badge b-replace">' + esc(count(x.changed, 'день', 'дня', 'дней')) + '</span><span class="aka">было: ' + esc(wasTxt || 'нет в графике') + '</span>' : '<span class="badge s-past">уже так</span>') + '</td>' +
          '<td data-label="Проверка">' + (notes || gone ? '<ul class="xi-notes">' + notes + gone + '</ul>' : '<span class="xi-ok">✓ пар в эти дни нет</span>') + '</td></tr>';
      });
      $('pi-preview').innerHTML = '<p class="hint">Недели графика учебного процесса поменяются на эти сроки (это видно в ленте года, во вкладках «График», «Часы» и «Вычитка»). Пары группы в эти дни будут отменены правкой администратора, кроме пар самой практики. Сроки экзаменов по умолчанию пары не отменяют.</p>' +
        '<div class="table-wrap"><table class="smart xi-table pi-table"><thead><tr><th scope="col"><span class="sr-only">Применить</span></th><th scope="col">Группа</th><th scope="col">Что</th><th scope="col">Когда</th><th scope="col">График</th><th scope="col">Проверка</th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
    } else {
      it.forEach(function (x) { if (!x.apply) return; nDays++; nPairs += x.list.length; });
      $('pi-status').textContent = 'Расписание практики: ' + count(nPairs, 'пара', 'пары', 'пар') + ' на ' + count(it.length, 'день', 'дня', 'дней') + ' для ' + count(nGroups, 'группы', 'групп', 'групп') + '.' + loading;
      rows = it.map(function (x, i) {
        var notes = x.notes.map(function (n) { return '<li class="xi-' + n.level + '">' + esc(n.text) + '</li>'; }).join('');
        var list = x.list.map(function (p) { return '<li>' + esc(slotLabel(p.start)) + ': ' + esc(p.name) + (p.teacher ? ', ' + esc(p.teacher) : '') + '</li>'; }).join('');
        var gone = x.gone.length ? '<li><label><input type="checkbox" data-pi-cancel="' + i + '"' + (x.cancel ? ' checked' : '') + '> Убрать остальные пары дня: ' + esc(x.gone.map(function (r) { return (DB.slotByStart[r.start] || {}).n + ' пара ' + splitCode(r.disc).name; }).join(', ')) + '</label></li>' : '';
        var repl = x.replaced.length ? '<li class="xi-info">Заменится: ' + esc(x.replaced.map(function (r) { return (DB.slotByStart[r.start] || {}).n + ' пара ' + splitCode(r.disc).name; }).join(', ')) + '</li>' : '';
        return '<tr' + (x.apply ? '' : ' class="is-off"') + '><td data-label="Применить"><input type="checkbox" data-pi-apply="' + i + '"' + (x.apply ? ' checked' : '') + ' aria-label="Применить ' + esc(x.group) + ' ' + esc(x.date) + '"></td>' +
          '<td data-label="Группа"><span class="grp" translate="no">' + esc(x.group) + '</span></td>' +
          '<td data-label="Дата">' + esc(dayLabel(x.date)) + '</td>' +
          '<td class="title" data-label="Пары практики"><ul class="xi-notes">' + list + '</ul></td>' +
          '<td data-label="Проверка">' + (notes || gone || repl ? '<ul class="xi-notes">' + repl + gone + notes + '</ul>' : '<span class="xi-ok">✓ других пар в этот день нет</span>') + '</td></tr>';
      });
      $('pi-preview').innerHTML = '<p class="hint">В каждый день из файла пары группы заменятся парами практики, а в графике учебного процесса этот день станет днём практики. Всё сохраняется правками администратора, их можно убрать в разделе «Изменения».</p>' +
        '<div class="table-wrap"><table class="smart xi-table pi-table"><thead><tr><th scope="col"><span class="sr-only">Применить</span></th><th scope="col">Группа</th><th scope="col">Дата</th><th scope="col">Пары практики</th><th scope="col">Проверка</th></tr></thead><tbody>' + rows.join('') + '</tbody></table></div>';
    }
    $('pi-save').disabled = !it.some(function (x) { return x.apply; });
    $('pi-save').textContent = 'Применить' + (nDays && pi.kind === 'periods' ? ': график — ' + count(nDays, 'день', 'дня', 'дней') : '') + (nPairs ? (pi.kind === 'periods' ? ', отменить ' + count(nPairs, 'пару', 'пары', 'пар') : ': ' + count(nPairs, 'пара', 'пары', 'пар') + ' практики') : '');
  }
  // the week strings of the PDF schedule overrides after the given days get their letters
  function piGrafikItems(dayLetters, note) {
    var weeks = {};
    Object.keys(dayLetters).forEach(function (k) {
      var p = k.split('|'), group = p[0], iso = p[1], gf = grafikFor(group);
      if (!gf) return;
      var mon = isoOfUtc(mondayOfUtc(utcOfIso(iso))), wd = isoToDate(iso).getDay();
      if (GRAFIK.data.weeks.indexOf(mon) < 0 || wd < 1 || wd > 5) return;
      var wk = gf.key + '|' + mon;
      if (!weeks[wk]) { var cur = GOV.map[gf.key] && GOV.map[gf.key][mon]; weeks[wk] = { group: gf.key, week: mon, days: (cur && cur.days ? cur.days : '.....').split(''), note: note }; }
      weeks[wk].days[wd - 1] = dayLetters[k];
    });
    return Object.keys(weeks).map(function (k) { var w = weeks[k]; return { group: w.group, week: w.week, days: w.days.join(''), note: w.note }; });
  }
  function rowWas(group, day, start) {
    return JSON.stringify(slotRows(group, day, start).filter(function (x) { return !x.ghost && !x.edit; }).map(function (x) { return { disc: x.disc, teacher: x.teacher, room: x.room, sub: x.sub }; }));
  }
  function cancelEditFor(row, note) {
    return { id: row.edit ? row.edit.id : '', kind: 'cancel', group: row.group, date: row.day, start: row.start, end: row.end, sub: +row.sub || 0, note: note, was: row.edit ? row.edit.was : rowWas(row.group, row.day, row.start) };
  }
  function savePracticeImport() {
    var note = 'практика из «' + (pi.source || 'документа') + '»', letters = {}, edits = [], drop = [], seen = {};
    function cancel(r, why) {
      var k = r.group + '|' + r.day + '|' + r.start + '|' + (r.sub || 0);
      if (seen[k]) return;
      seen[k] = true;
      if (r.edit && r.edit.kind === 'add') drop.push(r.edit.id); else edits.push(cancelEditFor(r, why));
    }
    pi.items.forEach(function (x) {
      if (!x.apply) return;
      if (pi.kind === 'periods') {
        x.days.forEach(function (d) { if (grafikLetter(x.p.group, d) !== x.p.kind) letters[x.p.group + '|' + d] = x.p.kind; });
        if (x.cancel) x.gone.forEach(function (r) { cancel(r, note + ': ' + letterText(x.p.kind)); });
        return;
      }
      if (grafikLetter(x.group, x.date) !== x.letter) letters[x.group + '|' + x.date] = x.letter;
      if (x.cancel) x.gone.forEach(function (r) { cancel(r, note + ': день практики'); });
      x.list.forEach(function (p) {
        var here = slotRows(x.group, x.date, p.start).filter(function (r) { return !r.ghost; });
        var mine = here.filter(function (r) { return r.edit && r.edit.kind !== 'cancel'; })[0];
        edits.push({ id: mine ? mine.edit.id : '', kind: here.length && !(mine && mine.edit.kind === 'add') ? 'replace' : 'add', group: x.group, date: x.date, start: p.start, end: slotEndOf(p.start), sub: 0,
          disc: p.name, teacher: p.teacher, room: p.room, note: note, was: mine ? mine.edit.was || '' : rowWas(x.group, x.date, p.start) });
        seen[x.group + '|' + x.date + '|' + p.start + '|0'] = true;
      });
    });
    var gItems = piGrafikItems(letters, note);
    var btn = $('pi-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    var batch = 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), savedWeeks = [];
    function chunks(list, n) { var out = []; for (var i = 0; i < list.length; i += n) out.push(list.slice(i, i + n)); return out; }
    chunks(gItems, 150).reduce(function (pr, ch) {
      return pr.then(function () { return Api.grafikSave(ch, note).then(function (saved) { setGrafik(saved); savedWeeks = savedWeeks.concat(saved); }); });
    }, Promise.resolve()).then(function () {
      return chunks(edits, 100).reduce(function (pr, ch) { return pr.then(function () { return Api.save(ch, { batch: batch, comment: note }).then(mergeEdits); }); }, Promise.resolve());
    }).then(function () {
      return drop.length ? Api.remove(drop, { comment: note }).then(function (removed) { setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok'); }) : null;
    }).then(function () {
      HIST.status = 'idle';
      $('pi-dialog').close();
      recompute();
      render();
      toast('Практика внесена: в графике ' + count(savedWeeks.length, 'неделя', 'недели', 'недель') + ', правок пар — ' + (edits.length + drop.length) + '. Всё можно вернуть в разделе «Изменения» и на вкладке «График».');
    }, function (err) {
      $('pi-status').textContent = err.message + ' Часть изменений могла сохраниться — проверьте раздел «Изменения».';
      btn.disabled = false;
      btn.textContent = 'Применить';
      recompute();
      render();
    });
  }

  // ---------- reset what was uploaded ----------
  var EXAM_NOTE = /^(экзамен|зачёт|диф\. зачёт): /i, PRACTICE_NOTE = /^практика из «/;
  function resetExams() {
    var cancels = EDITS.list.filter(function (e) { return e.kind === 'cancel' && EXAM_NOTE.test(e.note || ''); });
    var n = EXI.raw.length;
    if (!n && !cancels.length) { toast('Загруженных экзаменов нет: сбрасывать нечего.'); return; }
    if (!confirm('Очистить список экзаменов и зачётов (' + n + ')' + (cancels.length ? ' и вернуть ' + count(cancels.length, 'пару', 'пары', 'пар') + ', отменённые под экзамены' : '') + '?\n\nСписок в базе сайта будет очищен. Прошлый список можно вернуть кнопкой «Откатить» в журнале правок.')) return;
    var btn = $('e-xreset');
    btn.disabled = true;
    (n ? Api.examsImport([], 'сброс загруженных экзаменов') : Promise.resolve()).then(function () {
      return cancels.length ? Api.remove(cancels.map(function (e) { return e.id; }), { comment: 'сброс загруженных экзаменов' }) : null;
    }).then(function () {
      HIST.status = 'idle';
      return loadEdits();
    }).then(function () {
      toast('Сброшено: экзаменов — ' + n + ', возвращено пар — ' + cancels.length + '.');
    }, function (err) { toast(err.message, true); }).then(function () { btn.disabled = false; });
  }
  function resetPractice() {
    var edits = EDITS.list.filter(function (e) { return PRACTICE_NOTE.test(e.note || ''); });
    var weeks = GOV.list.filter(function (g) { return PRACTICE_NOTE.test(g.note || ''); });
    if (!edits.length && !weeks.length) { toast('Загруженной практики нет: сбрасывать нечего.'); return; }
    if (!confirm('Убрать всё, что внесла загрузка практики: ' + count(weeks.length, 'неделю', 'недели', 'недель') + ' в графике и ' + count(edits.length, 'правку', 'правки', 'правок') + ' пар? Расписание и график вернутся как в таблице и PDF.')) return;
    var items = weeks.map(function (g) { return { group: g.group, week: g.week, days: '', note: '' }; }), out = [];
    for (var i = 0; i < items.length; i += 150) out.push(items.slice(i, i + 150));
    out.reduce(function (pr, ch) { return pr.then(function () { return Api.grafikSave(ch, 'сброс загруженной практики').then(function () { setGrafik(ch); }); }); }, Promise.resolve()).then(function () {
      return edits.length ? Api.remove(edits.map(function (e) { return e.id; }), { comment: 'сброс загруженной практики' }).then(function (removed) { setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok'); }) : null;
    }).then(function () {
      HIST.status = 'idle';
      recompute();
      render();
      toast('Практика сброшена: недель в графике — ' + weeks.length + ', правок пар — ' + edits.length + '.');
    }, function (err) { toast(err.message, true); recompute(); render(); });
  }

  // Excel export of the filtered list in the import format (ТЗ БП 1.5 «Экспорт в Excel»)
  var xlsxPromise = null;
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    if (!xlsxPromise) {
      xlsxPromise = new Promise(function (resolve, reject) {
        var s = document.createElement('script');
        s.src = XLSX_URL;
        s.onload = function () { window.XLSX ? resolve(window.XLSX) : reject(new Error('xlsx')); };
        s.onerror = function () { xlsxPromise = null; reject(new Error('xlsx')); };
        document.head.appendChild(s);
      });
    }
    return xlsxPromise;
  }
  function exportExams() {
    var today = todayISO(), keepType = ex.type;
    if (keepType === 'retake') ex.type = 'all';
    var recs = EX.records.filter(function (r) { return matchesExam({ rec: r, date: r.date, time: r.time, retake: false }, today); });
    ex.type = keepType;
    var btn = $('e-export');
    btn.disabled = true;
    loadXlsx().then(function (X) {
      var aoa = [['№', 'Группа', 'Дисциплина', 'Дата и время проведения', 'Аудитория', 'Преподаватель', 'Тип аттестации', 'Дата и время пересдачи']]
        .concat(sortItems(recs.map(function (r) { return { rec: r, date: r.date, time: r.time, retake: false }; })).map(function (it, i) {
          var r = it.rec;
          return [i + 1, r.group, r.discipline, fmtShort(r.date) + (r.time ? ' ' + r.time : ''), r.room, r.teacher, r.type, r.retakeDate ? fmtShort(r.retakeDate) + (r.retakeTime ? ' ' + r.retakeTime : '') : ''];
        }));
      var ws = X.utils.aoa_to_sheet(aoa), wb = X.utils.book_new();
      ws['!cols'] = [5, 10, 52, 22, 12, 24, 14, 22].map(function (w) { return { wch: w }; });
      X.utils.book_append_sheet(wb, ws, 'Экзамены');
      X.writeFile(wb, 'ekzameny' + (ex.group.length === 1 ? '-' + slug(ex.group[0]) : '') + '.xlsx');
    }).catch(function () {
      alert('Не удалось загрузить модуль Excel. Проверьте подключение к интернету.');
    }).then(function () { btn.disabled = false; });
  }
  // ТЗ БП 5.5 «Скачать шаблон»: the header row and two filled examples
  function downloadTemplate() {
    var btn = $('e-template');
    btn.disabled = true;
    loadXlsx().then(function (X) {
      var aoa = [
        ['№', 'Группа', 'Дисциплина / Наименование модуля', 'Дата и время проведения', '№ аудитории', 'Преподаватель (ФИО)', 'Тип аттестации', 'Дата и время пересдачи'],
        [1, 'ИС2-25', 'ЕН.01 Дискретная математика', '23 июня 2027 г 10 часов 00 минут', '305', 'Иванова А.С.', 'Экзамен', '10 июля 2027 г 10 часов 00 минут'],
        [2, 'ИС2-25', 'ОП.02 Архитектура аппаратных средств', '25.06.2027 12:30', '309', 'Петров Д.М.', 'Диф. зачёт', '']
      ];
      var ws = X.utils.aoa_to_sheet(aoa), wb = X.utils.book_new();
      ws['!cols'] = [5, 10, 52, 30, 12, 24, 14, 30].map(function (w) { return { wch: w }; });
      X.utils.book_append_sheet(wb, ws, (CFG.examSheet || 'Экзамены').slice(0, 31));
      X.writeFile(wb, 'shablon-ekzamenov.xlsx');
    }).catch(function () {
      alert('Не удалось загрузить модуль Excel. Проверьте подключение к интернету.');
    }).then(function () { btn.disabled = false; });
  }

  // ---------- column settings dialog ----------
  function renderColList() {
    var order = columnOrder();
    $('col-list').innerHTML = order.map(function (k, i) {
      var c = colByKey(k), locked = k === 'disc';
      return '<li><label><input type="checkbox" data-col="' + k + '"' + (isHidden(k) ? '' : ' checked') + (locked ? ' disabled' : '') + '>' + c.label + (locked ? ' <span class="hint">(всегда)</span>' : '') + '</label>' +
        '<button type="button" class="btn btn-icon" data-move="' + k + ':-1" aria-label="Переместить «' + c.label + '» выше"' + (i === 0 ? ' disabled' : '') + '>↑</button>' +
        '<button type="button" class="btn btn-icon" data-move="' + k + ':1" aria-label="Переместить «' + c.label + '» ниже"' + (i === order.length - 1 ? ' disabled' : '') + '>↓</button></li>';
    }).join('');
  }
  function ensureColPrefs() { if (!colPrefs) colPrefs = { order: columnOrder(), hidden: [] }; }

  // ---------- filter sheet on phones ----------
  function toggleSheet(open) {
    $('e-sheet').classList.toggle('open', open);
    $('e-backdrop').hidden = !open;
    $('e-sheet-open').setAttribute('aria-expanded', open ? 'true' : 'false');
    if (open) $('e-sheet-close').focus(); else $('e-sheet-open').focus();
  }

  // ---------- admin: sign-in, the pair editor, password ----------
  function renderAccount() {
    var u = Api.user(), btn = $('account-btn');
    $('account-label').textContent = u ? u.name : 'Вход для администратора';
    $('account-short').textContent = u ? 'Админ' : 'Вход';
    btn.title = u ? u.name + ': меню администратора' : 'Вход для администратора';
    $('foot-admin').textContent = u ? 'Меню администратора' : 'Вход для администратора';
    btn.classList.toggle('is-in', !!u);
    btn.setAttribute('aria-label', u ? 'Администратор ' + u.name + ': меню' : 'Войти как администратор');
    document.documentElement.classList.toggle('is-admin', !!u);
    $('g-pi-field').hidden = !u;
    $('demo-note').hidden = !(u && Api.mode === 'demo');
  }
  function openLogin() {
    $('login-error').textContent = '';
    $('login-demo').hidden = Api.mode !== 'demo';
    $('login-form').reset();
    $('login-dialog').showModal();
    $('login-name').focus();
  }
  function openAccount() {
    var u = Api.user();
    $('account-who').textContent = u.name + ' (' + u.login + ', ' + u.role + ')';
    $('account-mode').hidden = Api.mode !== 'demo';
    $('account-dialog').showModal();
  }
  // every place a pair can be edited from passes the slot: day, start time and the groups in it
  var ed = null;
  function slotRows(group, day, start) {
    var s = DB.sheets[group];
    return ((s && s.byDay[day]) || []).filter(function (r) { return r.start === start; });
  }
  function datalistFill(id, values) {
    $(id).innerHTML = uniq(values).sort(cmpRu).map(function (v) { return '<option value="' + esc(v) + '"></option>'; }).join('');
  }
  function openEditor(o) {
    var groups = o.groups.filter(Boolean);
    if (!groups.length) return;
    ed = { day: o.day, start: o.start, groups: groups, group: groups[0], all: groups.length > 1, kind: o.kind || '', editId: o.editId || '', chooseGroup: !!o.chooseGroup, presetTeacher: o.teacher || '', presetRoom: o.room || '' };
    $('ed-error').textContent = '';
    $('ed-repeat').value = 'year';
    $('ed-repeat-until').value = Rules.shift(academicEnd(), -1);
    $('ed-repeat-until').min = ed.day;
    $('ed-repeat-until').max = Rules.shift(academicEnd(), -1);
    $('ed-study-only').checked = true;
    var gsel = $('ed-group');
    gsel.innerHTML = (ed.chooseGroup ? '<option value="">Выберите группу для пары</option>' : groups.length > 1 ? '<option value="*">Все группы потока: ' + esc(groups.join(', ')) + '</option>' : '') +
      groups.map(function (g) { return '<option value="' + esc(g) + '">' + esc(g) + '</option>'; }).join('');
    gsel.value = ed.chooseGroup ? '' : groups.length > 1 && !o.editId ? '*' : ed.group;
    $('ed-group-field').hidden = !ed.chooseGroup && groups.length < 2;
    $('ed-slot').innerHTML = DB.slots.map(function (s) { return '<option value="' + esc(s.start) + '">' + s.n + ' пара, ' + esc(s.start) + '–' + esc(s.end) + '</option>'; }).join('');
    $('ed-slot').value = o.start;
    $('ed-move-slot').innerHTML = $('ed-slot').innerHTML;
    $('ed-move-date').value = o.day;
    $('ed-move-slot').value = o.start;
    var rooms = [], discs = [];
    Object.keys(DB.sheets).forEach(function (g) { DB.sheets[g].rows.forEach(function (r) { if (r.room) rooms.push(r.room); }); });
    groups.forEach(function (g) {
      (DB.sheets[g] ? DB.sheets[g].rows : []).forEach(function (r) { if (r.disc) discs.push(r.disc); });
      var p = planFor(g);
      if (p) p.subjects.forEach(function (s) { discs.push(s.name); });
    });
    datalistFill('dl-disc', discs.concat(catalogNames('subject')));
    datalistFill('dl-teacher', DB.teachers);
    datalistFill('dl-room', rooms);
    edLoadSlot();
    $('edit-dialog').showModal();
  }
  // fills the form from what is in the chosen slot now: an edit if there is one, else the lesson
  function edLoadSlot() {
    var g = $('ed-group').value === '*' ? ed.groups[0] : $('ed-group').value;
    ed.group = g;
    ed.start = $('ed-slot').value || ed.start;
    var rows = slotRows(g, ed.day, ed.start);
    var edit = null;
    rows.forEach(function (r) { if (r.edit && (ed.editId ? r.edit.id === ed.editId : r.edit.kind !== 'add') && !edit) edit = r.edit; });
    if (!edit && !ed.editId) {
      var liveRows = rows.filter(function (r) { return !r.ghost; });
      if (liveRows.length === 1 && liveRows[0].edit && liveRows[0].edit.kind === 'add') edit = liveRows[0].edit;
    }
    if (!edit && ed.editId) edit = EDITS.list.filter(function (e) { return e.id === ed.editId; })[0] || null;
    if (ed.chooseGroup) edit = null;
    var base = [];
    rows.forEach(function (r) {
      if (r.edit) { if (r.edit.kind !== 'add') (r.was || []).forEach(function (w) { if (base.indexOf(w) < 0) base.push(w); }); }
      else base.push(r);
    });
    ed.edit = edit;
    ed.base = base;
    var s = DB.slotByStart[ed.start];
    $('ed-title').textContent = cap(WEEKDAYS[isoToDate(ed.day).getDay()]) + ', ' + fmtDayMonth(ed.day) + (s ? ', ' + s.n + ' пара' : '');
    $('ed-now').innerHTML = base.length
      ? 'По основному расписанию: ' + base.map(function (r) { return '<strong>' + esc(splitCode(r.disc).name) + '</strong>' + (r.teacher ? ', ' + esc(r.teacher) : '') + (r.room ? ', ' + esc(roomLabel(r.room)) : '') + (r.sub ? ', ' + r.sub + ' подгруппа' : ''); }).join('; ') + '.'
      : 'По основному расписанию в это время у группы пары нет.';
    if (edit) $('ed-now').innerHTML += ' Сейчас действует правка: ' + esc(edit.kind === 'cancel' ? 'пара отменена' : [edit.disc, edit.teacher, edit.room].filter(Boolean).join(', ')) + '.';
    if (ed.chooseGroup && !g) $('ed-now').textContent = 'Выберите группу для новой пары. ' + (ed.presetTeacher ? 'Преподаватель: ' + ed.presetTeacher : ed.presetRoom ? 'Аудитория: ' + ed.presetRoom : '');
    else if (ed.chooseGroup && base.length) $('ed-now').innerHTML += ' У выбранной группы это время занято. Выберите другую группу или добавьте параллельную пару.';
    // ТЗ, этап 4: «возможность добавить две пары в одно время» — «Ещё пара» ставит пару рядом с этой
    var addOnly = edit && edit.kind === 'add';
    var kinds = addOnly ? ['add', 'move'] : base.length || edit ? ['replace', 'cancel', 'move', 'add'] : ['add'];
    var kind = edit ? edit.kind : ed.kind && kinds.indexOf(ed.kind) >= 0 ? ed.kind : kinds[0];
    ed.occupied = !addOnly && (base.length > 0 || !!edit);
    $('ed-kind').querySelector('[data-value="add"]').textContent = ed.occupied ? 'Ещё пара' : addOnly ? 'Изменить' : 'Поставить пару';
    Array.prototype.forEach.call($('ed-kind').querySelectorAll('button'), function (b) {
      var k = b.getAttribute('data-value');
      b.hidden = kinds.indexOf(k) < 0;
    });
    var src = edit && edit.kind !== 'cancel' ? edit : kind === 'add' && ed.occupied ? {} : base[0] || { disc: '', teacher: '', room: '', sub: 0 };
    $('ed-disc').value = src.disc || '';
    $('ed-teacher').value = ed.presetTeacher || src.teacher || '';
    $('ed-room').value = ed.presetRoom || src.room || '';
    $('ed-sub').value = String(edit ? +edit.sub || 0 : 0);
    $('ed-note').value = edit ? edit.note || '' : '';
    $('ed-remove').hidden = !edit;
    $('ed-remove-series').hidden = !(edit && /^r/.test(edit.batch || ''));
    edSetKind(kind);
  }
  function edSetKind(kind) {
    ed.kind = kind;
    Array.prototype.forEach.call($('ed-kind').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === kind ? 'true' : 'false');
    });
    $('ed-fields').hidden = kind === 'cancel';
    $('ed-move').hidden = kind !== 'move';
    $('ed-slot-field').hidden = kind !== 'add';
    $('ed-save').textContent = kind === 'cancel' ? 'Отменить пару' : kind === 'move' ? 'Перенести пару' : kind === 'add' ? (ed.occupied ? 'Добавить ещё пару' : 'Сохранить пару') : 'Сохранить замену';
    if (ed.chooseGroup) $('ed-save').disabled = !ed.group || !isLive(ed.group);
    edRepeatPreview();
    syncRoving();
  }
  function edSave() {
    if (ed.chooseGroup && (!ed.group || !isLive(ed.group))) { $('ed-error').textContent = 'Выберите группу и дождитесь загрузки её расписания.'; $('ed-group').focus(); return; }
    var targets = $('ed-group').value === '*' ? ed.groups : [ed.group];
    var sub = +$('ed-sub').value || 0, note = $('ed-note').value.trim();
    var fields = { disc: $('ed-disc').value.trim(), teacher: $('ed-teacher').value.trim(), room: $('ed-room').value.trim() };
    if (ed.kind !== 'cancel' && !fields.disc) { $('ed-error').textContent = 'Укажите дисциплину.'; $('ed-disc').focus(); return; }
    var list = [], link = 'l' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    var mineAdd = ed.edit && ed.edit.kind === 'add' ? ed.edit.id : '';
    targets.forEach(function (g) {
      if (!list) return;
      var base = slotRows(g, ed.day, ed.start).filter(function (r) { return !r.edit; }).concat(g === ed.group ? ed.base : []);
      var was = JSON.stringify(uniq(base).map(function (r) { return { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub }; }));
      var mine = g === ed.group && ed.edit && ed.edit.kind !== 'add' ? ed.edit.id : '';
      var slot = { group: g, date: ed.day, start: ed.start, end: (DB.slotByStart[ed.start] || {}).end || '', sub: sub, note: note, was: was };
      if (ed.kind === 'cancel') list.push(Object.assign({ id: mine, kind: 'cancel' }, slot));
      else if (ed.kind === 'move' && mineAdd && g === ed.group) {
        // an added pair just goes to its new time
        var nd = $('ed-move-date').value, ns = $('ed-move-slot').value;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(nd) || !ns) { list = null; return; }
        list.push(Object.assign({}, fields, { id: mineAdd, kind: 'add', group: g, date: nd, start: ns, end: (DB.slotByStart[ns] || {}).end || '', sub: sub, note: note }));
      } else if (ed.kind === 'move') {
        var to = $('ed-move-date').value, toStart = $('ed-move-slot').value;
        if (!/^\d{4}-\d{2}-\d{2}$/.test(to) || !toStart) { list = null; return; }
        if (to === ed.day && toStart === ed.start) { list = null; return; }
        list.push(Object.assign({ id: mine, kind: 'cancel', link: link }, slot));
        var hasThere = slotRows(g, to, toStart).some(function (r) { return !r.ghost; });
        list.push(Object.assign({}, fields, { group: g, date: to, start: toStart, end: (DB.slotByStart[toStart] || {}).end || '', sub: sub, kind: hasThere ? 'replace' : 'add', link: link,
          note: note || 'перенос с ' + fmtShort(ed.day) + ', ' + slotLabel(ed.start), was: JSON.stringify(slotRows(g, to, toStart).filter(function (r) { return !r.ghost; }).map(function (r) { return { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub }; })) }));
      } else if (ed.kind === 'add') list.push(Object.assign({ id: g === ed.group ? mineAdd : '', kind: 'add' }, slot, fields, { was: '' }));
      else list.push(Object.assign({ id: mine, kind: base.length ? 'replace' : 'add' }, slot, fields));
    });
    if (!list) { $('ed-error').textContent = 'Выберите другую дату или пару для переноса.'; return; }
    try {
      list = Rules.expand(list, edRepeatDates(), ed.day, function (next) {
        var existing = EDITS.list.filter(function (x) { return x.group === next.group && x.date === next.date && normTime(x.start) === next.start && +x.sub === +next.sub && (next.kind === 'add' ? x.kind === 'add' && x.disc === next.disc : x.kind !== 'add'); })[0];
        var was = slotRows(next.group, next.date, next.start).filter(function (r) { return !r.ghost && !r.edit; });
        return { id: existing ? existing.id : '', was: JSON.stringify(was.map(function (r) { return { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub }; })) };
      });
    } catch (err) { $('ed-error').textContent = err.message; return; }
    var btn = $('ed-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    $('ed-error').textContent = '';
    var batch = list.length > 1 ? 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6) : '';
    var savedAll = [], chunks = [];
    for (var n = 0; n < list.length; n += 100) chunks.push(list.slice(n, n + 100));
    chunks.reduce(function (chain, chunk) { return chain.then(function () {
      return Api.save(chunk, { batch: batch, comment: list.length > 1 ? 'повторение правки на следующие недели' : '' }).then(function (saved) { savedAll = savedAll.concat(saved); mergeEdits(saved); });
    }); }, Promise.resolve()).then(function () {
      HIST.status = 'idle'; if (state.view === 'history') loadHistory(true);
      $('edit-dialog').close();
      recompute();
      render();
      toast('Сохранено ' + count(savedAll.length, 'изменение', 'изменения', 'изменений') + '. Часы и вычитка пересчитаны.');
    }, function (err) {
      $('ed-error').textContent = err.message + (savedAll.length ? ' Часть правок уже сохранена; их можно убрать одной пачкой в разделе «Изменения».' : '');
      if (!Api.user()) { renderAccount(); }
    }).then(function () { btn.disabled = false; edSetKind(ed.kind); });
  }
  function edRepeatDates() {
    if (!ed || $('ed-repeat').value === 'once') return [ed.day];
    var groups = $('ed-group').value === '*' ? ed.groups : [ed.group];
    var until = $('ed-repeat-until').value;
    return Rules.dates(ed.day, until, $('ed-repeat').value === 'alternate' ? 14 : 7, function (day) {
      var delta = daysBetween(ed.day, day), dest = ed.kind === 'move' && $('ed-move-date').value ? Rules.shift($('ed-move-date').value, delta) : day;
      if (dest > until || dest >= academicEnd()) return false;
      if (!$('ed-study-only').checked) return true;
      return groups.every(function (g) { return [day, dest].every(function (date) {
        if (isHoliday(date) || isoToDate(date).getDay() === 0) return false;
        var gw = grafikWeek(g, isoOfUtc(mondayOfUtc(utcOfIso(date)))), wd = (isoToDate(date).getDay() + 6) % 7;
        if (GRAFIK.data && !gw) return false;
        return !gw || wd >= 5 || gw.days.charAt(wd) === 'T';
      }); });
    });
  }
  function edRepeatPreview() {
    if (!ed) return;
    if (ed.chooseGroup && !ed.group) { $('ed-repeat-preview').textContent = 'После выбора группы будут показаны недели повторения.'; return; }
    var once = $('ed-repeat').value === 'once';
    $('ed-repeat-until').disabled = once;
    $('ed-study-only').disabled = once;
    try { var dates = edRepeatDates(); $('ed-repeat-preview').textContent = 'Правка затронет ' + count(dates.length, 'неделю', 'недели', 'недель') + ', с ' + fmtShort(dates[0]) + ' по ' + fmtShort(dates[dates.length - 1]) + '. Сохранится в базе для всех посетителей.'; }
    catch (err) { $('ed-repeat-preview').textContent = err.message; }
  }
  function mergeEdits(saved) {
    var ids = saved.map(function (e) { return e.id; });
    setEdits(EDITS.list.filter(function (e) { return ids.indexOf(e.id) < 0; }).concat(saved), 'ok');
  }
  // the admin keeps the edit over a changed table: what is there now becomes the replaced lesson
  function keepEdit(id) {
    var e = EDITS.list.filter(function (x) { return x.id === id; })[0];
    if (!e) return;
    var row = ((DB.sheets[e.group] && DB.sheets[e.group].rows) || []).filter(function (r) { return r.edit && r.edit.id === id && r.conflict; })[0];
    var now = row ? row.conflict.now : [];
    var upd = Object.assign({}, e, { was: JSON.stringify(now.map(function (r) { return { disc: r.disc, teacher: r.teacher, room: r.room, sub: r.sub }; })) });
    Api.save([upd], { comment: 'конфликт с таблицей: правка оставлена' }).then(function (saved) {
      mergeEdits(saved);
      HIST.status = 'idle';
      recompute();
      render();
      toast('Правка оставлена, конфликт снят.');
    }, function (err) { toast(err.message, true); });
  }
  function removeEdit(id, after) {
    var e = EDITS.list.filter(function (x) { return x.id === id; })[0];
    return Api.remove([id]).then(function (removed) {
      var linked = e && e.link ? e.link : null;
      setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0 && !(linked && x.link === linked); }), 'ok');
      HIST.status = 'idle'; if (state.view === 'history') loadHistory(true);
      recompute();
      render();
      toast('Правка убрана, пара снова идёт по основному расписанию.');
      if (after) after();
    }, function (err) { toast(err.message, true); });
  }
  var toastTimer = null;
  function toast(text, isError) {
    var el = $('toast');
    el.textContent = text;
    el.classList.toggle('is-error', !!isError);
    el.hidden = false;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.hidden = true; }, 4000);
  }

  // ---------- themes ----------
  // until a visitor picks a theme, it follows the system: Космос in the dark, Рассвет in the light
  var DARK = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : { matches: false };
  function savedTheme() {
    var t = null;
    try { t = localStorage.getItem(THEME_KEY); } catch (e) { t = null; }
    return THEMES.some(function (x) { return x.id === t; }) ? t : null;
  }
  function currentTheme() { return savedTheme() || (DARK.matches ? 'cosmos' : 'dawn'); }
  function applyTheme(id, save) {
    document.documentElement.setAttribute('data-theme', id);
    if (save) { try { localStorage.setItem(THEME_KEY, id); } catch (e) { /* storage unavailable */ } }
    var meta = document.querySelector('meta[name="theme-color"]:not([media])');
    if (meta) meta.setAttribute('content', getComputedStyle(document.documentElement).getPropertyValue('--topbar').trim() || '#1f2a7a');
    renderThemes();
  }
  function renderThemes() {
    var cur = currentTheme();
    $('theme-list').innerHTML = THEMES.map(function (t) {
      return '<li><button type="button" class="theme-opt" data-theme-pick="' + t.id + '" aria-pressed="' + (t.id === cur ? 'true' : 'false') + '">' +
        '<span class="swatch" data-theme="' + t.id + '" aria-hidden="true"><i></i><i></i><i></i></span><span class="theme-name">' + esc(t.name) + '</span><span class="theme-note">' + esc(t.note) + '</span></button></li>';
    }).join('');
  }

  // ---------- page ----------
  function termLabel() {
    var t = new Date(), y = t.getFullYear(), m = t.getMonth();
    var start = m >= 6 ? y : y - 1;
    var autumn = m >= 8 || m === 0;
    return start + '–' + (start + 1) + ' учебный год · ' + (autumn ? 'I семестр (осенний)' : 'II семестр (весенний)');
  }
  function renderStatus() {
    var el = $('status'), text, st;
    if (state.view === 'exams') {
      if (EX.status === 'loading') { st = 'loading'; text = 'Загружаю расписание экзаменов из базы расписания…'; }
      else if (EX.status === 'error') { st = 'error'; text = 'Не удалось загрузить расписание экзаменов из базы расписания.'; }
      else if (EX.source === 'cache') { st = 'stale'; text = 'Google-таблица не ответила. Показана сохранённая копия, загруженная ' + fmtStamp(EX.at) + '.'; }
      else {
        st = 'ok';
        var docs = EXF.records.length ? 'графики экзаменов и зачётной недели' + (EXF.period ? ', ' + EXF.period : '') : '';
        if (CFG.dataSource === 'database') text = 'Экзамены, зачёты и пересдачи из общей базы сайта, обновлены ' + fmtStamp(EXI.at) + '. Записей: ' + EX.records.length + '.';
        else if (EXS.status === 'ok') text = 'Свежие данные из базы расписания, загружены ' + fmtStamp(EX.at) + (docs ? '; к ним добавлены ' + docs : '') + '.';
        else if (docs) text = 'Показаны ' + docs + '.';
        else text = 'Google-таблица прочитана ' + fmtStamp(EX.at) + ': листа с экзаменами в ней пока нет.';
      }
    } else {
      var shown = [];
      if (state.mode === 'group' && state.group && DB.sheets[state.group]) shown = [DB.sheets[state.group]];
      if (state.mode !== 'group') shown = Object.keys(DB.sheets).map(function (g) { return DB.sheets[g]; });
      var anyFailed = net.listFailed || Object.keys(net.failed).length > 0;
      if (net.bulk) { st = 'loading'; text = 'Загружаю расписание групп из базы расписания: ' + net.bulk.done + ' из ' + net.bulk.total + '…'; }
      else if (net.pending > 0) { st = 'loading'; text = 'Загружаю свежее расписание из базы расписания…'; }
      else if (anyFailed) {
        var cached = shown.filter(function (s) { return s.source === 'cache'; });
        var hasData = state.mode === 'group' ? DB.groups.length > 0 && (!state.group || !!DB.sheets[state.group]) : shown.length > 0;
        if (hasData) {
          st = 'stale';
          var oldest = cached.length ? Math.min.apply(null, cached.map(function (s) { return s.at; })) : 0;
          text = 'Google-таблица не ответила. ' + (oldest ? 'Показана сохранённая копия, загруженная ' + fmtStamp(oldest) + '.' : 'Показана сохранённая копия.');
        } else { st = 'error'; text = 'Не удалось загрузить расписание из базы расписания.'; }
      } else if (net.listLive) {
        var live = shown.filter(function (s) { return s.source === 'live'; });
        var at = live.length ? Math.max.apply(null, live.map(function (s) { return s.at; })) : Date.now();
        st = 'ok'; text = 'Свежие данные из базы расписания, загружены ' + fmtStamp(at) + '.' + (syncMinutes() ? ' Сайт сам перечитывает таблицу каждые ' + count(syncMinutes(), 'минуту', 'минуты', 'минут') + '.' : '');
      } else { st = 'loading'; text = 'Загружаю расписание из базы расписания…'; }
    }
    if (state.view === 'hours' && PLAN.source === 'cache' && st === 'ok') { st = 'stale'; text = 'Учебный план не обновился: показана сохранённая копия от ' + fmtStamp(PLAN.at) + '.'; }
    if (EDITS.status === 'error') { if (st === 'ok') st = 'stale'; text += ' Замены от администраторов не загрузились.'; }
    else if (EDITS.status === 'cache') { if (st === 'ok') st = 'stale'; text += ' Замены показаны по сохранённой копии от ' + fmtStamp(EDITS.at) + '.'; }
    else if (EDITS.status === 'ok' && EDITS.list.length && st === 'ok') text += ' Учтены замены администраторов.';
    var guards = Object.keys(net.guards);
    if (guards.length) { if (st === 'ok') st = 'stale'; text += ' Не разобралось листов: ' + guards.length + ', для них показана прошлая версия.'; }
    if (net.refreshing) {
      st = 'loading';
      text = (isAdmin() ? 'Импортирую таблицу' : 'Обновляю расписание') + (net.refreshing.total ? ': ' + net.refreshing.done + ' из ' + net.refreshing.total + ' листов…' : '…');
    }
    el.setAttribute('data-state', st);
    $('status-text').textContent = text;
    $('retry').hidden = !(st === 'stale' || st === 'error') || !!net.refreshing;
    var rb = $('refresh');
    rb.disabled = !!net.refreshing;
    rb.textContent = isAdmin() ? 'Обновить и импортировать' : 'Обновить';
    rb.title = isAdmin() ? 'Заново прочитать таблицу и учебный план, сравнить с прошлым импортом и записать изменения в журнал' : 'Заново прочитать таблицу и учебный план';
    // ТЗ, этап 2: «админ получает предупреждение»
    var gb = $('guards'), notes = Object.keys(net.notes);
    gb.hidden = !(isAdmin() && (guards.length || notes.length));
    if (!gb.hidden) {
      gb.innerHTML = (guards.length ? '<p><strong>Предупреждение для администратора: парсер не справился</strong></p><ul>' + guards.map(function (n) { return '<li>' + esc(net.guards[n]) + '</li>'; }).join('') + '</ul>' +
        '<p class="hint">Данные этих листов не затёрты. При «Обновить и импортировать» предупреждение попадёт в журнал «Изменения в базе расписания».</p>' : '') +
        (notes.length ? '<p><strong>Проверьте листы</strong></p><ul>' + notes.map(function (n) { return '<li>' + esc(net.notes[n]) + '</li>'; }).join('') + '</ul>' : '');
    }
  }

  function render() {
    if (state.view !== 'lessons' && state.mode === 'room') state.mode = 'group';
    resolvePendingHash();
    var td = new Date();
    $('today-label').textContent = WEEKDAYS[td.getDay()][0].toUpperCase() + WEEKDAYS[td.getDay()].slice(1) + ', ' + fmtDayMonth(todayISO());
    $('term-label').textContent = termLabel();
    $('year').textContent = td.getFullYear();
    $('sysline').textContent = 'Системная информация: версия ' + SITE_VERSION + ' · ' + termLabel();

    var v = state.view, isExams = v === 'exams';
    VIEWS.forEach(function (k) {
      $('view-' + k).hidden = k !== v;
      $('tab-' + k).setAttribute('aria-selected', k === v ? 'true' : 'false');
    });
    // who the page is about: shared by lessons and hours
    var who = v === 'lessons' || v === 'hours';
    $('who').hidden = !who;
    $('half-field').hidden = v !== 'hours';
    $('page-title').textContent = VIEW_TITLES[v][0];
    $('page-lede').textContent = VIEW_TITLES[v][1];
    document.title = VIEW_TITLES[v][0] + ' · ТехноПортал';
    var upcoming = EDITS.list.filter(function (e) { return e.date >= todayISO(); }).length;
    $('tab-history-count').textContent = upcoming ? String(upcoming) : '';
    $('stamp').textContent = isExams
      ? (EX.status === 'ok' ? count(uniq(EX.records.map(function (r) { return r.group; })).length, 'группа', 'группы', 'групп') + ', ' + count(EX.records.length, 'запись', 'записи', 'записей') : '')
      : (DB.groups.length ? count(DB.groups.length, 'группа', 'группы', 'групп') + (DB.weeks.length ? ', ' + count(DB.weeks.length, 'неделя', 'недели', 'недель') + ' с ' + fmtDayMonth(DB.weeks[0].iso) : '') : '');
    renderStatus();
    renderAccount();
    if (window.RaspAdminTools) window.RaspAdminTools.render();
    if (who) renderWho();
    if (v === 'lessons') renderLessons();
    else if (v === 'hours') renderHours();
    else if (v === 'balance') renderBalance();
    else if (v === 'grafik') renderGrafik();
    else if (v === 'history') renderHistory();
    else renderExams();
    savePrefs();
    writeUrl();
    syncRoving();
    animateEntrance();
    queueThumbs();
  }
  // ---------- motion ----------
  var REDUCED = window.matchMedia ? window.matchMedia('(prefers-reduced-motion: reduce)') : { matches: false };
  // the chosen option of a switch and the open tab sit on one plate that slides, so a choice moves instead of blinking
  var thumbFrame = 0;
  function syncThumbs() {
    thumbFrame = 0;
    Array.prototype.forEach.call(document.querySelectorAll('.seg, .views'), function (g) {
      var t = g.querySelector(':scope > .seg-thumb');
      if (!g.offsetWidth) { if (t) t.removeAttribute('data-on'); return; }
      if (!t) {
        t = document.createElement('span');
        t.className = 'seg-thumb';
        t.setAttribute('aria-hidden', 'true');
        g.insertBefore(t, g.firstChild);
        g.classList.add('has-thumb');
      }
      var on = Array.prototype.filter.call(g.querySelectorAll('[aria-checked="true"], [aria-selected="true"]'), function (b) { return b.offsetWidth; })[0];
      if (!on) { t.removeAttribute('data-on'); return; }
      // a plate that was hidden appears in place instead of flying in from where it was
      var jump = !t.hasAttribute('data-on');
      if (jump) t.classList.add('no-anim');
      t.style.setProperty('--x', on.offsetLeft + 'px');
      t.style.setProperty('--y', on.offsetTop + 'px');
      t.style.setProperty('--w', on.offsetWidth + 'px');
      t.style.setProperty('--h', on.offsetHeight + 'px');
      t.setAttribute('data-on', '');
      if (jump) { void t.offsetWidth; t.classList.remove('no-anim'); }
    });
  }
  // a timer, not a frame callback: frames pause in background tabs, and the plates must be in place when the tab shows
  function queueThumbs() { if (!thumbFrame) thumbFrame = setTimeout(syncThumbs, 0); }
  // when the page starts showing something else, its blocks rise in one after another;
  // the minute-by-minute refresh of the same page does not replay it
  var ENTER = '.dcard, .tablebox, .hsum, .ribbon, .now, .reminder, .empty, .book, .tdiff, .guards, .course, .admin-bar';
  var enterKey = '', enterView = '', enterTimer = 0;
  function stopEntrance() {
    Array.prototype.forEach.call(document.querySelectorAll('.is-entering'), function (el) { el.classList.remove('is-entering'); });
  }
  function animateEntrance() {
    var view = $('view-' + state.view);
    var blocks = Array.prototype.filter.call(view.querySelectorAll(ENTER), function (b) { return b.offsetParent !== null; });
    var main = blocks.filter(function (b) { return b.matches('.dcard, .tablebox, .hsum, .book, .course, .empty'); }).length;
    var key = [state.view, state.mode, state.group, state.teacher, state.room, state.week, state.span, state.day, state.half,
      state.exPage, state.exTeacher, state.bookGroup, state.bookHalf, isAdmin(), main].join('|');
    if (key === enterKey) return;
    enterKey = key;
    if (REDUCED.matches) return;
    stopEntrance();
    blocks.forEach(function (b, i) { b.style.setProperty('--i', Math.min(i, 10)); });
    Array.prototype.forEach.call(view.querySelectorAll('.rw'), function (b, i) { b.style.setProperty('--k', i); });
    var els = [view];
    if (enterView !== state.view) els.push(document.querySelector('.headline'));
    enterView = state.view;
    void view.offsetWidth;
    els.forEach(function (el) { el.classList.add('is-entering'); });
    clearTimeout(enterTimer);
    enterTimer = setTimeout(stopEntrance, 1800);
  }
  // ARIA tabs and radio groups: Tab reaches the chosen button, arrows move between the rest
  var ROVE = '[role="tab"], [role="radio"]';
  function syncRoving() {
    Array.prototype.forEach.call(document.querySelectorAll('[role="tablist"], [role="radiogroup"]'), function (g) {
      var items = Array.prototype.filter.call(g.querySelectorAll(ROVE), function (b) { return !b.hidden; });
      var on = items.filter(function (b) { return b.getAttribute('aria-selected') === 'true' || b.getAttribute('aria-checked') === 'true'; })[0] || items[0];
      items.forEach(function (b) { b.tabIndex = b === on ? 0 : -1; });
    });
  }
  document.addEventListener('keydown', function (e) {
    var keys = ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End'];
    var b = e.target.closest && e.target.closest(ROVE);
    if (!b || keys.indexOf(e.key) < 0) return;
    var g = b.closest('[role="tablist"], [role="radiogroup"]');
    if (!g) return;
    var items = Array.prototype.filter.call(g.querySelectorAll(ROVE), function (x) { return !x.hidden && !x.disabled; });
    var i = items.indexOf(b), n = items.length;
    var j = e.key === 'Home' ? 0 : e.key === 'End' ? n - 1 : (i + (e.key === 'ArrowLeft' || e.key === 'ArrowUp' ? -1 : 1) + n) % n;
    e.preventDefault();
    items[j].click(); // the choice follows the focus
    items[j].focus();
    syncRoving();
  });
  function renderWho() {
    $('group-field').hidden = state.mode !== 'group';
    $('teacher-field').hidden = state.mode !== 'teacher';
    $('room-field').hidden = state.mode !== 'room';
    $('room').value = state.room;
    $('mode').querySelector('[data-value="room"]').hidden = state.view !== 'lessons';
    $('group').value = state.group;
    $('teacher').value = state.teacher;
    Array.prototype.forEach.call($('mode').querySelectorAll('button'), function (b) {
      b.setAttribute('aria-checked', b.getAttribute('data-value') === state.mode ? 'true' : 'false');
    });
    if (state.view !== 'lessons') $('weeknav').hidden = true;
    $('src-field').hidden = state.view !== 'lessons';
    $('span-field').hidden = state.view !== 'lessons';
    $('date-field').hidden = state.view !== 'lessons';
    $('show-src').checked = showSrc;
  }

  // ---------- events ----------
  function switchView(view) {
    if (view !== 'lessons' && state.mode === 'room') state.mode = 'group';
    if (view === state.view) return;
    if (view === 'exams' && !ex.group.length && !ex.teacher.length) {
      if (state.mode === 'group' && state.group) ex.group = [state.group];
      if (state.mode === 'teacher' && state.teacher) ex.teacher = [state.teacher];
      fillExamOptions();
    }
    if ((view === 'lessons' || view === 'hours') && state.view === 'exams') {
      if (ex.group.length === 1 && !state.group && DB.groups.indexOf(ex.group[0]) >= 0) { state.mode = 'group'; state.group = ex.group[0]; }
      else if (ex.teacher.length === 1 && !state.teacher) { state.mode = 'teacher'; state.teacher = ex.teacher[0]; }
    }
    if (view === 'history' && !HIST.group && state.mode === 'group' && state.group && EDITS.list.some(function (e) { return e.group === state.group; })) HIST.group = state.group;
    state.view = view;
    render();
    ensureView();
    window.scrollTo(0, 0);
  }
  $('half').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b) { state.half = +b.getAttribute('data-value'); render(); }
  });
  $('g-course').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b) { gcourse = b.getAttribute('data-value'); render(); }
  });
  $('grafik-body').addEventListener('click', function (e) {
    var ge = e.target.closest('[data-gedit]'), cell = e.target.closest('td[data-gk]');
    if (ge) { openGrafikEditor(ge.getAttribute('data-gedit')); return; }
    if (cell) { openGrafikEditor(cell.getAttribute('data-gk'), +cell.getAttribute('data-wi')); return; }
    var b = e.target.closest('[data-open-group]');
    if (!b) return;
    state.mode = 'group';
    state.group = b.getAttribute('data-open-group');
    state.week = '';
    switchView('lessons');
  });
  $('ribbon').addEventListener('click', function (e) {
    var b = e.target.closest('[data-rweek]');
    if (!b) return;
    var w = b.getAttribute('data-rweek'), i = nearestWeekIndex(w);
    state.week = i >= 0 ? DB.weeks[i].iso : w;
    render();
  });
  function openEditById(id) {
    var e0 = EDITS.list.filter(function (x) { return x.id === id; })[0];
    if (!e0) return;
    ensureGroup(e0.group);
    openEditor({ day: e0.date, start: normTime(e0.start), groups: [e0.group], editId: e0.id });
  }
  function dropEdit(id) {
    var e0 = EDITS.list.filter(function (x) { return x.id === id; })[0];
    if (e0 && confirm('Убрать правку? Пара ' + fmtDayMonth(e0.date) + ' у группы ' + e0.group + ' снова пойдёт по основному расписанию.')) removeEdit(id);
  }
  $('week').addEventListener('click', function (e) {
    var b = e.target.closest('[data-edit-day]'), a = e.target.closest('[data-add-day]');
    var oe = e.target.closest('[data-open-edit]'), ke = e.target.closest('[data-keep-edit]'), de = e.target.closest('[data-drop-edit]');
    if (oe) { openEditById(oe.getAttribute('data-open-edit')); return; }
    if (ke) { keepEdit(ke.getAttribute('data-keep-edit')); return; }
    if (de) { dropEdit(de.getAttribute('data-drop-edit')); return; }
    var sel = selection();
    if (!sel) return;
    if (b) {
      var day = b.getAttribute('data-edit-day'), start = b.getAttribute('data-edit-start');
      var groups = sel.mode === 'group' ? [sel.name] : uniq(rowsForDay(day, sel, true).filter(function (r) { return r.start === start; }).map(function (r) { return r.group; })).sort(cmpRu);
      openEditor({ day: day, start: start, groups: groups });
    } else if (a) {
      var aday = a.getAttribute('data-add-day'), astart = a.getAttribute('data-add-start');
      if (!astart) {
        var busy = rowsForDay(aday, sel).map(function (r) { return r.start; });
        var free = DB.slots.filter(function (s) { return busy.indexOf(s.start) < 0; })[0] || DB.slots[0];
        astart = free ? free.start : '8:30';
      }
      openEditor({ day: aday, start: astart, groups: sel.mode !== 'group' ? DB.groups : [sel.name], kind: 'add', chooseGroup: sel.mode !== 'group', teacher: sel.mode === 'teacher' ? sel.name : '', room: sel.mode === 'room' ? sel.name : '' });
    }
  });
  $('history-body').addEventListener('click', function (e) {
    var o = e.target.closest('[data-open-edit]'), r = e.target.closest('[data-remove-edit]');
    var ke = e.target.closest('[data-keep-edit]'), de = e.target.closest('[data-drop-edit]');
    var rb = e.target.closest('[data-remove-batch]'), rv = e.target.closest('[data-revert]');
    if (e.target.closest('[data-history-retry]')) { loadHistory(true); return; }
    if (ke) { keepEdit(ke.getAttribute('data-keep-edit')); return; }
    if (de) { dropEdit(de.getAttribute('data-drop-edit')); return; }
    var dm = e.target.closest('[data-drop-many]');
    if (dm) {
      var kind = dm.getAttribute('data-drop-many'), ids = [];
      if (kind === 'intable') ids = EDITS.list.filter(function (x) { return IN_TABLE[x.id]; }).map(function (x) { return x.id; });
      else Object.keys(DB.sheets).forEach(function (g) { DB.sheets[g].rows.forEach(function (r) { if (r.conflict && r.edit && !r.away && ids.indexOf(r.edit.id) < 0) ids.push(r.edit.id); }); });
      if (!ids.length) return;
      if (!confirm(kind === 'intable' ? 'Убрать ' + count(ids.length, 'правку', 'правки', 'правок') + ', которые уже есть в таблице?' : 'Во всех конфликтах взять расписание из таблицы? Будет убрано ' + count(ids.length, 'правка', 'правки', 'правок') + '.')) return;
      dm.disabled = true;
      Api.remove(ids, { comment: kind === 'intable' ? 'уже внесено в таблицу' : 'конфликт: взято из таблицы' }).then(function (removed) {
        setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok');
        HIST.status = 'idle'; loadHistory(true);
        recompute(); render();
        toast('Убрано ' + count(removed.length, 'правка', 'правки', 'правок') + ': пары идут по таблице.');
      }, function (err) { dm.disabled = false; toast(err.message, true); });
      return;
    }
    if (rb) {
      var batch = rb.getAttribute('data-remove-batch'), n = EDITS.list.filter(function (x) { return x.batch === batch; }).length;
      if (!confirm('Убрать всю массовую замену (' + count(n, 'пара', 'пары', 'пар') + ')? Пары снова пойдут по основному расписанию.')) return;
      Api.remove([], { batch: batch }).then(function (removed) {
        setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok');
        HIST.status = 'idle'; loadHistory(true);
        recompute(); render();
        toast('Массовая замена убрана: ' + count(removed.length, 'пара', 'пары', 'пар') + '.');
      }, function (err) { toast(err.message, true); });
      return;
    }
    if (rv) {
      if (!confirm('Откатить эту запись журнала? Правка вернётся к состоянию «было».')) return;
      Api.revert(rv.getAttribute('data-revert')).then(function () {
        HIST.status = 'idle';
        return loadEdits().then(function () { loadHistory(true); toast('Откат выполнен.'); });
      }, function (err) { toast(err.message, true); });
      return;
    }
    var id = (o || r) ? (o || r).getAttribute(o ? 'data-open-edit' : 'data-remove-edit') : '';
    var ed0 = EDITS.list.filter(function (x) { return x.id === id; })[0];
    if (!ed0) return;
    if (o) {
      ensureGroup(ed0.group);
      openEditor({ day: ed0.date, start: normTime(ed0.start), groups: [ed0.group], editId: ed0.id });
    } else if (confirm('Убрать правку? Пара ' + fmtDayMonth(ed0.date) + ' у группы ' + ed0.group + ' снова пойдёт по основному расписанию.')) removeEdit(id);
  });
  $('h-group').addEventListener('change', function (e) { HIST.group = e.target.value; render(); });
  $('h-type').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b) { histType = b.getAttribute('data-value'); render(); }
  });
  $('show-src').addEventListener('change', function (e) {
    showSrc = e.target.checked;
    try { localStorage.setItem(SOURCES_KEY, showSrc ? '1' : '0'); } catch (err) { /* storage unavailable */ }
    render();
  });
  $('h-empty-action').addEventListener('click', retry);

  function openAdmin() { if (Api.user()) openAccount(); else openLogin(); }
  $('account-btn').addEventListener('click', openAdmin);
  $('foot-admin').addEventListener('click', function (e) { e.preventDefault(); openAdmin(); });
  $('login-cancel').addEventListener('click', function () { $('login-dialog').close(); });
  $('login-form').addEventListener('submit', function (e) {
    e.preventDefault();
    var btn = $('login-submit');
    btn.disabled = true;
    btn.textContent = 'Вхожу…';
    $('login-error').textContent = '';
    Api.login($('login-name').value.trim(), $('login-pass').value).then(function (u) {
      $('login-dialog').close();
      HIST.status = 'idle';
      render();
      if (state.view === 'history') loadHistory(true);
      toast('Вы вошли как ' + u.name + '. Нажмите на карандаш у пары, чтобы изменить её.');
      loadEdits();
    }, function (err) {
      var d = err.data || {};
      $('login-error').textContent = err.message + (d.attemptsLeft != null && d.attemptsLeft > 0 ? '. Осталось попыток: ' + d.attemptsLeft + '.' : '');
      $('login-name').focus();
    }).then(function () { btn.disabled = false; btn.textContent = 'Войти'; });
  });
  $('account-close').addEventListener('click', function () { $('account-dialog').close(); });
  $('account-logout').addEventListener('click', function () {
    Api.logout().then(function () {
      $('account-dialog').close();
      HIST.status = 'idle';
      render();
      toast('Вы вышли из режима администратора.');
      loadEdits();
    });
  });
  $('account-password').addEventListener('click', function () {
    $('account-dialog').close();
    $('pass-form').reset();
    $('pass-error').textContent = '';
    $('pass-dialog').showModal();
  });
  $('pass-cancel').addEventListener('click', function () { $('pass-dialog').close(); });
  $('pass-form').addEventListener('submit', function (e) {
    e.preventDefault();
    var next = $('pass-new').value;
    if (next !== $('pass-again').value) { $('pass-error').textContent = 'Новый пароль и повтор не совпадают.'; $('pass-again').focus(); return; }
    if (next.length < 8) { $('pass-error').textContent = 'Новый пароль — не короче 8 символов.'; $('pass-new').focus(); return; }
    var btn = $('pass-submit');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    Api.password($('pass-old').value, next).then(function () {
      $('pass-dialog').close();
      toast('Пароль изменён.');
    }, function (err) { $('pass-error').textContent = err.message; $('pass-old').focus(); }).then(function () {
      btn.disabled = false;
      btn.textContent = 'Сменить пароль';
    });
  });
  $('theme-btn').addEventListener('click', function () { renderThemes(); $('theme-dialog').showModal(); });
  $('theme-list').addEventListener('click', function (e) {
    var b = e.target.closest('[data-theme-pick]');
    if (b) {
      var pick = b.getAttribute('data-theme-pick'), done = false;
      var apply = function () { if (!done) { done = true; applyTheme(pick, true); } };
      // the old and new themes crossfade; if the browser cannot draw the crossfade soon (a background tab),
      // the theme simply switches and the stalled crossfade is dropped so the page does not stay frozen on the old picture
      if (document.startViewTransition && !REDUCED.matches) {
        var vt = document.startViewTransition(apply), over = false;
        vt.finished.then(function () { over = true; }, function () { over = true; });
        vt.ready.catch(function () { /* a dropped crossfade is expected here */ });
        setTimeout(function () { apply(); if (!over) vt.skipTransition(); }, 700);
      } else apply();
    }
  });
  $('ed-kind').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b) edSetKind(b.getAttribute('data-value'));
  });
  $('ed-group').addEventListener('change', function () {
    edLoadSlot();
    if (!ed.chooseGroup || !ed.group || isLive(ed.group)) return;
    var editor = ed, group = ed.group;
    $('ed-save').disabled = true; $('ed-error').textContent = 'Загружаю расписание выбранной группы…';
    loadSheet(group).then(function () {
      recompute();
      if (ed !== editor || $('ed-group').value !== group) return;
      if (!isLive(group)) { $('ed-error').textContent = 'Не удалось загрузить группу. Выберите её снова, чтобы повторить.'; return; }
      $('ed-error').textContent = ''; edLoadSlot();
    });
  });
  $('ed-slot').addEventListener('change', edLoadSlot);
  ['ed-repeat-until', 'ed-study-only', 'ed-move-date', 'ed-move-slot'].forEach(function (id) { $(id).addEventListener('change', edRepeatPreview); });
  $('ed-repeat').addEventListener('change', function () {
    var mode = $('ed-repeat').value;
    $('ed-repeat-until').value = mode === 'term' && ed.day < half2Of(ed.group) ? Rules.shift(half2Of(ed.group), -1) : Rules.shift(academicEnd(), -1);
    edRepeatPreview();
  });
  $('ed-save').addEventListener('click', edSave);
  $('ed-remove-series').addEventListener('click', function () {
    if (!ed || !ed.edit || !ed.edit.batch) return;
    var batch = ed.edit.batch, n = EDITS.list.filter(function (x) { return x.batch === batch; }).length;
    if (!confirm('Убрать всю серию (' + count(n, 'правка', 'правки', 'правок') + ')?')) return;
    $('ed-remove-series').disabled = true;
    Api.remove([], { batch: batch, comment: 'убрана серия повторяющихся правок' }).then(function (removed) {
      setEdits(EDITS.list.filter(function (x) { return removed.indexOf(x.id) < 0; }), 'ok');
      HIST.status = 'idle'; recompute(); render(); $('edit-dialog').close(); toast('Серия убрана. Часы и вычитка пересчитаны.');
    }, function (err) { $('ed-error').textContent = err.message; }).then(function () { $('ed-remove-series').disabled = false; });
  });
  $('ed-cancel').addEventListener('click', function () { $('edit-dialog').close(); });
  $('ed-remove').addEventListener('click', function () {
    if (ed && ed.edit && confirm('Убрать правку? Пара снова пойдёт по основному расписанию.')) removeEdit(ed.edit.id, function () { $('edit-dialog').close(); });
  });
  $('balance-body').addEventListener('click', function (e) {
    var t;
    if ((t = e.target.closest('[data-bal-group]'))) {
      state.group = t.getAttribute('data-bal-group'); state.balAll = false; bal.open = ''; bal.wantAll = false;
      render(); ensureView();
      var d = $('balance-body').querySelector('.hsum');
      if (d && e.target.closest('.bal-all')) d.scrollIntoView({ block: 'start', behavior: REDUCED.matches ? 'auto' : 'smooth' });
      return;
    }
    if (e.target.closest('[data-bal-all]')) { state.balAll = !state.balAll; bal.open = ''; render(); ensureView(); return; }
    if ((t = e.target.closest('[data-bal-open]'))) {
      var k = t.getAttribute('data-bal-open');
      bal.open = bal.open === k ? '' : k;
      // free time has to be checked against the teacher's pairs in every group
      if (bal.open && /:add$/.test(bal.open)) { bal.wantAll = true; ensureView(); }
      render();
      return;
    }
    if ((t = e.target.closest('[data-bal-put]'))) { t.disabled = true; balPut(t.getAttribute('data-bal-put')); return; }
    if ((t = e.target.closest('[data-bal-cut]'))) { balCut(t.getAttribute('data-bal-cut')); return; }
    if ((t = e.target.closest('[data-bal-csv]'))) { balCsv(t.getAttribute('data-bal-csv')); return; }
    if (e.target.closest('[data-bal-until-reset]')) { delete bal.until[balanceKey(balanceGroup())]; render(); return; }
    if ((t = e.target.closest('[data-bal-until-set]'))) { bal.until[balanceKey(balanceGroup())] = t.getAttribute('data-bal-until-set'); render(); return; }
    if (e.target.closest('[data-bal-retry]')) retry();
  });
  $('balance-body').addEventListener('change', function (e) {
    if (e.target.id === 'bal-half') { state.half = +e.target.value; bal.open = ''; balMemo = {}; render(); ensureView(); return; }
    if (e.target.id === 'bal-group') { if (e.target.value) { state.group = e.target.value; state.mode = 'group'; state.balAll = false; bal.open = ''; render(); ensureView(); } return; }
    if (e.target.id === 'bal-until') {
      var v = e.target.value;
      if (/^\d{4}-\d{2}-\d{2}$/.test(v)) bal.until[balanceKey(balanceGroup())] = v; else delete bal.until[balanceKey(balanceGroup())];
      render();
    }
  });
  document.querySelector('.views').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-view]');
    if (b) switchView(b.getAttribute('data-view'));
  });
  $('mode').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    state.mode = b.getAttribute('data-value');
    render();
    ensureView();
  });
  $('group').addEventListener('change', function (e) { state.group = e.target.value; state.week = ''; render(); ensureView(); });
  $('room').addEventListener('change', function (e) { state.room = e.target.value; render(); });
  $('teacher').addEventListener('change', function (e) { state.teacher = e.target.value; render(); });
  $('picker-body').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-group]');
    if (b) { state.group = b.getAttribute('data-group'); render(); ensureView(); }
  });
  function shiftWeek(delta) {
    var day = spanDay();
    if (day) { state.span = 'day'; state.day = isoOfUtc(utcOfIso(day) + delta * DAY); render(); return; }
    var wi = nearestWeekIndex(state.week) + delta;
    if (wi >= 0 && wi < DB.weeks.length) { state.week = DB.weeks[wi].iso; render(); }
  }
  $('wk-prev').addEventListener('click', function () { shiftWeek(-1); });
  $('wk-next').addEventListener('click', function () { shiftWeek(1); });
  $('wk-today').addEventListener('click', function () {
    if (spanDay()) { state.span = 'today'; state.day = ''; } else state.week = '';
    render();
  });
  $('span').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    var day = spanDay();
    state.span = b.getAttribute('data-value');
    state.day = '';
    if (state.span === 'week' && day) state.week = isoOfUtc(mondayOfUtc(utcOfIso(day)));
    render();
  });
  $('goto-date').addEventListener('change', function (e) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(e.target.value)) return;
    state.span = 'day';
    state.day = e.target.value;
    render();
  });
  $('tdiff').addEventListener('click', function (e) {
    if (!e.target.closest('[data-tdiff-hide]') || !state.group) return;
    try { localStorage.removeItem(DIFF_KEY + state.group); } catch (err) { /* storage unavailable */ }
    render();
  });
  $('refresh').addEventListener('click', refreshAll);
  $('bulk-open').addEventListener('click', function () { openBulk(); });
  $('bulk-kind').addEventListener('click', function (e) { var b = e.target.closest('button'); if (b) bulkSetKind(b.getAttribute('data-value')); });
  ['bulk-teacher', 'bulk-from', 'bulk-to'].forEach(function (id) { $(id).addEventListener('change', bulkPreview); });
  $('bulk-save').addEventListener('click', bulkSave);
  $('bulk-cancel').addEventListener('click', function () { $('bulk-dialog').close(); });
  $('gf-group').addEventListener('change', gfLoad);
  $('gf-from').addEventListener('change', gfLoad);
  $('gf-to').addEventListener('change', gfSummary);
  $('gf-all').addEventListener('change', function (e) {
    if (!e.target.value) return;
    for (var j = 0; j < 5; j++) $('gf-d' + j).value = e.target.value;
    e.target.value = '';
  });
  $('gf-save').addEventListener('click', function () { gfSave(false); });
  $('gf-reset').addEventListener('click', function () { gfSave(true); });
  $('gf-cancel').addEventListener('click', function () { $('gf-dialog').close(); });
  $('ex-page').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (!b) return;
    state.exPage = b.getAttribute('data-value');
    if (state.exPage === 'book' && !state.bookGroup) state.bookGroup = state.group;
    if (state.exPage === 'teacher' && !state.exTeacher) state.exTeacher = state.teacher;
    render();
  });
  $('view-exams').addEventListener('change', function (e) {
    if (e.target.id === 'book-group') { state.bookGroup = e.target.value; render(); }
    if (e.target.id === 'ext-teacher') { state.exTeacher = e.target.value; render(); }
  });
  $('ex-book').addEventListener('click', function (e) {
    var b = e.target.closest('#book-half button');
    if (b) { state.bookHalf = +b.getAttribute('data-value'); render(); }
  });
  $('e-import').addEventListener('click', openExamImport);
  $('e-xreset').addEventListener('click', resetExams);
  document.addEventListener('click', function (e) { if (e.target.closest('[data-pi-reset]')) resetPractice(); });
  function takeExamFile(f) {
    if (!f) return;
    $('xi-status').textContent = 'Читаю «' + f.name + '»…';
    $('xi-preview').innerHTML = '';
    $('xi-save').disabled = true;
    readExamFile(f).then(function (res) { showExamPreview(res, f.name); }, function (err) { $('xi-status').textContent = 'Не удалось прочитать файл: ' + err.message + '.'; });
  }
  $('xi-file').addEventListener('change', function (e) { takeExamFile(e.target.files && e.target.files[0]); });
  // the page's group fills empty group cells only for the file being loaded in the dialog
  $('xi-dialog').addEventListener('close', function () { xiGroup = ''; });
  document.addEventListener('click', function (e) { if (e.target.closest('[data-pi-open]')) openPracticeImport(); });
  $('pi-file').addEventListener('change', function (e) { takePracticeFile(e.target.files && e.target.files[0]); });
  $('pi-read').addEventListener('click', function () {
    var u = $('pi-url').value.trim();
    if (!u) { $('pi-status').textContent = 'Вставьте ссылку на Google-таблицу, Google-документ или файл на Google Диске.'; $('pi-url').focus(); return; }
    $('pi-status').textContent = 'Открываю ссылку…';
    $('pi-preview').innerHTML = '';
    readGoogleLink(u).then(function (r) { takePracticeTables(r.tables, r.name); }, function (err) { $('pi-status').textContent = 'Не удалось открыть ссылку: ' + err.message; });
  });
  $('pi-save').addEventListener('click', savePracticeImport);
  $('pi-cancel').addEventListener('click', function () { $('pi-dialog').close(); });
  $('pi-preview').addEventListener('change', function (e) {
    var a = e.target.getAttribute('data-pi-apply'), c = e.target.getAttribute('data-pi-cancel');
    if (a != null && pi.items[+a]) pi.items[+a].apply = e.target.checked;
    if (c != null && pi.items[+c]) pi.items[+c].cancel = e.target.checked;
    if (a != null || c != null) renderPracticeReview();
  });
  $('xi-preview').addEventListener('change', function (e) {
    var k = e.target.getAttribute('data-xi-cancel');
    if (!k || !xi.review) return;
    var p = k.split(':'), it = xi.review.items[+p[0]], c = it && it.cancels[+p[1]];
    if (c) { c.on = e.target.checked; renderExamReview(); }
  });
  // a file dropped on the dialog, or on the exams page by an admin, is read right away
  function hasFiles(e) { return e.dataTransfer && Array.prototype.indexOf.call(e.dataTransfer.types || [], 'Files') >= 0; }
  // a file dropped on the exams page goes to the exam import, on the timetable or the schedule — to the practice import
  function dropTarget() {
    if ($('xi-dialog').open) return 'xi';
    if ($('pi-dialog').open) return 'pi';
    if (!isAdmin()) return '';
    return state.view === 'exams' ? 'xi' : state.view === 'lessons' || state.view === 'grafik' ? 'pi' : '';
  }
  document.addEventListener('dragover', function (e) {
    var t = hasFiles(e) ? dropTarget() : '';
    if (!t) return;
    e.preventDefault();
    $(t + '-drop').classList.add('is-over');
  });
  document.addEventListener('dragleave', function (e) { if (!e.relatedTarget) { $('xi-drop').classList.remove('is-over'); $('pi-drop').classList.remove('is-over'); } });
  document.addEventListener('drop', function (e) {
    var t = hasFiles(e) ? dropTarget() : '';
    if (!t) return;
    e.preventDefault();
    $(t + '-drop').classList.remove('is-over');
    if (t === 'xi') { if (!$('xi-dialog').open) openExamImport(); takeExamFile(e.dataTransfer.files[0]); }
    else { if (!$('pi-dialog').open) openPracticeImport(); takePracticeFile(e.dataTransfer.files[0]); }
  });
  $('xi-read').addEventListener('click', function () {
    var u = $('xi-url').value.trim();
    if (!u) { $('xi-status').textContent = 'Вставьте ссылку на Google-таблицу, Google-документ или файл на Google Диске.'; $('xi-url').focus(); return; }
    $('xi-status').textContent = 'Открываю ссылку…';
    $('xi-preview').innerHTML = '';
    readGoogleLink(u).then(function (r) { showExamPreview(examsFromTables(r.tables), r.name); }, function (err) { $('xi-status').textContent = 'Не удалось открыть ссылку: ' + err.message; });
  });
  $('xi-save').addEventListener('click', saveExamImport);
  $('xi-cancel').addEventListener('click', function () { $('xi-dialog').close(); });
  $('empty-action').addEventListener('click', function () {
    var a = this.getAttribute('data-action') || '';
    if (a === 'retry') retry();
    else if (a.indexOf('week:') === 0) { state.week = a.slice(5); render(); }
  });
  $('reminder').addEventListener('click', function (e) {
    if (e.target.id !== 'reminder-link') return;
    e.preventDefault();
    ex = exDefaults();
    ex.status = 'future';
    if (state.mode === 'group') ex.group = [state.group]; else ex.teacher = [state.teacher];
    state.view = 'exams';
    fillExamOptions();
    render();
  });
  $('retry').addEventListener('click', retry);

  function setEx(key, value) {
    ex[key] = value;
    ex.page = 1; ex.rpage = 1;
    render();
  }
  $('e-from').addEventListener('change', function (e) { setEx('from', e.target.value); });
  $('e-to').addEventListener('change', function (e) { setEx('to', e.target.value); });
  $('e-search').addEventListener('input', function (e) { setEx('search', e.target.value); });
  [['e-type', 'type'], ['e-status', 'status'], ['e-half', 'half']].forEach(function (pair) {
    $(pair[0]).addEventListener('click', function (e) {
      var b = e.target.closest('button');
      if (b) setEx(pair[1], b.getAttribute('data-value'));
    });
  });
  function resetExFilters() {
    var keep = { ordering: ex.ordering, pageSize: ex.pageSize };
    ex = Object.assign(exDefaults(), keep);
    fillExamOptions();
    render();
  }
  $('e-reset').addEventListener('click', resetExFilters);
  $('e-export').addEventListener('click', exportExams);
  $('e-template').addEventListener('click', downloadTemplate);
  $('e-sheet-open').addEventListener('click', function () { toggleSheet(true); });
  $('e-sheet-close').addEventListener('click', function () { toggleSheet(false); });
  $('e-backdrop').addEventListener('click', function () { toggleSheet(false); });

  $('e-cols').addEventListener('click', function () { renderColList(); $('col-dialog').showModal(); });
  $('col-list').addEventListener('change', function (e) {
    var k = e.target.getAttribute('data-col');
    if (!k) return;
    ensureColPrefs();
    colPrefs.hidden = colPrefs.hidden.filter(function (h) { return h !== k; });
    if (!e.target.checked) colPrefs.hidden.push(k);
    saveColPrefs();
    render();
  });
  $('col-list').addEventListener('click', function (e) {
    var b = e.target.closest('button[data-move]');
    if (!b) return;
    var parts = b.getAttribute('data-move').split(':'), k = parts[0], d = +parts[1];
    ensureColPrefs();
    var order = columnOrder(), i = order.indexOf(k), j = i + d;
    if (j < 0 || j >= order.length) return;
    order.splice(i, 1);
    order.splice(j, 0, k);
    colPrefs.order = order;
    saveColPrefs();
    renderColList();
    var again = $('col-list').querySelector('[data-move="' + k + ':' + d + '"]');
    if (again && !again.disabled) again.focus();
    render();
  });
  $('col-reset').addEventListener('click', function () { colPrefs = null; saveColPrefs(); renderColList(); render(); });

  $('e-fav-save').addEventListener('click', function () { $('fav-name').value = ''; $('fav-dialog').showModal(); });
  $('fav-cancel').addEventListener('click', function () { $('fav-dialog').close(); });
  $('fav-form').addEventListener('submit', function () {
    var name = $('fav-name').value.trim();
    if (!name) return;
    var favs = loadFavs();
    favs.push({ name: name, qs: exParams(ex, new URLSearchParams(), false).toString() });
    saveFavs(favs);
    renderFavs();
  });
  $('e-favs').addEventListener('click', function (e) {
    var a = e.target.closest('button[data-fav]'), d = e.target.closest('button[data-fav-del]');
    var favs = loadFavs();
    if (a && favs[+a.getAttribute('data-fav')]) {
      var keepSize = ex.pageSize;
      ex = parseExParams(new URLSearchParams(favs[+a.getAttribute('data-fav')].qs));
      ex.pageSize = keepSize;
      fillExamOptions();
      render();
    }
    if (d) { favs.splice(+d.getAttribute('data-fav-del'), 1); saveFavs(favs); renderFavs(); }
  });
  $('view-exams').addEventListener('click', function (e) {
    var s = e.target.closest('button[data-sort]');
    if (s) {
      // click cycles: ascending ▲ → descending ▼ → no sorting (ТЗ БП 1.5)
      var k = s.getAttribute('data-sort');
      ex.ordering = ex.ordering === k ? '-' + k : ex.ordering === '-' + k ? '' : k;
      render();
      return;
    }
    var p = e.target.closest('button[data-page]');
    if (p && !p.disabled) {
      var parts = p.getAttribute('data-page').split(':');
      ex[parts[0]] = +parts[1];
      render();
      return;
    }
    if (e.target.closest('[data-reset]')) resetExFilters();
    if (e.target.closest('[data-template]')) downloadTemplate();
    if (e.target.closest('[data-retry]')) retry();
  });
  $('view-exams').addEventListener('change', function (e) {
    if (!e.target.hasAttribute('data-size')) return;
    ex.pageSize = +e.target.value;
    ex.page = 1; ex.rpage = 1;
    render();
  });
  window.addEventListener('popstate', function () { readUrl(); fillOptions(); render(); ensureView(); });

  // ---------- start ----------
  (function boot() {
    applyTheme(currentTheme());
    if (DARK.addEventListener) DARK.addEventListener('change', function () { if (!savedTheme()) applyTheme(currentTheme()); });
    ['ms-group', 'ms-teacher', 'ms-disc'].forEach(msInit);
    renderFavs();
    readUrl();
    var lc = cacheGet(CFG.groupListSheet), cc = CFG.calendarSheet ? cacheGet(CFG.calendarSheet) : null, ec = CFG.examSheet ? cacheGet(CFG.examSheet) : null;
    var pc = CFG.planSheet ? cacheGet(CFG.planSheet) : null;
    if (lc) applyGroupList(lc.csv);
    if (cc) applyCalendar(cc.csv);
    if (ec) applyExams(ec.csv, 'cache', ec.t);
    if (pc) applyPlan(pc.csv, 'cache', pc.t);
    if (Api.mode !== 'off') {
      try {
        var saved = JSON.parse(localStorage.getItem(STATE_KEY) || 'null');
        if (saved && saved.state) { setState(saved.state, 'loading'); EDITS.at = saved.t; }
      } catch (e) { /* storage unavailable */ }
    }
    if (state.view === 'lessons' && state.mode === 'group' && state.group) {
      var c = cacheGet(state.group);
      if (c) applySheet(state.group, c.csv, 'cache', c.t);
    }
    recompute();
    fillOptions();
    render();
    loadLists();
    loadGrafik();
    loadEdits();
    if (window.ResizeObserver) {
      new ResizeObserver(function () {
        document.documentElement.style.setProperty('--who-h', ($('who').offsetHeight + 16) + 'px');
      }).observe($('who'));
    }
    // inside ТехноПортал (?embed=1) the page tells the portal its height, so the frame has no scrollbar
    if (document.documentElement.classList.contains('embed') && window.parent !== window && window.ResizeObserver) {
      new ResizeObserver(function () {
        window.parent.postMessage({ type: 'raspisanie:height', height: document.documentElement.scrollHeight }, '*');
      }).observe(document.body);
    }
    // keep "today" and the current pair fresh while the page stays open
    setInterval(function () { if (state.view === 'lessons' && !document.querySelector('dialog[open]')) render(); }, 60000);
    // new replacements from the admins show up without reloading the page
    setInterval(function () { if (!document.hidden && !document.querySelector('dialog[open]')) loadEdits(); }, 180000);
    // and so do the changes in the college table
    setInterval(function () { if (!document.hidden && syncDue() && !document.querySelector('dialog[open]:not(#xi-dialog)')) syncNow(); }, 30000);
    document.addEventListener('visibilitychange', function () {
      if (document.hidden) return;
      if (syncMinutes() && Date.now() - net.synced > 60000) syncNow();
      if (Date.now() - EDITS.at > 60000) loadEdits();
    });
    // the sliding plates follow every change of choice, a dialog opening, the fonts arriving and the window size
    if (window.MutationObserver) new MutationObserver(queueThumbs).observe(document.body, { subtree: true, attributes: true, attributeFilter: ['aria-checked', 'aria-selected', 'hidden', 'open'] });
    window.addEventListener('resize', queueThumbs);
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(queueThumbs);
    if (wantLogin) openAdmin();
    // when the site is opened from this computer, expose the data for checks in the console
    if (/^(localhost|127\.0\.0\.1)$/.test(location.hostname)) {
      window.__raspisanie = { DB: DB, net: net, ensureAll: ensureAll, get EX() { return EX; }, applyExams: applyExams, get EDITS() { return EDITS; }, get PLAN() { return PLAN; }, get GRAFIK() { return GRAFIK; }, countFor: countFor, planFor: planFor, half2Of: half2Of, get GOV() { return GOV; }, get TJ() { return TJ; }, refreshAll: refreshAll, syncNow: syncNow, diffRows: diffRows, examsFromTables: examsFromTables, tablesFromHtml: tablesFromHtml, readExamFile: readExamFile, showExamPreview: showExamPreview, openExamImport: openExamImport, piPeriods: piPeriods, piPairs: piPairs, readDocTables: readDocTables };
    }
  })();
})();

