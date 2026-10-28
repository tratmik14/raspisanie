(function () {
  'use strict';

  var CFG = window.SCHEDULE_CONFIG || {};
  var DAY = 864e5;
  var PREFS_KEY = 'raspisanie.prefs.v1';
  var CACHE_KEY = 'raspisanie.sheet.v1:';
  var CONCURRENCY = 6;
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  var MONTHS_GEN = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  var WEEKDAYS = ['воскресенье', 'понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота'];
  var WEEKDAYS_MON = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота', 'Воскресенье'];
  var WEEKDAY_NAMES = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];
  var TEACHER_RE = /^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.?)?$/;
  var GROUP_RE = /^[А-ЯЁA-Z][А-ЯЁA-Zа-яё0-9]*-\d{2}[А-ЯЁа-яё]?$/;
  var TYPE_KEY = { 'Экзамен': 'exam', 'Диф. зачёт': 'diff', 'Зачёт': 'credit' };
  var SITE_VERSION = '3.4';
  var COLS_KEY = 'raspisanie.columns.v1';
  var FAVS_KEY = 'raspisanie.favs.v1';
  var THEME_KEY = 'raspisanie.theme.v1';
  var STATE_KEY = 'raspisanie.state.v1';
  var GOOD_KEY = 'raspisanie.good.v1';
  var DIFF_KEY = 'raspisanie.diff.v1:';
  var SOURCES_KEY = 'raspisanie.sources.v1';
  var VIEWS = ['lessons', 'hours', 'grafik', 'exams', 'history'];
  var VIEW_TITLES = {
    lessons: ['Расписание занятий', 'Пары на каждую неделю семестра: для группы или для преподавателя.'],
    hours: ['Часы по предметам', 'Сколько часов по плану, сколько уже проведено и сколько осталось.'],
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
    { id: 'portal', name: 'Портал', note: 'цвета ТехноПортала, светлая или тёмная по системе' },
    { id: 'notebook', name: 'Тетрадь', note: 'клетка, синие чернила и красные поля' },
    { id: 'board', name: 'Доска', note: 'зелёная доска и мел' },
    { id: 'cosmos', name: 'Космос', note: 'ночное небо и оранжевый «Восток»' },
    { id: 'blueprint', name: 'Чертёж', note: 'синька и белые линии' },
    { id: 'lilac', name: 'Сирень', note: 'светлая, сиреневая с мятным' }
  ];
  var Api = window.RaspApi;
  var Hours = window.RaspHours;
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
  function splitTeachers(s) { return String(s || '').split(/\s*[,;]\s*/).filter(Boolean); }
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
  function sheetUrl(name) {
    return 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(CFG.sheetId) +
      '/gviz/tq?tqx=out:csv&headers=0&sheet=' + encodeURIComponent(name);
  }
  function fetchSheet(name) {
    return fetch(sheetUrl(name), { cache: 'no-store' }).then(function (res) {
      var type = res.headers.get('content-type') || '';
      if (!res.ok || type.indexOf('csv') < 0) throw new Error('Google Sheets answered ' + res.status + ' ' + type);
      return res.text();
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
      if (/препод|фио|экзаменатор|педагог/.test(h)) { if (map.teacher == null) map.teacher = i; return; }
      if (/тип|вид|форм/.test(h)) { if (map.type == null) map.type = i; return; }
      if (/дата|срок/.test(h)) { if (map.date == null) map.date = i; return; }
      if (/время|начал/.test(h)) { if (map.time == null) map.time = i; return; }
    });
    return map;
  }
  // Returns null when the sheet does not look like an exam list (Google answers with the
  // first sheet of the workbook when the requested one does not exist).
  function parseExamSheet(rows, defaults) {
    var headerIdx = -1, map = null;
    for (var i = 0; i < Math.min(rows.length, 20); i++) {
      var mp = mapExamHeader(rows[i] || []);
      var hits = ['group', 'discipline', 'date', 'teacher', 'type'].filter(function (k) { return mp[k] != null; }).length;
      if (hits >= 3) { headerIdx = i; map = mp; break; }
    }
    if (!map) return null;
    var out = { records: [], skipped: 0 }, seen = {};
    function get(row, key) { return map[key] == null ? '' : row[map[key]]; }
    for (var r = headerIdx + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var discipline = cellText(get(row, 'discipline')), group = cellText(get(row, 'group')) || (defaults && defaults.group) || '';
      var dateRaw = get(row, 'date');
      if (!discipline && !cellText(dateRaw)) continue;
      var dt = parseExamDate(dateRaw), type = normType(get(row, 'type'));
      if (!group || !discipline || !dt || dt.error || !type) { out.skipped++; continue; }
      var time = dt.time || (map.time != null ? parseTimeText(get(row, 'time')) : '');
      var rt = parseExamDate(get(row, 'retake')), retakeDate = '', retakeTime = '';
      if (rt && !rt.error) {
        retakeDate = rt.date;
        retakeTime = rt.time || (map.retakeTime != null ? parseTimeText(get(row, 'retakeTime')) : '');
        if ((retakeDate + retakeTime) <= (dt.date + time)) { retakeDate = ''; retakeTime = ''; }
      }
      var key = [group, discipline, dt.date, time].join('|').toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      out.records.push({
        group: group, discipline: discipline, date: dt.date, time: time, room: cellText(get(row, 'room')),
        teacher: cellText(get(row, 'teacher')), type: type, retakeDate: retakeDate, retakeTime: retakeTime
      });
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
  var net = { pending: 0, failed: {}, listLive: false, listFailed: false, bulk: null, guards: {}, refreshing: false };
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
    var seen = {}, records = [];
    EXS.records.concat(EXF.records, EXI.records).forEach(function (r) {
      var k = [r.group, r.discipline, r.date, r.time].join('|').toLowerCase();
      if (!seen[k]) { seen[k] = true; records.push(r); }
    });
    var extra = EXF.records.length + EXI.records.length > 0;
    EX = {
      status: extra ? 'ok' : EXS.status === 'loading' || EXF.status === 'loading' ? 'loading' : EXS.status, records: records, skipped: EXS.skipped,
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
    try { localStorage.setItem(STATE_KEY, JSON.stringify({ t: Date.now(), state: { edits: EDITS.list, grafik: GOV.list, exams: EXI.raw } })); }
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
    return key ? PLAN.groups[key] : null;
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
  function applyEdits(group, rows) {
    var list = EDITS.list.filter(function (e) { return e.group === group; });
    if (!list.length) return rows;
    var out = rows.slice();
    list.slice().sort(function (a, b) { return (a.kind === 'add' ? 1 : 0) - (b.kind === 'add' ? 1 : 0); }).forEach(function (e) {
      var start = normTime(e.start), was = [], sub = +e.sub || 0;
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
      var cells = [].concat.apply([], scan.blocks.map(function (b) { return b.cells; }));
      DB.weeks.forEach(function (w) { emit(w.ms, cells, true); });
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
    sheet.rows = applyEdits(group, rows);
    sheet.byDay = {};
    sheet.rows.forEach(function (r) { (sheet.byDay[r.day] || (sheet.byDay[r.day] = [])).push(r); });
  }

  // ---------- state and the address bar (ТЗ БП 1.5: filters live in query parameters) ----------
  // span: 'week' | 'today' | 'tomorrow' | 'day' (a date from the calendar, in state.day)
  // exPage: 'list' | 'book' (зачётка) | 'teacher'
  var state = { view: 'lessons', mode: 'group', group: '', teacher: '', week: '', half: null, span: 'week', day: '', exPage: 'list', bookGroup: '', bookHalf: 0, exTeacher: '' };
  var ex = exDefaults();
  var pendingHash = ''; // an old #group link that waits for the group list
  var wantLogin = false; // #admin opens the login right away
  try {
    var saved = JSON.parse(localStorage.getItem(PREFS_KEY) || 'null');
    if (saved && typeof saved === 'object') {
      ['mode', 'group', 'teacher'].forEach(function (k) { if (typeof saved[k] === 'string') state[k] = saved[k]; });
      if (saved.span === 'today' || saved.span === 'tomorrow') state.span = saved.span;
    }
  } catch (e) { /* storage unavailable */ }
  if (state.mode !== 'teacher') state.mode = 'group';
  function savePrefs() {
    try { localStorage.setItem(PREFS_KEY, JSON.stringify({ mode: state.mode, group: state.group, teacher: state.teacher, span: state.span === 'day' ? 'week' : state.span })); }
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
      if (p.get('teacher')) { state.mode = 'teacher'; state.teacher = p.get('teacher'); }
      else if (p.get('group')) { state.mode = 'group'; state.group = p.get('group'); }
      // a link to a week opens that week, whatever the visitor chose last time
      if (/^\d{4}-\d{2}-\d{2}$/.test(p.get('week') || '')) { state.week = p.get('week'); state.span = 'week'; }
      if (/^[012]$/.test(p.get('half') || '')) state.half = +p.get('half');
      if (p.get('span') === 'today' || p.get('span') === 'tomorrow' || p.get('span') === 'week') state.span = p.get('span');
      if (/^\d{4}-\d{2}-\d{2}$/.test(p.get('day') || '')) { state.span = 'day'; state.day = p.get('day'); }
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
        if (state.mode === 'teacher' && state.teacher) p.set('teacher', state.teacher);
        else if (state.mode === 'group' && state.group) p.set('group', state.group);
      }
      if (state.view === 'lessons') {
        if (state.span === 'day' && state.day) p.set('day', state.day);
        else if (state.span === 'today' || state.span === 'tomorrow') p.set('span', state.span);
        else if (state.week && state.week !== currentWeekIso()) p.set('week', state.week);
      }
      if (state.view === 'hours' && state.half != null) p.set('half', state.half);
    }
    if (document.documentElement.classList.contains('embed')) p.set('embed', '1');
    var qs = p.toString(), target = location.pathname + (qs ? '?' + qs : '');
    if (target === location.pathname + location.search && !location.hash) return;
    try { history.replaceState(null, '', target); } catch (e) { /* ignore */ }
  }

  // ---------- loading ----------
  function isLive(group) { return !!DB.sheets[group] && (DB.sheets[group].source === 'live' || !!DB.sheets[group].fresh); }
  function allLive() { return DB.groups.length > 0 && DB.groups.every(isLive); }
  // ---------- protection against a broken sheet (ТЗ, этап 2) ----------
  // A sheet that suddenly parses into nothing (the layout changed, a tab was renamed) does not replace
  // the last good copy: the page keeps that copy, and admins see a warning and get it in the import journal.
  var good = {};
  try { good = JSON.parse(localStorage.getItem(GOOD_KEY) || '{}') || {}; } catch (e) { good = {}; }
  function brokenParse(name, size, what) {
    if (!size) return 'На листе «' + name + '» не нашлось ' + what;
    var before = good[name] || 0;
    if (before >= 20 && size < before * 0.3) return 'На листе «' + name + '» ' + what.replace(/^ни одно[йг]о? /, '') + ' стало в ' + Math.round(before / size) + ' раз меньше, чем в прошлый раз';
    return '';
  }
  function scanSize(scan) { return scan.blocks.reduce(function (a, b) { return a + b.cells.length; }, 0); }
  function acceptSheet(name, csv, size, what) {
    var why = brokenParse(name, size, what);
    if (why) { net.guards[name] = why + '. Показана прошлая версия: проверьте, не изменился ли формат листа.'; return false; }
    delete net.guards[name];
    cachePut(name, csv);
    good[name] = size;
    try { localStorage.setItem(GOOD_KEY, JSON.stringify(good)); } catch (e) { /* storage unavailable */ }
    return true;
  }
  function loadSheet(group) {
    net.pending++;
    return fetchSheet(group).then(function (csv) {
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
      resolvePendingHash();
      if (state.teacher && DB.teachers.indexOf(state.teacher) < 0) state.teacher = '';
      fillOptions();
      render();
      if ($('bulk-dialog').open) { bulkFillTeachers(); bulkPreview(); }
    });
  }
  function ensureView() {
    if (state.view === 'history') { loadHistory(); loadTableJournal(); }
    if (state.view !== 'lessons' && state.view !== 'hours') return;
    if (state.mode === 'teacher') ensureAll();
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
    }).then(function () { render(); });
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
      (DB.sheets[g].byDay[iso] || []).forEach(function (r) { if (r.teacher === sel.name && (!r.ghost || withGhosts)) out.push(r); });
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
  // «Основная» — пара из Google-таблицы; «Замена» — администратор поставил вместо неё другую;
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
    if (sel.mode === 'teacher') e.groups.sort(cmpRu).forEach(function (g) { meta.push('<span class="grp" translate="no">' + esc(g) + '</span>'); });
    else if (e.teacher) meta.push('<span>' + esc(e.teacher) + '</span>');
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
    var admin = isAdmin(), canAdd = admin && sel.mode === 'group';
    var head = '<header class="dhead"><span class="dname">' + WEEKDAYS_MON[wd] + '</span><span class="ddate">' + fmtDayMonth(dayIso) + '</span>' +
      (rel ? '<span class="rel">' + rel + '</span>' : '') + '</header>';
    var addBtn = canAdd ? '<button type="button" class="btn btn-quiet btn-small add-pair" data-add-day="' + dayIso + '">' + PLUS + 'Добавить пару</button>' : '';
    if (!rows.length) return '<article class="dcard' + cls + '">' + head + '<p class="dempty">Пар нет</p>' + addBtn + '</article>';
    var byPair = {};
    rows.forEach(function (r) { var n = slotOf(r).n; (byPair[n] || (byPair[n] = [])).push(r); });
    var nums = Object.keys(byPair).map(Number).sort(function (a, b) { return a - b; });
    var nowMin = new Date().getHours() * 60 + new Date().getMinutes();
    var items = [];
    for (var p = nums[0]; p <= nums[nums.length - 1]; p++) {
      var s = DB.slots[p - 1];
      if (!byPair[p]) {
        items.push('<li class="pair gap"><span class="pnum"><b>' + p + '</b></span><span>Окно' + (s ? ', ' + esc(s.start) + '–' + esc(s.end) : '') + '</span>' +
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
    return '<article class="dcard' + cls + '">' + head + '<ol class="pairs">' + items.join('') + '</ol>' + addBtn + '</article>';
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

  // ТЗ БП 5.8: a reminder about exams within the next 3 calendar days.
  // Credits and retakes are not shown in the general schedule, so the reminder leaves them out too.
  function renderReminder(sel, today) {
    var el = $('reminder');
    if (EX.status !== 'ok' || !sel) { el.hidden = true; return; }
    var soon = examItems().filter(function (it) {
      if (it.retake || TYPE_KEY[it.rec.type] !== 'exam') return false;
      var d = daysBetween(today, it.date);
      if (d < 0 || d > 3) return false;
      return sel.mode === 'group' ? it.rec.group === sel.name : splitTeachers(it.rec.teacher).indexOf(sel.name) >= 0;
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
      if (net.listFailed) showEmpty('Не удалось загрузить расписание из Google-таблицы. Проверьте подключение к интернету. Если ошибка повторяется, проверьте, что таблица открыта по ссылке.', 'Повторить', 'retry');
      else showEmpty('Загружаю расписание из Google-таблицы…');
      return;
    }
    if (!sel) {
      if (state.mode === 'group') $('picker').hidden = false;
      else showEmpty(net.bulk ? 'Собираю список преподавателей: загружено ' + net.bulk.done + ' из ' + net.bulk.total + ' групп…' : 'Выберите преподавателя, чтобы увидеть его пары.');
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
    if (total || (isAdmin() && sel.mode === 'group')) weekEl.innerHTML = cards.join('');
    else {
      var other = findWeekWithLessons(sel, wi);
      var why = 'На эту неделю пар нет: расписание ещё не заполнено или идёт практика.';
      if (gw && !/T/.test(gw.days)) why = 'На этой неделе у группы ' + weekTypeText(gw.days) + ' по графику учебного процесса, пар в расписании нет.';
      showEmpty(sel.mode === 'teacher' && net.bulk ? 'Загружаю расписание групп: ' + net.bulk.done + ' из ' + net.bulk.total + '…' : why,
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
  // ТЗ, этап 2: «импорт из таблицы и учебного плана по нажатию кнопки» (без расписания по таймеру)
  function refreshAll() {
    if (net.refreshing) return;
    var admin = isAdmin();
    net.refreshing = { done: 0, total: 0 };
    net.failed = {};
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
  function countFor(group, teacher) {
    var sheet = DB.sheets[group];
    if (!sheet) return null;
    var rows = sheet.rows.filter(function (r) { return !r.ghost && r.disc && (!teacher || r.teacher === teacher); });
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
        }).join('') + '</ol><p class="pad hint">Здесь группы, чьё расписание открыто на этом компьютере. Чтобы проверить все, нажмите «Обновить и импортировать».</p></section>';
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
    var head = '<section class="tablebox"><div class="table-head"><h2>Изменения в таблице колледжа</h2><span class="count">' +
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
    if (ex.half !== 'all' && halfOf(it.date) !== ex.half) return false;
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
      case 'disc': return '<td class="title" data-label="Дисциплина">' + discHTML(r.discipline) + (it.retake ? ' (Пересдача)' : '') + '</td>';
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
    if (EX.status === 'loading') { emptyEl.innerHTML = '<p>Загружаю расписание экзаменов из Google-таблицы…</p>'; emptyEl.hidden = false; return; }
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
      emptyEl.innerHTML = '<p>' + (EX.records.length ? 'По этим фильтрам ничего не нашлось.' : 'На листе «' + esc(CFG.examSheet) + '» пока нет ни одной записи.') + '</p>' +
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
  function acadStart() { var t = new Date(); return ymd(t.getMonth() >= 6 ? t.getFullYear() : t.getFullYear() - 1, 8, 1); }
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
    var exams = [], credits = [];
    plan.subjects.forEach(function (sj) {
      var f = sj.form[half - 1];
      if (!f) return;
      var item = { s: sj, f: f, hours: half === 1 ? sj.s1 : sj.s2, rec: findExam(g, sj.name, half) };
      if (f === 'Э') exams.push(item); else if (f === 'ДЗ' || f === 'З' || f === 'ЗаО') credits.push(item);
    });
    function page(title, list) {
      return '<section class="book-page"><h3>' + esc(title) + '</h3>' + (list.length ? '<table class="book-table"><thead><tr><th scope="col">№</th><th scope="col">Дисциплина</th><th scope="col" class="n">Часов</th>' +
        '<th scope="col">Дата</th><th scope="col">Преподаватель</th></tr></thead><tbody>' + list.map(function (it, i) {
          return '<tr><td class="n">' + (i + 1) + '</td><td>' + discHTML(it.s.name) + (it.f !== 'Э' ? ' <abbr class="hint" title="' + esc(Hours.FORM_NAMES[it.f] || '') + '">' + esc(it.f) + '</abbr>' : '') + '</td>' +
            '<td class="n">' + (it.hours || '—') + '</td><td>' + examWhen(it.rec) + '</td><td>' + esc(it.s.teacher || '—') + '</td></tr>';
        }).join('') + '</tbody></table>' : '<p class="hint">В этом семестре по плану нет.</p>') + '</section>';
    }
    box.innerHTML = controls + '<div class="book" aria-label="Разворот зачётной книжки">' +
      '<div class="book-cap"><strong>Зачётная книжка</strong><span>Группа <span translate="no">' + esc(g) + '</span>' + (plan.title ? ', ' + esc(plan.title) : '') + ', ' + half + ' семестр ' + esc(GRAFIK.data ? GRAFIK.data.title.replace(/^.*?(\d{4}\/\d{4}).*$/, '$1') : '') + '</span></div>' +
      page('Экзамены', exams) + page('Зачёты и дифференцированные зачёты', credits) + '</div>' +
      '<p class="hint">Список и часы — по учебному плану группы (лист «' + esc(CFG.planSheet) + '»), даты — по расписанию экзаменов. Контрольные работы в зачётку не вносятся.</p>';
  }
  function renderTeacherExams() {
    var box = $('ex-teacher'), t = state.exTeacher || state.teacher;
    var names = uniq(DB.teachers.concat(EX.records.map(function (r) { return splitTeachers(r.teacher)[0] || ''; }))).sort(cmpRu);
    var controls = '<section class="filters"><div class="filter-row"><div class="field f-entity"><label for="ext-teacher">Преподаватель</label><select id="ext-teacher" name="ext-teacher">' +
      '<option value="">Выберите преподавателя</option>' + names.map(function (n) { return '<option value="' + esc(n) + '"' + (n === t ? ' selected' : '') + '>' + esc(n) + '</option>'; }).join('') + '</select></div></div></section>';
    if (!t) { box.innerHTML = controls + '<p class="empty">Выберите преподавателя, чтобы увидеть его экзамены и зачёты.</p>'; return; }
    var surname = t.split(/\s+/)[0].toLowerCase().replace(/ё/g, 'е'), items = [], used = [];
    function hasT(str) { return String(str || '').toLowerCase().replace(/ё/g, 'е').indexOf(surname) >= 0; }
    PLAN.order.forEach(function (key) {
      var p = PLAN.groups[key], g = DB.groups.indexOf(key) >= 0 ? key : resolveKey(key, DB.groups, null) || key;
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
          '<td data-label="Форма">' + formHTML(it.form) + '</td><td data-label="Дата">' + examWhen(it.rec) + '</td></tr>';
      }).join('') + '</tbody></table></div></section>' +
      '<p class="hint">Из учебных планов групп (где преподаватель указан у предмета) и расписания экзаменов.</p>';
  }

  // ---------- exams loaded from a document (ТЗ, этап 8: «экзамены в документе, импорт») ----------
  var MAMMOTH_URL = 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.8.0/mammoth.browser.min.js';
  var mammothPromise = null, xi = { records: [], skipped: 0, source: '' };
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
  // the last group code in the text, the one nearest to the table
  var GROUP_IN_TEXT = /([А-ЯЁ][А-ЯЁа-яё]*\d?)\s*-\s*(\d{2})(?!\d)/g;
  function groupFromText(text) {
    var m, last = '';
    GROUP_IN_TEXT.lastIndex = 0;
    while ((m = GROUP_IN_TEXT.exec(text || ''))) last = m[1] + '-' + m[2];
    return last;
  }
  function nodeText(td) {
    var ps = td.querySelectorAll('p');
    return ps.length ? Array.prototype.map.call(ps, function (q) { return q.textContent; }).join('\n') : td.textContent;
  }
  // every table of an HTML document, with the text right above it (it often names the group)
  function tablesFromHtml(html) {
    var doc = new DOMParser().parseFromString(html, 'text/html');
    return Array.prototype.map.call(doc.querySelectorAll('table'), function (t) {
      var ctx = '', el = t.previousElementSibling, n = 0;
      while (el && n < 4 && el.tagName !== 'TABLE') { ctx = el.textContent + ' ' + ctx; el = el.previousElementSibling; n++; }
      return { rows: Array.prototype.map.call(t.querySelectorAll('tr'), function (tr) { return Array.prototype.map.call(tr.children, nodeText); }), context: ctx };
    });
  }
  function examsFromTables(tables) {
    var records = [], skipped = 0;
    tables.forEach(function (t) {
      var res = parseExamSheet(t.rows, { group: groupFromText(t.context) });
      if (res) { records = records.concat(res.records); skipped += res.skipped; }
    });
    return { records: records, skipped: skipped };
  }
  function readExamFile(file) {
    var name = file.name || '', ext = name.toLowerCase().split('.').pop();
    return file.arrayBuffer().then(function (buf) {
      if (ext === 'docx') {
        return loadMammoth().then(function (m) { return m.convertToHtml({ arrayBuffer: buf }); }).then(function (r) { return examsFromTables(tablesFromHtml(r.value)); });
      }
      if (ext === 'html' || ext === 'htm') return examsFromTables(tablesFromHtml(new TextDecoder().decode(buf)));
      // SheetJS reads a CSV as single-byte text, so Cyrillic in UTF-8 goes through the page's own parser
      if (ext === 'csv') return examsFromTables([{ rows: parseCSV(new TextDecoder().decode(buf).replace(/^﻿/, '')), context: name }]);
      return loadXlsx().then(function (X) {
        var wb = X.read(buf, { type: 'array' });
        return examsFromTables(wb.SheetNames.map(function (sn) {
          var rows = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '' }).map(function (r) { return r.map(function (v) { return v == null ? '' : String(v); }); });
          return { rows: rows, context: sn };
        }));
      });
    });
  }
  function openExamImport() {
    xi = { records: [], skipped: 0, source: '' };
    $('xi-file').value = '';
    $('xi-url').value = '';
    $('xi-status').textContent = '';
    $('xi-preview').innerHTML = '';
    $('xi-save').disabled = true;
    $('xi-dialog').showModal();
  }
  function showExamPreview(res, source) {
    xi = { records: res.records, skipped: res.skipped, source: source };
    $('xi-save').disabled = !res.records.length;
    $('xi-save').textContent = res.records.length ? 'Заменить список экзаменов (' + res.records.length + ')' : 'Заменить список экзаменов';
    if (!res.records.length) {
      $('xi-status').textContent = 'В документе не нашлось таблицы с экзаменами. Нужны колонки: Группа (или группа в заголовке над таблицей), Дисциплина, Дата, Тип аттестации — как в шаблоне.';
      $('xi-preview').innerHTML = '';
      return;
    }
    var groups = uniq(res.records.map(function (r) { return r.group; })).sort(cmpRu);
    $('xi-status').textContent = 'Прочитано ' + count(res.records.length, 'запись', 'записи', 'записей') + ' для ' + count(groups.length, 'группы', 'групп', 'групп') + (res.skipped ? ', не прочитано строк: ' + res.skipped : '') + '.';
    $('xi-preview').innerHTML = '<div class="table-wrap"><table class="smart"><thead><tr><th scope="col">Группа</th><th scope="col">Дисциплина</th><th scope="col">Дата</th><th scope="col">Тип</th></tr></thead><tbody>' +
      res.records.slice(0, 8).map(function (r) {
        return '<tr><td data-label="Группа"><span class="grp" translate="no">' + esc(r.group) + '</span></td><td class="title" data-label="Дисциплина">' + discHTML(r.discipline) + '</td>' +
          '<td data-label="Дата">' + esc(fmtShort(r.date)) + (r.time ? ', ' + esc(r.time) : '') + '</td><td data-label="Тип">' + esc(r.type) + '</td></tr>';
      }).join('') + '</tbody></table></div>' + (res.records.length > 8 ? '<p class="hint">И ещё ' + (res.records.length - 8) + '.</p>' : '');
  }
  function saveExamImport() {
    var btn = $('xi-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    Api.examsImport(xi.records, xi.source).then(function (r) {
      HIST.status = 'idle';
      $('xi-dialog').close();
      return loadEdits().then(function () { toast('Список экзаменов заменён: ' + count(r.count, 'запись', 'записи', 'записей') + '. Прошлый список можно вернуть кнопкой «Откатить» в журнале правок.'); });
    }, function (err) { $('xi-status').textContent = err.message; btn.disabled = false; btn.textContent = 'Заменить список экзаменов'; });
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
    $('foot-admin').textContent = u ? 'Меню администратора' : 'Вход для администратора';
    btn.classList.toggle('is-in', !!u);
    btn.setAttribute('aria-label', u ? 'Администратор ' + u.name + ': меню' : 'Войти как администратор');
    document.documentElement.classList.toggle('is-admin', !!u);
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
    $(id).innerHTML = uniq(values).sort(cmpRu).slice(0, 400).map(function (v) { return '<option value="' + esc(v) + '"></option>'; }).join('');
  }
  function openEditor(o) {
    var groups = o.groups.filter(Boolean);
    if (!groups.length) return;
    ed = { day: o.day, start: o.start, groups: groups, group: groups[0], all: groups.length > 1, kind: o.kind || '', editId: o.editId || '' };
    $('ed-error').textContent = '';
    var gsel = $('ed-group');
    gsel.innerHTML = (groups.length > 1 ? '<option value="*">Все группы потока: ' + esc(groups.join(', ')) + '</option>' : '') +
      groups.map(function (g) { return '<option value="' + esc(g) + '">' + esc(g) + '</option>'; }).join('');
    gsel.value = groups.length > 1 && !o.editId ? '*' : ed.group;
    $('ed-group-field').hidden = groups.length < 2;
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
    datalistFill('dl-disc', discs);
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
    if (!edit && ed.editId) edit = EDITS.list.filter(function (e) { return e.id === ed.editId; })[0] || null;
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
    $('ed-teacher').value = src.teacher || '';
    $('ed-room').value = src.room || '';
    $('ed-sub').value = String(edit ? +edit.sub || 0 : 0);
    $('ed-note').value = edit ? edit.note || '' : '';
    $('ed-remove').hidden = !edit;
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
    syncRoving();
  }
  function edSave() {
    var targets = $('ed-group').value === '*' ? ed.groups : [ed.group];
    var sub = +$('ed-sub').value || 0, note = $('ed-note').value.trim();
    var fields = { disc: $('ed-disc').value.trim(), teacher: $('ed-teacher').value.trim(), room: $('ed-room').value.trim() };
    if (ed.kind !== 'cancel' && !fields.disc) { $('ed-error').textContent = 'Укажите дисциплину.'; $('ed-disc').focus(); return; }
    var list = [], link = 'l' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
    var mineAdd = ed.edit && ed.edit.kind === 'add' ? ed.edit.id : '';
    targets.forEach(function (g) {
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
    var btn = $('ed-save');
    btn.disabled = true;
    btn.textContent = 'Сохраняю…';
    $('ed-error').textContent = '';
    var batch = targets.length > 1 ? 'b' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6) : '';
    Api.save(list, { batch: batch }).then(function (saved) {
      mergeEdits(saved);
      HIST.status = 'idle'; if (state.view === 'history') loadHistory(true);
      $('edit-dialog').close();
      recompute();
      render();
      toast(ed.kind === 'cancel' ? 'Пара отменена.' : ed.kind === 'move' ? 'Пара перенесена.' : 'Правка сохранена.');
    }, function (err) {
      $('ed-error').textContent = err.message;
      if (!Api.user()) { renderAccount(); }
    }).then(function () { btn.disabled = false; edSetKind(ed.kind); });
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
  function currentTheme() {
    var t = null;
    try { t = localStorage.getItem(THEME_KEY); } catch (e) { t = null; }
    return THEMES.some(function (x) { return x.id === t; }) ? t : 'portal';
  }
  function applyTheme(id) {
    document.documentElement.setAttribute('data-theme', id);
    try { localStorage.setItem(THEME_KEY, id); } catch (e) { /* storage unavailable */ }
    var meta = document.querySelector('meta[name="theme-color"]:not([media])');
    if (meta) meta.setAttribute('content', getComputedStyle(document.documentElement).getPropertyValue('--topbar').trim() || '#22577a');
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
      if (EX.status === 'loading') { st = 'loading'; text = 'Загружаю расписание экзаменов из Google-таблицы…'; }
      else if (EX.status === 'error') { st = 'error'; text = 'Не удалось загрузить расписание экзаменов из Google-таблицы.'; }
      else if (EX.source === 'cache') { st = 'stale'; text = 'Google-таблица не ответила. Показана сохранённая копия, загруженная ' + fmtStamp(EX.at) + '.'; }
      else {
        st = 'ok';
        var docs = EXF.records.length ? 'графики экзаменов и зачётной недели' + (EXF.period ? ', ' + EXF.period : '') : '';
        if (EXS.status === 'ok') text = 'Свежие данные из Google-таблицы, загружены ' + fmtStamp(EX.at) + (docs ? '; к ним добавлены ' + docs : '') + '.';
        else if (docs) text = 'Показаны ' + docs + '.';
        else text = 'Google-таблица прочитана ' + fmtStamp(EX.at) + ': листа с экзаменами в ней пока нет.';
      }
    } else {
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
    var gb = $('guards');
    gb.hidden = !(isAdmin() && guards.length);
    if (!gb.hidden) {
      gb.innerHTML = '<p><strong>Предупреждение для администратора: парсер не справился</strong></p><ul>' + guards.map(function (n) { return '<li>' + esc(net.guards[n]) + '</li>'; }).join('') + '</ul>' +
        '<p class="hint">Данные этих листов не затёрты. При «Обновить и импортировать» предупреждение попадёт в журнал «Изменения в таблице колледжа».</p>';
    }
  }

  function render() {
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
    if (who) renderWho();
    if (v === 'lessons') renderLessons();
    else if (v === 'hours') renderHours();
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
    var key = [state.view, state.mode, state.group, state.teacher, state.week, state.span, state.day, state.half,
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
      openEditor({ day: aday, start: astart, groups: [sel.name], kind: 'add' });
    }
  });
  $('history-body').addEventListener('click', function (e) {
    var o = e.target.closest('[data-open-edit]'), r = e.target.closest('[data-remove-edit]');
    var ke = e.target.closest('[data-keep-edit]'), de = e.target.closest('[data-drop-edit]');
    var rb = e.target.closest('[data-remove-batch]'), rv = e.target.closest('[data-revert]');
    if (e.target.closest('[data-history-retry]')) { loadHistory(true); return; }
    if (ke) { keepEdit(ke.getAttribute('data-keep-edit')); return; }
    if (de) { dropEdit(de.getAttribute('data-drop-edit')); return; }
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
      var apply = function () { if (!done) { done = true; applyTheme(pick); } };
      // the old and new themes crossfade; if the browser cannot draw the crossfade soon, the theme simply switches
      if (document.startViewTransition && !REDUCED.matches) { document.startViewTransition(apply); setTimeout(apply, 400); }
      else apply();
    }
  });
  $('ed-kind').addEventListener('click', function (e) {
    var b = e.target.closest('button');
    if (b) edSetKind(b.getAttribute('data-value'));
  });
  $('ed-group').addEventListener('change', edLoadSlot);
  $('ed-slot').addEventListener('change', edLoadSlot);
  $('ed-save').addEventListener('click', edSave);
  $('ed-cancel').addEventListener('click', function () { $('edit-dialog').close(); });
  $('ed-remove').addEventListener('click', function () {
    if (ed && ed.edit && confirm('Убрать правку? Пара снова пойдёт по основному расписанию.')) removeEdit(ed.edit.id, function () { $('edit-dialog').close(); });
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
  $('xi-file').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    if (!f) return;
    $('xi-status').textContent = 'Читаю «' + f.name + '»…';
    readExamFile(f).then(function (res) { showExamPreview(res, f.name); }, function (err) { $('xi-status').textContent = 'Не удалось прочитать файл: ' + err.message; });
  });
  $('xi-read').addEventListener('click', function () {
    var u = $('xi-url').value.trim();
    if (!u) { $('xi-status').textContent = 'Вставьте ссылку на Google-документ.'; $('xi-url').focus(); return; }
    $('xi-status').textContent = 'Открываю документ…';
    Api.fetchDoc(u).then(function (html) { showExamPreview(examsFromTables(tablesFromHtml(html)), 'Google-документ'); }, function (err) { $('xi-status').textContent = err.message; });
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
    // the sliding plates follow every change of choice, a dialog opening, the fonts arriving and the window size
    if (window.MutationObserver) new MutationObserver(queueThumbs).observe(document.body, { subtree: true, attributes: true, attributeFilter: ['aria-checked', 'aria-selected', 'hidden', 'open'] });
    window.addEventListener('resize', queueThumbs);
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(queueThumbs);
    if (wantLogin) openAdmin();
    // when the site is opened from this computer, expose the data for checks in the console
    if (/^(localhost|127\.0\.0\.1)$/.test(location.hostname)) {
      window.__raspisanie = { DB: DB, net: net, ensureAll: ensureAll, get EX() { return EX; }, applyExams: applyExams, get EDITS() { return EDITS; }, get PLAN() { return PLAN; }, get GRAFIK() { return GRAFIK; }, countFor: countFor, planFor: planFor, half2Of: half2Of, get GOV() { return GOV; }, get TJ() { return TJ; }, refreshAll: refreshAll, diffRows: diffRows, examsFromTables: examsFromTables, tablesFromHtml: tablesFromHtml, readExamFile: readExamFile, showExamPreview: showExamPreview, openExamImport: openExamImport };
    }
  })();
})();
