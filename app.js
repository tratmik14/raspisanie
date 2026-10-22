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
  var WEEKDAYS_MON = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота'];
  var WEEKDAY_NAMES = ['понедельник', 'вторник', 'среда', 'четверг', 'пятница', 'суббота', 'воскресенье'];
  var TEACHER_RE = /^[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.\s?(?:[А-ЯЁ]\.?)?$/;
  var GROUP_RE = /^[А-ЯЁA-Z][А-ЯЁA-Zа-яё0-9]*-\d{2}[А-ЯЁа-яё]?$/;
  var TYPE_KEY = { 'Экзамен': 'exam', 'Диф. зачёт': 'diff', 'Зачёт': 'credit' };
  var EX_DEFAULTS = { group: '', teacher: '', disc: '', type: 'all', status: 'all', from: '', to: '', search: '', ordering: '', page: 1, pageSize: 20, rpage: 1 };

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
  function fmtDayMonth(iso) { var d = isoToDate(iso); return d.getDate() + ' ' + MONTHS_GEN[d.getMonth()]; }
  function fmtShort(iso) { var p = isoParts(iso); return pad(p.d) + '.' + pad(p.m) + '.' + p.y; }
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
  function parseExamSheet(rows) {
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
      var discipline = cellText(get(row, 'discipline')), group = cellText(get(row, 'group'));
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
  // exam list: status 'loading' | 'ok' | 'none' (no such sheet) | 'error'
  var EX = { status: 'loading', records: [], skipped: 0, source: '', at: 0 };
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
  function applyExams(csv, source, at) {
    var res = parseExamSheet(parseCSV(csv));
    if (!res) { EX = { status: 'none', records: [], skipped: 0, source: source, at: at }; return; }
    EX = { status: 'ok', records: res.records, skipped: res.skipped, source: source, at: at };
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

  // ---------- state and the address bar (ТЗ БП 1.5: filters live in query parameters) ----------
  var state = { view: 'lessons', mode: 'group', group: '', teacher: '', week: '' };
  var ex = Object.assign({}, EX_DEFAULTS);
  var pendingHash = ''; // an old #group link that waits for the group list
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
  function readUrl() {
    var p = new URLSearchParams(location.search);
    if (p.get('view') === 'exams') {
      state.view = 'exams';
      ex = Object.assign({}, EX_DEFAULTS, {
        group: p.get('group') || '', teacher: p.get('teacher') || '', disc: p.get('discipline') || '',
        type: /^(exam|credit|diff|retake)$/.test(p.get('type')) ? p.get('type') : 'all',
        status: /^(future|past)$/.test(p.get('status')) ? p.get('status') : 'all',
        from: p.get('date_from') || '', to: p.get('date_to') || '', search: p.get('search') || '',
        ordering: p.get('ordering') || '', page: Math.max(1, +p.get('page') || 1), rpage: Math.max(1, +p.get('rpage') || 1),
        pageSize: [10, 20, 50, 100].indexOf(+p.get('page_size')) >= 0 ? +p.get('page_size') : 20
      });
    } else {
      state.view = 'lessons';
      if (p.get('teacher')) { state.mode = 'teacher'; state.teacher = p.get('teacher'); }
      else if (p.get('group')) { state.mode = 'group'; state.group = p.get('group'); }
      if (/^\d{4}-\d{2}-\d{2}$/.test(p.get('week') || '')) state.week = p.get('week');
    }
    var h = '';
    try { h = decodeURIComponent(location.hash.replace(/^#/, '')).toLowerCase(); } catch (e) { h = ''; }
    if (h) pendingHash = h;
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
      if (ex.group) p.set('group', ex.group);
      if (ex.teacher) p.set('teacher', ex.teacher);
      if (ex.disc) p.set('discipline', ex.disc);
      if (ex.type !== 'all') p.set('type', ex.type);
      if (ex.status !== 'all') p.set('status', ex.status);
      if (ex.from) p.set('date_from', ex.from);
      if (ex.to) p.set('date_to', ex.to);
      if (ex.search) p.set('search', ex.search);
      if (ex.ordering) p.set('ordering', ex.ordering);
      if (ex.page > 1) p.set('page', ex.page);
      if (ex.rpage > 1) p.set('rpage', ex.rpage);
      if (ex.pageSize !== 20) p.set('page_size', ex.pageSize);
    } else {
      if (state.mode === 'teacher' && state.teacher) p.set('teacher', state.teacher);
      else if (state.mode === 'group' && state.group) p.set('group', state.group);
      if (state.week && state.week !== currentWeekIso()) p.set('week', state.week);
    }
    var qs = p.toString(), target = location.pathname + (qs ? '?' + qs : '');
    if (target === location.pathname + location.search && !location.hash) return;
    try { history.replaceState(null, '', target); } catch (e) { /* ignore */ }
  }

  // ---------- loading ----------
  function isLive(group) { return DB.sheets[group] && DB.sheets[group].source === 'live'; }
  function allLive() { return DB.groups.length > 0 && DB.groups.every(isLive); }
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
    });
  }
  function ensureView() {
    if (state.view !== 'lessons') return;
    if (state.mode === 'teacher') ensureAll();
    else if (state.group) ensureGroup(state.group);
  }
  function loadExams() {
    if (!CFG.examSheet) { EX = { status: 'none', records: [], skipped: 0, source: '', at: 0 }; return Promise.resolve(); }
    return fetchSheet(CFG.examSheet).then(function (csv) {
      applyExams(csv, 'live', Date.now());
    }, function () {
      var c = cacheGet(CFG.examSheet);
      if (c) applyExams(c.csv, 'cache', c.t);
      else EX = { status: 'error', records: [], skipped: 0, source: '', at: 0 };
    });
  }
  function loadLists() {
    net.pending++;
    net.listFailed = false;
    render();
    var cal = CFG.calendarSheet ? fetchSheet(CFG.calendarSheet).catch(function () { return null; }) : Promise.resolve(null);
    var lists = Promise.all([fetchSheet(CFG.groupListSheet), cal]).then(function (res) {
      applyGroupList(res[0]);
      if (res[1]) applyCalendar(res[1]);
      net.listLive = true;
    }, function () {
      net.listFailed = true;
    });
    return Promise.all([lists, loadExams()]).then(function () {
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
    if (!net.listLive || EX.status === 'error') loadLists(); else ensureView();
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
        return '<button type="button" class="chip" data-group="' + esc(g) + '">' + esc(g) + '</button>';
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

  // ТЗ БП 5.8: a reminder about exams and credits within the next 3 calendar days
  function renderReminder(sel, today) {
    var el = $('reminder');
    if (EX.status !== 'ok' || !sel) { el.hidden = true; return; }
    var soon = examItems().filter(function (it) {
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
    $('reminder').hidden = true;
    $('picker').hidden = true;
    weekEl.innerHTML = '';
    var sheet = sel && sel.mode === 'group' ? DB.sheets[sel.name] : null;
    $('tpl-note').hidden = !(sheet && sheet.template);

    var weeks = DB.weeks, wi = nearestWeekIndex(state.week || currentWeekIso());
    $('weeknav').hidden = !(sel && wi >= 0);

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

    state.week = weeks[wi].iso;
    var mondayMs = weeks[wi].ms;
    $('wk-label').textContent = fmtRange(weeks[wi].iso, isoOfUtc(mondayMs + 5 * DAY));
    $('wk-sub').textContent = 'Неделя ' + (Math.round((mondayMs - weeks[0].ms) / (7 * DAY)) + 1) +
      (weeks[wi].par === 1 ? ' · чётная' : weeks[wi].par === 2 ? ' · нечётная' : '');
    $('wk-prev').disabled = wi <= 0;
    $('wk-next').disabled = wi >= weeks.length - 1;
    $('wk-today').hidden = nearestWeekIndex(currentWeekIso()) === wi;

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
  function matchesExam(it, today) {
    var r = it.rec;
    if (ex.group && r.group !== ex.group) return false;
    if (ex.teacher && splitTeachers(r.teacher).indexOf(ex.teacher) < 0) return false;
    if (ex.disc && r.discipline !== ex.disc) return false;
    if (ex.type !== 'all' && ex.type !== 'retake' && TYPE_KEY[r.type] !== ex.type) return false;
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
      case 'group': return '<td data-label="Группа"><span class="grp">' + esc(r.group) + '</span></td>';
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
    var cols = EX_COLUMNS.filter(function (c) { return !(c.key === 'group' && ex.group); });
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
  function fillExamOptions() {
    var recs = EX.records;
    var groups = uniq(recs.map(function (r) { return r.group; })).sort(cmpRu);
    var teachers = uniq([].concat.apply([], recs.map(function (r) { return splitTeachers(r.teacher); }))).sort(cmpRu);
    var discs = uniq(recs.filter(function (r) { return !ex.group || r.group === ex.group; }).map(function (r) { return r.discipline; })).sort(cmpRu);
    $('e-group').innerHTML = optionsHTML('Все группы', groups, ex.group);
    $('e-teacher').innerHTML = optionsHTML('Все преподаватели', teachers, ex.teacher);
    $('e-disc').innerHTML = optionsHTML('Все дисциплины', discs, ex.disc);
  }
  function renderExams() {
    var today = todayISO();
    $('e-group').value = ex.group;
    $('e-teacher').value = ex.teacher;
    $('e-disc').value = ex.disc;
    $('e-from').value = ex.from;
    $('e-to').value = ex.to;
    if (document.activeElement !== $('e-search')) $('e-search').value = ex.search;
    [['e-type', ex.type], ['e-status', ex.status]].forEach(function (pair) {
      Array.prototype.forEach.call($(pair[0]).querySelectorAll('button'), function (b) {
        b.setAttribute('aria-checked', b.getAttribute('data-value') === pair[1] ? 'true' : 'false');
      });
    });
    var dirty = ['group', 'teacher', 'disc', 'type', 'status', 'from', 'to', 'search'].some(function (k) { return ex[k] !== EX_DEFAULTS[k]; });
    $('e-reset').hidden = !dirty;

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
        '<li>Дата — «23.06.2027 10:00» или «23 июня 2027 г 10 часов 00 минут». Тип — Экзамен, Зачёт или Диф. зачёт.</li></ol>';
      emptyEl.hidden = false;
      return;
    }
    var all = examItems().filter(function (it) { return matchesExam(it, today); });
    var main = ex.type === 'retake' ? [] : sortItems(all.filter(function (it) { return !it.retake; }));
    var ret = sortItems(all.filter(function (it) { return it.retake; }));
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
      X.writeFile(wb, 'ekzameny' + (ex.group ? '-' + slug(ex.group) : '') + '.xlsx');
    }).catch(function () {
      alert('Не удалось загрузить модуль Excel. Проверьте подключение к интернету.');
    }).then(function () { btn.disabled = false; });
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
      else { st = 'ok'; text = EX.status === 'none' ? 'Google-таблица прочитана ' + fmtStamp(EX.at) + ': листа с экзаменами в ней пока нет.' : 'Свежие данные из Google-таблицы, загружены ' + fmtStamp(EX.at) + '.'; }
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
    el.setAttribute('data-state', st);
    $('status-text').textContent = text;
    $('retry').hidden = !(st === 'stale' || st === 'error');
  }

  function render() {
    resolvePendingHash();
    var td = new Date();
    $('today-label').textContent = WEEKDAYS[td.getDay()][0].toUpperCase() + WEEKDAYS[td.getDay()].slice(1) + ', ' + td.getDate() + ' ' + MONTHS_GEN[td.getMonth()];
    $('term-label').textContent = termLabel();
    $('year').textContent = td.getFullYear();

    var isExams = state.view === 'exams';
    $('view-lessons').hidden = isExams;
    $('view-exams').hidden = !isExams;
    $('tab-lessons').setAttribute('aria-selected', isExams ? 'false' : 'true');
    $('tab-exams').setAttribute('aria-selected', isExams ? 'true' : 'false');
    $('page-title').textContent = isExams ? 'Расписание экзаменов' : 'Расписание занятий';
    $('page-lede').textContent = isExams
      ? 'Экзамены, зачёты и пересдачи по группам и преподавателям.'
      : 'Пары на каждую неделю семестра: для группы или для преподавателя.';
    document.title = (isExams ? 'Расписание экзаменов' : 'Расписание занятий') + ' · ТехноПортал';
    $('stamp').textContent = isExams
      ? (EX.status === 'ok' ? count(uniq(EX.records.map(function (r) { return r.group; })).length, 'группа', 'группы', 'групп') + ' · ' + count(EX.records.length, 'запись', 'записи', 'записей') : '')
      : (DB.groups.length ? count(DB.groups.length, 'группа', 'группы', 'групп') + (DB.weeks.length ? ' · ' + count(DB.weeks.length, 'неделя', 'недели', 'недель') + ' с ' + fmtDayMonth(DB.weeks[0].iso) : '') : '');
    renderStatus();
    if (isExams) renderExams(); else renderLessons();
    savePrefs();
    writeUrl();
  }

  // ---------- events ----------
  function switchView(view) {
    if (view === state.view) return;
    if (view === 'exams' && !ex.group && !ex.teacher) {
      if (state.mode === 'group' && state.group) ex.group = state.group;
      if (state.mode === 'teacher' && state.teacher) ex.teacher = state.teacher;
      fillExamOptions();
    }
    if (view === 'lessons') {
      if (ex.group && !state.group && DB.groups.indexOf(ex.group) >= 0) { state.mode = 'group'; state.group = ex.group; }
      else if (ex.teacher && !state.teacher) { state.mode = 'teacher'; state.teacher = ex.teacher; }
    }
    state.view = view;
    render();
    ensureView();
  }
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
  $('reminder').addEventListener('click', function (e) {
    if (e.target.id !== 'reminder-link') return;
    e.preventDefault();
    ex = Object.assign({}, EX_DEFAULTS, { status: 'future' });
    if (state.mode === 'group') ex.group = state.group; else ex.teacher = state.teacher;
    state.view = 'exams';
    fillExamOptions();
    render();
  });
  $('retry').addEventListener('click', retry);

  function setEx(key, value) {
    ex[key] = value;
    ex.page = 1; ex.rpage = 1;
    if (key === 'group') { ex.disc = ''; fillExamOptions(); }
    render();
  }
  $('e-group').addEventListener('change', function (e) { setEx('group', e.target.value); });
  $('e-teacher').addEventListener('change', function (e) { setEx('teacher', e.target.value); });
  $('e-disc').addEventListener('change', function (e) { setEx('disc', e.target.value); });
  $('e-from').addEventListener('change', function (e) { setEx('from', e.target.value); });
  $('e-to').addEventListener('change', function (e) { setEx('to', e.target.value); });
  $('e-search').addEventListener('input', function (e) { setEx('search', e.target.value); });
  ['e-type', 'e-status'].forEach(function (id) {
    $(id).addEventListener('click', function (e) {
      var b = e.target.closest('button');
      if (b) setEx(id === 'e-type' ? 'type' : 'status', b.getAttribute('data-value'));
    });
  });
  function resetExFilters() {
    var keep = { ordering: ex.ordering, pageSize: ex.pageSize };
    ex = Object.assign({}, EX_DEFAULTS, keep);
    fillExamOptions();
    render();
  }
  $('e-reset').addEventListener('click', resetExFilters);
  $('e-export').addEventListener('click', exportExams);
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
    readUrl();
    var lc = cacheGet(CFG.groupListSheet), cc = CFG.calendarSheet ? cacheGet(CFG.calendarSheet) : null, ec = CFG.examSheet ? cacheGet(CFG.examSheet) : null;
    if (lc) applyGroupList(lc.csv);
    if (cc) applyCalendar(cc.csv);
    if (ec) applyExams(ec.csv, 'cache', ec.t);
    if (state.view === 'lessons' && state.mode === 'group' && state.group) {
      var c = cacheGet(state.group);
      if (c) applySheet(state.group, c.csv, 'cache', c.t);
    }
    recompute();
    fillOptions();
    render();
    loadLists();
    // keep "today" and the current pair fresh while the page stays open
    setInterval(function () { if (state.view === 'lessons') render(); }, 60000);
    // when the site is opened from this computer, expose the data for checks in the console
    if (/^(localhost|127\.0\.0\.1)$/.test(location.hostname)) window.__raspisanie = { DB: DB, net: net, ensureAll: ensureAll, get EX() { return EX; }, applyExams: applyExams };
  })();
})();
