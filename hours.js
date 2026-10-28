// Учебная нагрузка: плановые часы с листа «0 Учебные графики» и подсчёт проведённых пар.
// Модуль без DOM: сайт берёт его как window.RaspHours, проверки в Node — через require().
(function (root) {
  'use strict';

  function cellText(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
  function low(s) { return cellText(s).toLowerCase().replace(/ё/g, 'е'); }

  // ---------- reading the plan sheet ----------
  // "3/46" → 46 hours in the term (3 a week); "2,8/ 48", "1/18+4 проект", "12/128 4сам" → the number after
  // the slash; a lone "144" is the hours themselves; "-" or an empty cell is no hours.
  function termHours(v) {
    var s = cellText(v).replace(/,/g, '.');
    if (!s || /^[-–—]+$/.test(s)) return 0;
    var m = /(\d+(?:\.\d+)?)\s*\/\s*(\d+)/.exec(s);
    if (m) return +m[2];
    m = /^(\d+)(?!\s*[.,]\d)/.exec(s);
    return m && !/нед/.test(s) ? +m[1] : 0;
  }
  // "78+18" → 78 classroom hours and 18 of consultations
  function yearHours(v) {
    var m = /(\d+)(?:\s*\+\s*(\d+))?/.exec(cellText(v));
    return m ? { hours: +m[1], cons: m[2] ? +m[2] : 0 } : { hours: 0, cons: 0 };
  }
  function weeksOf(v) { var m = /(\d+(?:[.,]\d+)?)\s*нед/.exec(cellText(v)); return m ? +m[1].replace(',', '.') : 0; }

  // "-/Э 6к/6э" → ['', 'Э']; "З /Э 12К/6Э" → ['З', 'Э']; a single mark belongs to the term that has hours
  function parseForm(v, s1, s2) {
    // «к/р» is one mark (контрольная работа), not two terms split by the slash
    var s = cellText(v).replace(/\*/g, '').replace(/к\s*[./]\s*р\.?/gi, 'КР');
    if (!s || /^[-–—]$/.test(s)) return ['', ''];
    var m = /^([А-ЯЁа-яёA-Za-z.\s-]*?)\s*\/\s*([А-ЯЁа-яёA-Za-z.-]*)/.exec(s);
    var a, b;
    if (m) { a = m[1]; b = m[2]; }
    else {
      var one = (/^[А-ЯЁа-яёA-Za-z.]+/.exec(s) || [''])[0];
      if (s2 && !s1) { a = ''; b = one; } else { a = one; b = ''; }
    }
    return [normForm(a), normForm(b)];
  }
  function normForm(f) {
    var s = low(f).replace(/[\s.-]/g, '');
    if (!s) return '';
    if (s === 'э' || s === 'экз') return 'Э';
    if (s === 'дз' || s === 'дифзач') return 'ДЗ';
    if (s === 'зао') return 'ЗаО';
    if (s === 'з' || s === 'зач') return 'З';
    if (s === 'кр' || s === 'к/р') return 'К/Р';
    if (s === 'др') return 'ДР';
    if (/проект/.test(s)) return 'Проект';
    return cellText(f).toUpperCase();
  }
  var FORM_NAMES = { 'Э': 'экзамен', 'ДЗ': 'дифференцированный зачёт', 'ЗаО': 'зачёт с оценкой', 'З': 'зачёт', 'К/Р': 'контрольная работа', 'ДР': 'другая форма', 'Проект': 'защита проекта' };

  function mapHeader(row) {
    var map = {}, prac = [];
    row.forEach(function (v, i) {
      var h = low(v);
      if (!h) return;
      if (/наименован/.test(h)) { if (map.name == null) map.name = i; return; }
      if (/форм/.test(h)) { if (map.form == null) map.form = i; return; }
      if (/препод|фио/.test(h)) { if (map.teacher == null) map.teacher = i; return; }
      if ((/практик/.test(h) || /^пр\.?\s*\d/.test(h)) && !/недел.*сем|сем.*недел/.test(h)) { prac.push(i); return; }
      if (/в год/.test(h)) { if (map.year == null) map.year = i; return; }
      // «в 1 сем.», «во 2 сем.», «Кол.ч в Нед. в 1 с.»
      if (/(^|\D)2\s*с(ем)?(\.|\s|$)/.test(h)) { if (map.s2 == null) map.s2 = i; return; }
      if (/(^|\D)1\s*с(ем)?(\.|\s|$)/.test(h)) { if (map.s1 == null) map.s1 = i; return; }
      if (/^№|п\/п/.test(h)) { if (map.num == null) map.num = i; }
    });
    // a practice column before the 2nd-term column belongs to the 1st term
    prac.forEach(function (i) {
      if (map.s2 != null && i < map.s2) { if (map.p1 == null) map.p1 = i; }
      else if (map.p2 == null) map.p2 = i;
    });
    return map;
  }

  var GROUP_LINE = /Групп[аы]\s+([А-ЯЁA-Z][А-ЯЁа-яёA-Za-z0-9]*)\s*[-–]\s*(\d{2}[а-яё]?)/;
  // Rows come straight from the CSV of the sheet. Each block: «Учебный график», the profession,
  // «Группа ГД - 26 Курс 1», the term line, a header row, a row with weeks per term, subjects.
  function parsePlan(rows) {
    var groups = {}, order = [], cur = null, map = null, title = '';
    rows.forEach(function (raw) {
      var row = (raw || []).map(cellText);
      var first = row.filter(Boolean)[0] || '';
      if (/^учебный график$/i.test(first)) { cur = null; map = null; title = ''; return; }
      if (!cur && /^(професси|специальност)/i.test(first)) { title = first; return; }
      if (!cur && title && !GROUP_LINE.test(first) && row.filter(Boolean).length === 1 && !/^семестр/i.test(first)) { title += ' ' + first; return; }
      var gm = GROUP_LINE.exec(first);
      if (gm) {
        var code = gm[1] + '-' + gm[2], cm = /курс\s*(\d)/i.exec(first);
        cur = { group: code, course: cm ? +cm[1] : 0, title: cleanTitle(title), weeks: [0, 0], subjects: [] };
        if (!groups[code]) order.push(code);
        groups[code] = cur;
        map = null;
        return;
      }
      if (!cur) return;
      if (!map) {
        var mp = mapHeader(row);
        if (mp.name != null && (mp.s1 != null || mp.s2 != null || mp.year != null)) map = mp;
        return;
      }
      var name = row[map.name] || '';
      if (!name) {
        if (map.s1 != null && weeksOf(row[map.s1])) cur.weeks[0] = weeksOf(row[map.s1]);
        if (map.s2 != null && weeksOf(row[map.s2])) cur.weeks[1] = weeksOf(row[map.s2]);
        return;
      }
      if (/наименован/i.test(name)) return;
      // a module row («ПМ 01 …») only carries the module exam, its hours are in the МДК rows below
      if (/^\s*ПМ[а-яё]?[\s.]*\d/.test(name)) return;
      var s1 = map.s1 != null ? termHours(row[map.s1]) : 0;
      var s2 = map.s2 != null ? termHours(row[map.s2]) : 0;
      var p1 = map.p1 != null ? termHours(row[map.p1]) : 0;
      var p2 = map.p2 != null ? termHours(row[map.p2]) : 0;
      var y = map.year != null ? yearHours(row[map.year]) : { hours: 0, cons: 0 };
      if (!s1 && !s2 && !p1 && !p2 && !y.hours) return; // a module heading such as «ПМ.05 …»
      var formRaw = map.form != null ? row[map.form] : '';
      cur.subjects.push({
        name: name,
        code: codeOf(name).code,
        year: y.hours || (s1 + s2 + p1 + p2),
        cons: y.cons,
        s1: s1 + p1,
        s2: s2 + p2,
        form: parseForm(formRaw, s1 + p1, s2 + p2),
        formRaw: formRaw,
        teacher: map.teacher != null ? row[map.teacher] : ''
      });
    });
    return { groups: groups, order: order };
  }
  function cleanTitle(t) {
    return cellText(t).replace(/^(Профессия|Специальность)\s*/i, '').replace(/[«»"]/g, '').replace(/\s*\(ТОП-50\)/i, '').trim();
  }

  // ---------- matching timetable names to plan names ----------
  // Codes such as МДК.05.01, УП. 02, МДКн 02.01, УПд.12.01, ОП.03 → "МДК-05-01"
  function codeOf(name) {
    var m = /^\s*([А-ЯЁA-Z]{2,5})[а-яё]?\s*\.?\s*(\d{1,2})(?:\s*\.\s*(\d{1,2}))?(?:\s*\.\s*(\d{1,2}))?\.?(?=\s|$|[^\d])/.exec(name);
    if (!m) {
      // «Учебная практика ПМ.03»: the code of the module stands at the end
      var t = /\s(ПМ|МДК|УП|ПП)[а-яё]?\s*\.?\s*(\d{1,2})(?:\s*\.\s*(\d{1,2}))?\s*$/.exec(name);
      if (!t) return { code: '', rest: name };
      var kind = /^учеб/i.test(name) ? 'УП' : /^производ/i.test(name) ? 'ПП' : t[1];
      return { code: [kind, pad2(t[2])].concat(t[3] ? [pad2(t[3])] : []).join('-'), rest: name.slice(0, t.index) };
    }
    var parts = [m[1]];
    [m[2], m[3], m[4]].forEach(function (p) { if (p) parts.push(pad2(p)); });
    return { code: parts.join('-'), rest: name.slice(m[0].length) };
  }
  function pad2(p) { return p.length < 2 ? '0' + p : p; }
  var STOP = { 'и': 1, 'в': 1, 'с': 1, 'по': 1, 'для': 1, 'на': 1, 'из': 1, 'к': 1, 'о': 1, 'об': 0, 'со': 1 };
  var SYN = [[/^специальн/, 'професс'], [/^професс/, 'професс']];
  function words(s) {
    return low(s).replace(/\([^)]*\)/g, ' ').replace(/\d\s*п\s*\/?\s*гр\.?/g, ' ').replace(/подгр\w*/g, ' ')
      .replace(/[«»"'.,:;!?()\-–—*\/+]/g, ' ').split(/\s+/).filter(Boolean);
  }
  function tokens(s) {
    return words(s).filter(function (w) { return !STOP[w] && w.length >= 3 && !/^\d+$/.test(w); }).map(function (w) {
      for (var i = 0; i < SYN.length; i++) if (SYN[i][0].test(w)) return SYN[i][1];
      return w;
    });
  }
  function acronym(s) {
    return words(s).filter(function (w) { return !STOP[w] && !/^\d+$/.test(w); }).map(function (w) { return w[0]; }).join('');
  }
  // "матем" = "математическая" (a short form), "технологии" = "технология" (only the ending differs),
  // but "информатика" ≠ "информационные"
  function sameWord(a, b) {
    var s = a.length <= b.length ? a : b, l = s === a ? b : a;
    if (s.length >= 3 && l.indexOf(s) === 0) return true;
    var k = 0;
    while (k < s.length && s[k] === l[k]) k++;
    if (k >= Math.max(5, s.length - 2)) return true;
    // a typo in a long word: «налогооблажение» = «налогообложение»
    return s.length >= 8 && l.length - s.length <= 2 && editDistance(s, l) <= 2;
  }
  function editDistance(a, b) {
    var prev = [], cur, i, j;
    for (j = 0; j <= b.length; j++) prev[j] = j;
    for (i = 1; i <= a.length; i++) {
      cur = [i];
      for (j = 1; j <= b.length; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return prev[b.length];
  }
  // 100 same code · 90 same words · 75 an abbreviation (ОПД = Основы проектной деятельности) · 64..80 word overlap
  function matchScore(planName, ttName) {
    var a = codeOf(planName), b = codeOf(ttName);
    if (a.code && b.code) {
      if (a.code === b.code) return 100;
      if (a.code.indexOf(b.code + '-') === 0 || b.code.indexOf(a.code + '-') === 0) return 95;
      return 0;
    }
    var ra = a.code ? a.rest : planName, rb = b.code ? b.rest : ttName;
    var wa = words(ra).join(' '), wb = words(rb).join(' ');
    var cap = a.code || b.code ? 70 : 100;
    if (wa && wa === wb) return Math.min(cap, 90);
    var shortA = cellText(ra), shortB = cellText(rb);
    if (/^[А-ЯЁ]{2,6}$/.test(shortA) && acronym(rb) === low(shortA)) return Math.min(cap, 75);
    if (/^[А-ЯЁ]{2,6}$/.test(shortB) && acronym(ra) === low(shortB)) return Math.min(cap, 75);
    var ta = tokens(ra), tb = tokens(rb);
    if (!ta.length || !tb.length) return 0;
    var used = {}, hit = 0;
    ta.forEach(function (x) {
      for (var j = 0; j < tb.length; j++) if (!used[j] && sameWord(x, tb[j])) { used[j] = true; hit++; return; }
    });
    var dice = 2 * hit / (ta.length + tb.length);
    return dice >= 0.6 ? Math.min(cap, Math.round(40 + dice * 40)) : 0;
  }
  // timetable discipline → index of the plan subject, for one group
  function linkSubjects(subjects, discNames) {
    var out = {};
    discNames.forEach(function (d) {
      var best = -1, bestScore = 59;
      subjects.forEach(function (s, i) {
        var sc = matchScore(s.name, d);
        if (sc > bestScore) { best = i; bestScore = sc; }
      });
      out[d] = best;
    });
    return out;
  }

  // ---------- the academic year from the PDF schedule ----------
  // First Monday of the 2nd half: the week after the winter holidays in the schedule of the group.
  function secondHalfStart(grafikGroup, weeks) {
    if (grafikGroup && weeks) {
      var days = grafikGroup.days, inK = false;
      for (var i = 0; i < days.length; i++) {
        var isK = days[i].indexOf('K') >= 0 && /^K+$/.test(days[i].replace(/T/g, '')) && weeks[i].slice(5, 7) !== '06';
        var m = +weeks[i].slice(5, 7);
        if (isK && (m === 12 || m === 1)) inK = true;
        else if (inK) return weeks[i];
      }
    }
    return null;
  }

  // ---------- counting ----------
  // rows: the lessons of one group after the admin edits; cancelled lessons are already out.
  // Every pair is the same number of academic hours; one pair given to both subgroups counts once.
  function countHours(rows, opt) {
    var per = opt.hoursPerPair || 2, nowIso = opt.today, nowMin = opt.nowMin;
    var res = {};
    var seen = {};
    rows.forEach(function (r) {
      var key = r.day + '|' + r.start + '|' + r.disc;
      if (seen[key]) return;
      seen[key] = true;
      var half = r.day < opt.half2 ? 1 : 2;
      var e = res[r.disc] || (res[r.disc] = { disc: r.disc, h: [null, { done: 0, all: 0, last: '', next: '' }, { done: 0, all: 0, last: '', next: '' }], teachers: {} });
      var x = e.h[half];
      x.all += per;
      var past = r.day < nowIso || (r.day === nowIso && toMin(r.end) <= nowMin);
      if (past) { x.done += per; if (r.day > x.last) x.last = r.day; }
      else if (!x.next || r.day < x.next) x.next = r.day;
      if (r.teacher) e.teachers[r.teacher] = true;
    });
    return res;
  }
  function toMin(t) { var p = String(t).split(':'); return (+p[0]) * 60 + (+p[1] || 0); }

  // Plan subjects with hours from the timetable, plus timetable lessons that the plan does not list.
  function summarize(plan, counted, half) {
    var subjects = plan ? plan.subjects : [];
    var link = linkSubjects(subjects, Object.keys(counted));
    var lines = subjects.map(function (s) {
      return { subject: s, names: [], plan: half === 0 ? s.s1 + s.s2 : half === 1 ? s.s1 : s.s2, done: 0, all: 0, last: '', next: '', teachers: {} };
    });
    var extra = [];
    Object.keys(counted).forEach(function (d) {
      var c = counted[d], i = link[d];
      var halves = half === 0 ? [1, 2] : [half];
      var target = i >= 0 ? lines[i] : null;
      if (!target) {
        target = { subject: null, names: [], plan: null, done: 0, all: 0, last: '', next: '', teachers: {} };
        extra.push(target);
      }
      target.names.push(d);
      halves.forEach(function (h) {
        var x = c.h[h];
        target.done += x.done; target.all += x.all;
        if (x.last > target.last) target.last = x.last;
        if (x.next && (!target.next || x.next < target.next)) target.next = x.next;
      });
      Object.keys(c.teachers).forEach(function (t) { target.teachers[t] = true; });
    });
    lines.forEach(function (l) { l.left = Math.max(0, l.plan - l.done); });
    extra = extra.filter(function (l) { return l.all > 0; });
    return { lines: lines.filter(function (l) { return l.plan > 0 || l.all > 0; }), extra: extra };
  }

  var api = {
    parsePlan: parsePlan, termHours: termHours, yearHours: yearHours, parseForm: parseForm, FORM_NAMES: FORM_NAMES,
    codeOf: codeOf, matchScore: matchScore, linkSubjects: linkSubjects, countHours: countHours, summarize: summarize,
    secondHalfStart: secondHalfStart
  };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.RaspHours = api;
})(typeof window !== 'undefined' ? window : this);
