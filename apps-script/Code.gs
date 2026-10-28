/**
 * Расписание · сервер правок и истории изменений (Google Apps Script).
 *
 * Скрипт живёт в отдельной Google-таблице администраторов (не в таблице расписания колледжа):
 * Расширения → Apps Script → вставить этот файл → Развернуть → Веб-приложение
 * (Выполнять от имени: «Я», Доступ: «Все»). Ссылку на веб-приложение — в config.js, поле adminApi.
 *
 * Листы создаются сами при первом запуске:
 *   «Правки»              — действующие и отменённые изменения расписания;
 *   «История»             — журнал правок: кто, когда, что было и что стало;
 *   «График»              — правки графика учебного процесса (сдвинутые практики и т. п.);
 *   «Снимки»              — последняя импортированная версия каждого листа расписания;
 *   «Изменения таблицы»   — что изменилось в таблице колледжа между импортами, предупреждения парсера;
 *   «Экзамены»            — экзамены и зачёты, загруженные из документа;
 *   «Админы»              — логины, роли и хэши паролей (сами пароли нигде не хранятся).
 * Первого администратора добавьте через меню таблицы «Расписание → Добавить администратора».
 * Тот же файл сайт запускает у себя в демо-режиме (gas-sim.js), поэтому правила одни и те же.
 */

var VERSION = '2.0';
var SHEETS = {
  edits: 'Правки', history: 'История', admins: 'Админы', grafik: 'График', snaps: 'Снимки',
  journal: 'Изменения таблицы', exams: 'Экзамены', examsPrev: 'Экзамены (прошлая версия)'
};
var EDIT_COLS = ['id', 'created', 'updated', 'author', 'status', 'group', 'date', 'start', 'end', 'kind', 'disc', 'teacher', 'room', 'sub', 'note', 'was', 'link', 'type', 'batch'];
var EDIT_HEAD = ['ID', 'Создано', 'Изменено', 'Автор', 'Статус', 'Группа', 'Дата', 'Начало', 'Конец', 'Вид', 'Дисциплина', 'Преподаватель', 'Аудитория', 'Подгруппа', 'Комментарий', 'Было (JSON)', 'Связь', 'Тип записи', 'Массовая замена'];
var HISTORY_COLS = ['id', 'time', 'login', 'name', 'action', 'editId', 'group', 'date', 'start', 'before', 'after', 'comment', 'type', 'target', 'batch'];
var HISTORY_HEAD = ['ID', 'Время', 'Логин', 'Имя', 'Действие', 'ID записи', 'Группа', 'Дата', 'Начало', 'Было (JSON)', 'Стало (JSON)', 'Комментарий', 'Тип записи', 'Что меняли', 'Массовая замена'];
var ADMIN_COLS = ['login', 'name', 'role', 'salt', 'hash', 'created', 'lastLogin', 'active'];
var ADMIN_HEAD = ['Логин', 'Имя', 'Роль', 'Соль', 'Хэш пароля', 'Создан', 'Последний вход', 'Активен'];
var GRAFIK_COLS = ['id', 'created', 'updated', 'author', 'status', 'group', 'week', 'days', 'note'];
var GRAFIK_HEAD = ['ID', 'Создано', 'Изменено', 'Автор', 'Статус', 'Группа', 'Неделя (понедельник)', 'Дни пн–пт', 'Комментарий'];
var SNAP_COLS = ['group', 'part', 'at', 'hash', 'data'];
var SNAP_HEAD = ['Группа', 'Часть', 'Импортировано', 'Контрольная сумма', 'Данные'];
var JOURNAL_COLS = ['id', 'time', 'login', 'name', 'kind', 'group', 'date', 'start', 'sub', 'before', 'after', 'note'];
var JOURNAL_HEAD = ['ID', 'Время', 'Логин', 'Имя', 'Вид', 'Группа', 'Дата', 'Начало', 'Подгруппа', 'Было', 'Стало', 'Пояснение'];
var EXAM_COLS = ['id', 'group', 'discipline', 'date', 'time', 'room', 'teacher', 'type', 'retakeDate', 'retakeTime', 'source', 'imported', 'author'];
var EXAM_HEAD = ['ID', 'Группа', 'Дисциплина', 'Дата', 'Время', 'Аудитория', 'Преподаватель', 'Тип аттестации', 'Дата пересдачи', 'Время пересдачи', 'Источник', 'Загружено', 'Кто загрузил'];
var KINDS = { replace: 'Замена', cancel: 'Отмена', add: 'Добавлена' };
var WEEK_CODES = /^[TAKUPXDG.]{5}$/;   // the 5 weekdays of the PDF schedule, «.» keeps the day as in the PDF
var SESSION_TTL = 21600;   // 6 hours, the longest CacheService keeps a value
var MAX_FAILS = 5;         // wrong passwords in a row before the login is locked
var LOCK_SECONDS = 900;    // 15 minutes
var HASH_ROUNDS = 200;
var MAX_ITEMS = 200;       // edits or schedule weeks in one request (a teacher's month of pairs fits)
var CELL_LIMIT = 45000;    // a Sheets cell holds 50 000 characters; snapshots are split below that

// ---------- entry points ----------
function doGet(e) { return respond(route((e && e.parameter) || {})); }
function doPost(e) {
  var body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { body = {}; }
  return respond(route(body));
}
function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function route(req) {
  try {
    switch (String(req.action || 'ping')) {
      case 'ping': return { ok: true, version: VERSION, admins: readAdmins().filter(isActive).length };
      case 'state': return state(req);
      case 'edits': return listEdits(req);
      case 'history': return history(req);
      case 'login': return login(req);
      case 'logout': return logout(req);
      case 'me': return { ok: true, user: needSession(req) };
      case 'save': return save(req);
      case 'remove': return remove(req);
      case 'revert': return revert(req);
      case 'grafikSave': return grafikSave(req);
      case 'snapshots': return snapshots(req);
      case 'importTable': return importTable(req);
      case 'tableJournal': return tableJournal(req);
      case 'examsImport': return examsImport(req);
      case 'fetchDoc': return fetchDoc(req);
      case 'password': return changePassword(req);
      default: return fail('Неизвестное действие');
    }
  } catch (err) {
    return fail(err && err.userMessage ? err.userMessage : 'Ошибка сервера: ' + (err && err.message ? err.message : err));
  }
}
function fail(message, extra) {
  var out = { ok: false, error: message };
  if (extra) for (var k in extra) out[k] = extra[k];
  return out;
}
function userError(message) { var e = new Error(message); e.userMessage = message; return e; }

// ---------- sheets ----------
function book() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet(name, cols, head) {
  var ss = book(), sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // plain text keeps "2026-10-24" and "8:30" from turning into dates and times
    sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  }
  return sh;
}
function readRows(name, cols, head) {
  var sh = sheet(name, cols, head), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, cols.length).getValues().map(function (r, i) {
    var o = { _row: i + 2 };
    cols.forEach(function (c, j) { o[c] = asText(r[j]); });
    return o;
  });
}
function asText(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function rowValues(cols, obj) { return cols.map(function (c) { return obj[c] == null ? '' : String(obj[c]); }); }
// a new sheet has 1000 rows; past them getRange fails, so add rows (as plain text) in advance
function ensureRows(sh, cols, last) {
  if (last <= sh.getMaxRows()) return;
  var from = sh.getMaxRows(), add = Math.max(500, last - from);
  sh.insertRowsAfter(from, add);
  sh.getRange(from + 1, 1, add, cols.length).setNumberFormat('@');
}
function writeRow(name, cols, head, obj, row) {
  if (!row) { appendRows(name, cols, head, [obj]); return; }
  sheet(name, cols, head).getRange(row, 1, 1, cols.length).setValues([rowValues(cols, obj)]);
}
// one write for many rows: a bulk replacement of 50 pairs stays one call
function appendRows(name, cols, head, objs) {
  if (!objs.length) return;
  var sh = sheet(name, cols, head), at = sh.getLastRow() + 1;
  ensureRows(sh, cols, at + objs.length - 1);
  sh.getRange(at, 1, objs.length, cols.length).setValues(objs.map(function (o) { return rowValues(cols, o); }));
}
function replaceRows(name, cols, head, objs) {
  var sh = sheet(name, cols, head), n = sh.getLastRow();
  if (n >= 2) sh.getRange(2, 1, n - 1, cols.length).clearContent();
  if (!objs.length) return;
  ensureRows(sh, cols, objs.length + 1);
  sh.getRange(2, 1, objs.length, cols.length).setValues(objs.map(function (o) { return rowValues(cols, o); }));
}
// Writes must reach the sheet before the lock goes, or the next request may read old rows.
function release(lock) { SpreadsheetApp.flush(); lock.releaseLock(); }
function withLock(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { release(lock); }
}
function readEdits() {
  return readRows(SHEETS.edits, EDIT_COLS, EDIT_HEAD).map(function (e) { e.sub = +e.sub || 0; return e; });
}
function readAdmins() { return readRows(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD); }
function isActive(a) { return a.login && String(a.active).toLowerCase() !== 'нет' && String(a.active).toLowerCase() !== 'false'; }
function nowIso() { return new Date().toISOString(); }
function newId(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function str(v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max); }

// ---------- passwords and sessions ----------
function hex(bytes) {
  return bytes.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length < 2 ? '0' + v : v; }).join('');
}
function hashPassword(password, salt) {
  var v = salt + ':' + password;
  for (var i = 0; i < HASH_ROUNDS; i++) v = hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, v, Utilities.Charset.UTF_8));
  return v;
}
function sameText(a, b) {
  a = String(a); b = String(b);
  var diff = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function cache() { return CacheService.getScriptCache(); }
function login(req) {
  var name = String(req.login || '').trim().toLowerCase(), password = String(req.password || '');
  if (!name || !password) return fail('Введите логин и пароль');
  var c = cache();
  if (c.get('lock:' + name)) return fail('Слишком много неудачных попыток. Попробуйте через 15 минут.', { locked: true });
  var admin = readAdmins().filter(function (a) { return isActive(a) && a.login.toLowerCase() === name; })[0];
  if (!admin || !sameText(hashPassword(password, admin.salt), admin.hash)) {
    var fails = (+c.get('fail:' + name) || 0) + 1;
    if (fails >= MAX_FAILS) { c.put('lock:' + name, '1', LOCK_SECONDS); c.remove('fail:' + name); }
    else c.put('fail:' + name, String(fails), LOCK_SECONDS);
    return fail('Неверный логин или пароль', { attemptsLeft: Math.max(0, MAX_FAILS - fails) });
  }
  c.remove('fail:' + name);
  var token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  var user = { login: admin.login, name: admin.name || admin.login, role: admin.role || 'редактор' };
  c.put('s:' + token, JSON.stringify(user), SESSION_TTL);
  admin.lastLogin = nowIso();
  writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, admin, admin._row);
  return { ok: true, token: token, user: user, expires: Date.now() + SESSION_TTL * 1000 };
}
function logout(req) { if (req.token) cache().remove('s:' + req.token); return { ok: true }; }
function session(req) {
  if (!req.token) return null;
  var v = cache().get('s:' + req.token);
  return v ? JSON.parse(v) : null;
}
function needSession(req) {
  var u = session(req);
  // a disabled admin loses access at once, not when the session runs out
  if (u && !readAdmins().some(function (a) { return isActive(a) && a.login === u.login; })) { cache().remove('s:' + req.token); u = null; }
  if (!u) throw userError('Сессия истекла. Войдите снова.');
  return u;
}
function adminNames() {
  var names = {};
  readAdmins().forEach(function (a) { names[a.login] = a.name || a.login; });
  return names;
}
function changePassword(req) {
  var u = needSession(req), pass = String(req.next || '');
  if (pass.length < 8) return fail('Новый пароль — не короче 8 символов');
  return withLock(function () {
    var admin = readAdmins().filter(function (a) { return a.login === u.login; })[0];
    if (!admin || !sameText(hashPassword(String(req.old || ''), admin.salt), admin.hash)) return fail('Текущий пароль неверный');
    admin.salt = Utilities.getUuid();
    admin.hash = hashPassword(pass, admin.salt);
    writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, admin, admin._row);
    return { ok: true };
  });
}

// ---------- history ----------
function historyEntry(user, o) {
  return {
    id: newId('h'), time: nowIso(), login: user.login, name: user.name, action: o.action, editId: o.editId || '',
    group: o.group || '', date: o.date || '', start: o.start || '', before: o.before || '', after: o.after || '',
    comment: o.comment || '', type: o.type || '', target: o.target || 'edit', batch: o.batch || ''
  };
}
// Everyone sees what changed; who changed it is shown only to signed-in admins.
function history(req) {
  var u = session(req), limit = Math.max(1, Math.min(1000, parseInt(req.limit, 10) || 300));
  var group = String(req.group || '');
  var items = readRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD)
    .filter(function (h) { return !group || h.group === group; })
    .slice(-limit).reverse()
    .map(function (h) {
      var o = {};
      HISTORY_COLS.forEach(function (c) { if (c !== 'login' && c !== 'name') o[c] = h[c]; });
      o.target = o.target || 'edit';
      if (u) { o.login = h.login; o.name = h.name; }
      return o;
    });
  return { ok: true, items: items, full: !!u };
}

// ---------- edits ----------
function clean(e) {
  var out = {
    id: str(e.id, 40), group: str(e.group, 40), date: str(e.date, 10), start: str(e.start, 5), end: str(e.end, 5),
    kind: str(e.kind, 10), disc: str(e.disc, 300), teacher: str(e.teacher, 120), room: str(e.room, 60),
    sub: Math.max(0, Math.min(3, parseInt(e.sub, 10) || 0)), note: str(e.note, 300),
    was: String(e.was == null ? '' : typeof e.was === 'string' ? e.was : JSON.stringify(e.was)).slice(0, 3000), link: str(e.link, 40)
  };
  if (!out.group) throw userError('Не указана группа');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(out.date)) throw userError('Дата должна быть в виде ГГГГ-ММ-ДД');
  if (!/^\d{1,2}:\d{2}$/.test(out.start)) throw userError('Время начала должно быть в виде Ч:ММ');
  if (out.end && !/^\d{1,2}:\d{2}$/.test(out.end)) throw userError('Время конца должно быть в виде Ч:ММ');
  if (!KINDS[out.kind]) throw userError('Вид правки: replace, cancel или add');
  if (out.kind !== 'cancel' && !out.disc) throw userError('Укажите дисциплину');
  if (out.kind === 'cancel') { out.disc = ''; out.teacher = ''; out.room = ''; }
  return out;
}
function slotKey(e) { return [e.group, e.date, e.start, e.sub || 0].join('|'); }
function publicEdit(e) {
  var o = {};
  EDIT_COLS.forEach(function (c) { o[c] = e[c]; });
  o.sub = +o.sub || 0;
  return o;
}
function snapshot(e) {
  return e ? JSON.stringify({
    kind: e.kind, disc: e.disc, teacher: e.teacher, room: e.room, sub: e.sub, note: e.note, status: e.status, link: e.link || '',
    group: e.group, date: e.date, start: e.start, end: e.end
  }) : '';
}
// Тип записи: «замена» — пару заменили другой; «ручная правка» — пару добавили, отменили или перенесли.
// Записи основного расписания («основная») живут в таблице колледжа и сюда не попадают.
function recordType(e) { return e.kind === 'replace' && !e.link ? 'замена' : 'ручная правка'; }
function editEntry(user, action, edit, before, after, comment) {
  return historyEntry(user, {
    action: action, editId: edit.id, group: edit.group, date: edit.date, start: edit.start, before: before, after: after,
    comment: comment, type: recordType(edit), target: 'edit', batch: edit.batch
  });
}
function activeEdits(u, names) {
  return readEdits().filter(function (e) { return e.status === 'active'; }).map(function (e) {
    var o = publicEdit(e);
    if (u) o.authorName = names[o.author] || o.author;
    else delete o.author;
    return o;
  });
}
// Everyone gets the active edits; who made each one only goes to signed-in admins.
function listEdits(req) {
  var u = session(req);
  return { ok: true, edits: activeEdits(u, u ? adminNames() : {}), at: Date.now(), full: !!u };
}
// Everything the site lays over the college timetable, in one request.
function state(req) {
  var u = session(req), names = u ? adminNames() : {};
  var grafik = readRows(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD).filter(function (g) { return g.status === 'active'; }).map(function (g) {
    var o = { id: g.id, group: g.group, week: g.week, days: g.days, note: g.note, updated: g.updated };
    if (u) o.authorName = names[g.author] || g.author;
    return o;
  });
  var exams = readRows(SHEETS.exams, EXAM_COLS, EXAM_HEAD).map(function (x) {
    var o = {};
    EXAM_COLS.forEach(function (c) { if (c !== 'author') o[c] = x[c]; });
    return o;
  });
  return { ok: true, edits: activeEdits(u, names), grafik: grafik, exams: exams, at: Date.now(), full: !!u };
}
// One active replacement or cancellation per slot: saving onto such a slot updates it.
// Added pairs («add») never take a slot, so two pairs can stand at the same time.
function save(req) {
  var user = needSession(req);
  var list = Array.isArray(req.edits) ? req.edits : req.edit ? [req.edit] : [];
  if (!list.length || list.length > MAX_ITEMS) return fail('Передайте от 1 до ' + MAX_ITEMS + ' правок');
  var items = list.map(clean), batch = str(req.batch, 40);
  return withLock(function () {
    var all = readEdits(), saved = [], fresh = [], log = [], now = nowIso();
    items.forEach(function (it) {
      var cur = it.id ? all.filter(function (e) { return e.id === it.id; })[0] : null;
      if (!cur && it.kind !== 'add') cur = all.filter(function (e) { return e.status === 'active' && e.kind !== 'add' && slotKey(e) === slotKey(it); })[0] || null;
      if (cur) {
        var before = cur.status === 'active' ? snapshot(cur) : '';
        var next = publicEdit(cur);
        ['group', 'date', 'start', 'end', 'kind', 'disc', 'teacher', 'room', 'sub', 'note', 'link'].forEach(function (k) { next[k] = it[k]; });
        if (it.was) next.was = it.was;
        if (batch) next.batch = batch;
        next.status = 'active'; next.updated = now; next.author = user.login; next.type = recordType(next);
        writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, next, cur._row);
        log.push(editEntry(user, before ? 'update' : 'create', next, before, snapshot(next), req.comment));
        saved.push(next);
      } else {
        var e = publicEdit(it);
        e.id = newId('e'); e.created = now; e.updated = now; e.author = user.login; e.status = 'active'; e.type = recordType(e); e.batch = batch;
        all.push(e);
        fresh.push(e);
        log.push(editEntry(user, 'create', e, '', snapshot(e), req.comment));
        saved.push(e);
      }
    });
    appendRows(SHEETS.edits, EDIT_COLS, EDIT_HEAD, fresh);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, log);
    return { ok: true, edits: saved.map(function (e) { e = publicEdit(e); e.authorName = user.name; return e; }) };
  });
}
// Removing an edit brings back the lesson from the main timetable; the row stays as «deleted».
// One half of a move takes the other half along; {batch} removes a whole bulk replacement.
function remove(req) {
  var user = needSession(req);
  var ids = Array.isArray(req.ids) ? req.ids.map(String) : req.id ? [String(req.id)] : [];
  var batch = str(req.batch, 40);
  if (!ids.length && !batch) return fail('Не указаны правки');
  return withLock(function () {
    var all = readEdits(), links = {}, removed = [], log = [];
    all.forEach(function (e) { if (ids.indexOf(e.id) >= 0 && e.link) links[e.link] = true; });
    all.forEach(function (e) {
      var hit = ids.indexOf(e.id) >= 0 || (e.link && links[e.link]) || (batch && e.batch === batch);
      if (e.status !== 'active' || !hit) return;
      var before = snapshot(e);
      e.status = 'deleted'; e.updated = nowIso(); e.author = user.login;
      writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, publicEdit(e), e._row);
      log.push(editEntry(user, 'delete', e, before, '', req.comment));
      removed.push(e.id);
    });
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, log);
    return { ok: true, removed: removed };
  });
}
// Откат: the record goes back to what it was before the chosen journal entry.
function revert(req) {
  var user = needSession(req), id = str(req.id, 40);
  return withLock(function () {
    var h = readRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD).filter(function (x) { return x.id === id; })[0];
    if (!h) return fail('Запись журнала не найдена');
    var comment = 'откат записи журнала от ' + h.time.slice(0, 16).replace('T', ' ');
    if ((h.target || 'edit') === 'grafik') return revertGrafik(user, h, comment);
    if (h.target === 'exams') return revertExams(user, h, comment);
    var e = readEdits().filter(function (x) { return x.id === h.editId; })[0];
    if (!e) return fail('Правка из этой записи не найдена');
    var was = h.before ? JSON.parse(h.before) : null, prev = e.status === 'active' ? snapshot(e) : '';
    if (!was || was.status === 'deleted') e.status = 'deleted';
    else {
      ['kind', 'disc', 'teacher', 'room', 'sub', 'note', 'link', 'group', 'date', 'start', 'end'].forEach(function (k) { if (was[k] != null) e[k] = was[k]; });
      e.status = 'active';
    }
    e.updated = nowIso(); e.author = user.login; e.type = recordType(e);
    writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, publicEdit(e), e._row);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [editEntry(user, 'revert', e, prev, e.status === 'active' ? snapshot(e) : '', comment)]);
    return { ok: true, edit: publicEdit(e) };
  });
}

// ---------- the PDF schedule: shifted practice weeks and other corrections ----------
function grafikSave(req) {
  var user = needSession(req);
  var list = Array.isArray(req.items) ? req.items : [];
  if (!list.length || list.length > MAX_ITEMS) return fail('Передайте от 1 до ' + MAX_ITEMS + ' недель');
  var items = list.map(function (it) {
    var o = { group: str(it.group, 40), week: str(it.week, 10), days: str(it.days, 5).toUpperCase(), note: str(it.note, 300) };
    if (!o.group) throw userError('Не указана группа');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(o.week)) throw userError('Неделя — дата понедельника в виде ГГГГ-ММ-ДД');
    if (o.days && !WEEK_CODES.test(o.days)) throw userError('Дни недели: 5 знаков из T A K U P X D G или «.»');
    if (o.days === '.....') o.days = '';
    return o;
  });
  return withLock(function () {
    var all = readRows(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD), fresh = [], log = [], now = nowIso(), saved = [];
    items.forEach(function (it) {
      var cur = all.filter(function (g) { return g.status === 'active' && g.group === it.group && g.week === it.week; })[0];
      var before = cur ? JSON.stringify({ days: cur.days, note: cur.note }) : '';
      if (!it.days) {
        if (!cur) return;
        cur.status = 'deleted'; cur.updated = now; cur.author = user.login;
        writeRow(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD, cur, cur._row);
        log.push(historyEntry(user, { action: 'delete', editId: cur.id, group: it.group, date: it.week, before: before, comment: req.comment, target: 'grafik' }));
        return;
      }
      var g = cur || { id: newId('g'), created: now, status: 'active', group: it.group, week: it.week };
      g.days = it.days; g.note = it.note; g.updated = now; g.author = user.login;
      if (cur) writeRow(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD, g, cur._row);
      else { fresh.push(g); all.push(g); }
      log.push(historyEntry(user, { action: cur ? 'update' : 'create', editId: g.id, group: g.group, date: g.week, before: before, after: JSON.stringify({ days: g.days, note: g.note }), comment: req.comment, target: 'grafik' }));
      saved.push({ id: g.id, group: g.group, week: g.week, days: g.days, note: g.note, updated: g.updated, authorName: user.name });
    });
    appendRows(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD, fresh);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, log);
    return { ok: true, items: saved };
  });
}
function revertGrafik(user, h, comment) {
  var all = readRows(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD);
  var g = all.filter(function (x) { return x.id === h.editId; })[0];
  if (!g) return fail('Правка графика из этой записи не найдена');
  var was = h.before ? JSON.parse(h.before) : null, prev = g.status === 'active' ? JSON.stringify({ days: g.days, note: g.note }) : '';
  // another active row for the same week would hide the restored one
  all.forEach(function (x) {
    if (x.id !== g.id && x.status === 'active' && x.group === g.group && x.week === g.week) {
      x.status = 'deleted'; writeRow(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD, x, x._row);
    }
  });
  if (was) { g.days = was.days; g.note = was.note || ''; g.status = 'active'; } else g.status = 'deleted';
  g.updated = nowIso(); g.author = user.login;
  writeRow(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD, g, g._row);
  appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, {
    action: 'revert', editId: g.id, group: g.group, date: g.week, before: prev, after: g.status === 'active' ? JSON.stringify({ days: g.days, note: g.note }) : '', comment: comment, target: 'grafik'
  })]);
  return { ok: true };
}

// ---------- import of the college timetable: snapshots and «что изменилось» ----------
// The site parses the timetable itself; on import an admin's browser sends here the parsed lessons of
// every group, the differences from the previous snapshot and the sheets the parser could not read.
function snapshots(req) {
  needSession(req);
  var out = {};
  readRows(SHEETS.snaps, SNAP_COLS, SNAP_HEAD).sort(function (a, b) { return (+a.part || 0) - (+b.part || 0); }).forEach(function (s) {
    var o = out[s.group] || (out[s.group] = { hash: s.hash, at: s.at, data: '' });
    o.data += s.data;
  });
  return { ok: true, snapshots: out };
}
function importTable(req) {
  var user = needSession(req);
  var groups = Array.isArray(req.groups) ? req.groups : [], changes = Array.isArray(req.changes) ? req.changes : [];
  var warnings = Array.isArray(req.warnings) ? req.warnings : [];
  if (groups.length > 300 || changes.length > 20000 || warnings.length > 300) return fail('Слишком большой импорт');
  return withLock(function () {
    var now = nowIso(), names = {};
    groups.forEach(function (g) { names[str(g.group, 40)] = true; });
    var keep = readRows(SHEETS.snaps, SNAP_COLS, SNAP_HEAD).filter(function (s) { return !names[s.group]; });
    groups.forEach(function (g) {
      var data = String(g.data || ''), part = 0;
      do {
        keep.push({ group: str(g.group, 40), part: part, at: now, hash: str(g.hash, 40), data: data.slice(part * CELL_LIMIT, (part + 1) * CELL_LIMIT) });
        part++;
      } while (part * CELL_LIMIT < data.length);
    });
    replaceRows(SHEETS.snaps, SNAP_COLS, SNAP_HEAD, keep);
    var rows = changes.map(function (c) {
      return {
        id: newId('t'), time: now, login: user.login, name: user.name, kind: 'change', group: str(c.group, 40), date: str(c.date, 10),
        start: str(c.start, 5), sub: parseInt(c.sub, 10) || 0, before: str(c.before, 500), after: str(c.after, 500), note: ''
      };
    }).concat(warnings.map(function (w) {
      return { id: newId('t'), time: now, login: user.login, name: user.name, kind: 'warning', group: str(w.group, 40), date: '', start: '', sub: '', before: '', after: '', note: str(w.note, 500) };
    }));
    appendRows(SHEETS.journal, JOURNAL_COLS, JOURNAL_HEAD, rows);
    return { ok: true, groups: groups.length, changes: changes.length, warnings: warnings.length, at: now };
  });
}
// Changes are public; parser warnings and who ran the import are for admins.
function tableJournal(req) {
  var u = session(req), limit = Math.max(1, Math.min(3000, parseInt(req.limit, 10) || 600)), group = String(req.group || '');
  var items = readRows(SHEETS.journal, JOURNAL_COLS, JOURNAL_HEAD)
    .filter(function (j) { return (!group || j.group === group) && (u || j.kind !== 'warning'); })
    .slice(-limit).reverse()
    .map(function (j) {
      var o = { id: j.id, time: j.time, kind: j.kind, group: j.group, date: j.date, start: j.start, sub: +j.sub || 0, before: j.before, after: j.after, note: j.note };
      if (u) o.name = j.name;
      return o;
    });
  return { ok: true, items: items, full: !!u };
}

// ---------- exams loaded from a document ----------
function examsImport(req) {
  var user = needSession(req);
  var list = Array.isArray(req.records) ? req.records : [];
  if (list.length > 3000) return fail('Слишком много строк');
  var now = nowIso(), source = str(req.source, 200);
  var records = list.map(function (r) {
    var o = {
      id: newId('x'), group: str(r.group, 40), discipline: str(r.discipline, 300), date: str(r.date, 10), time: str(r.time, 5),
      room: str(r.room, 60), teacher: str(r.teacher, 200), type: str(r.type, 30), retakeDate: str(r.retakeDate, 10), retakeTime: str(r.retakeTime, 5),
      source: source, imported: now, author: user.login
    };
    if (!o.group || !o.discipline || !/^\d{4}-\d{2}-\d{2}$/.test(o.date)) throw userError('Строка без группы, дисциплины или даты: ' + o.discipline);
    return o;
  });
  return withLock(function () {
    var old = readRows(SHEETS.exams, EXAM_COLS, EXAM_HEAD);
    replaceRows(SHEETS.examsPrev, EXAM_COLS, EXAM_HEAD, old);
    replaceRows(SHEETS.exams, EXAM_COLS, EXAM_HEAD, records);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, {
      action: 'import', before: JSON.stringify({ count: old.length }), after: JSON.stringify({ count: records.length, source: source }), comment: req.comment, target: 'exams'
    })]);
    return { ok: true, count: records.length };
  });
}
function revertExams(user, h, comment) {
  var cur = readRows(SHEETS.exams, EXAM_COLS, EXAM_HEAD), prev = readRows(SHEETS.examsPrev, EXAM_COLS, EXAM_HEAD);
  replaceRows(SHEETS.exams, EXAM_COLS, EXAM_HEAD, prev);
  replaceRows(SHEETS.examsPrev, EXAM_COLS, EXAM_HEAD, cur);
  appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, {
    action: 'revert', before: JSON.stringify({ count: cur.length }), after: JSON.stringify({ count: prev.length }), comment: comment, target: 'exams'
  })]);
  return { ok: true, count: prev.length };
}
// A Google Doc opened «for everyone with the link» comes back as HTML; the site reads its tables.
function fetchDoc(req) {
  needSession(req);
  var m = /docs\.google\.com\/document\/d\/([\w-]{20,})/.exec(String(req.url || ''));
  if (!m) return fail('Нужна ссылка на Google-документ вида docs.google.com/document/d/…');
  var r = UrlFetchApp.fetch('https://docs.google.com/document/d/' + m[1] + '/export?format=html', { muteHttpExceptions: true, followRedirects: true });
  var type = String((r.getHeaders() || {})['Content-Type'] || '');
  if (r.getResponseCode() !== 200 || type.indexOf('text/html') < 0) {
    return fail('Документ не открылся. Откройте доступ «Все, у кого есть ссылка» или скачайте его как .docx и загрузите файлом.');
  }
  return { ok: true, html: r.getContentText() };
}

// ---------- spreadsheet menu for the owner ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Расписание')
    .addItem('Добавить администратора', 'menuAddAdmin')
    .addItem('Сбросить пароль администратора', 'menuResetPassword')
    .addItem('Отключить администратора', 'menuDisableAdmin')
    .addToUi();
  sheet(SHEETS.edits, EDIT_COLS, EDIT_HEAD);
  sheet(SHEETS.history, HISTORY_COLS, HISTORY_HEAD);
  sheet(SHEETS.grafik, GRAFIK_COLS, GRAFIK_HEAD);
  sheet(SHEETS.journal, JOURNAL_COLS, JOURNAL_HEAD);
  sheet(SHEETS.exams, EXAM_COLS, EXAM_HEAD);
  sheet(SHEETS.snaps, SNAP_COLS, SNAP_HEAD);
  sheet(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD);
}
function ask(ui, title, text) {
  var r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText().trim();
}
function addAdmin(login, name, pass) {
  var role = readAdmins().filter(isActive).length ? 'редактор' : 'главный';
  var salt = Utilities.getUuid();
  writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, { login: login, name: name || login, role: role, salt: salt, hash: hashPassword(pass, salt), created: nowIso(), lastLogin: '', active: 'да' });
}
function menuAddAdmin() {
  var ui = SpreadsheetApp.getUi();
  var login = ask(ui, 'Новый администратор', 'Логин (латиницей, например ivanova):');
  if (!login) return;
  login = login.toLowerCase();
  if (!/^[a-z0-9._-]{3,32}$/.test(login)) { ui.alert('Логин: 3–32 символа, латинские буквы, цифры, точка, дефис.'); return; }
  if (readAdmins().some(function (a) { return a.login.toLowerCase() === login; })) { ui.alert('Такой логин уже есть.'); return; }
  var name = ask(ui, 'Новый администратор', 'Имя для истории изменений (например, Иванова А. С.):');
  if (name === null) return;
  var pass = ask(ui, 'Новый администратор', 'Пароль (не короче 8 символов). Его увидите только вы — передайте лично:');
  if (!pass || pass.length < 8) { ui.alert('Пароль должен быть не короче 8 символов.'); return; }
  addAdmin(login, name, pass);
  ui.alert('Готово: ' + login + ' может входить на сайт.');
}
function menuResetPassword() {
  var ui = SpreadsheetApp.getUi();
  var login = ask(ui, 'Сброс пароля', 'Логин администратора:');
  if (!login) return;
  var admin = readAdmins().filter(function (a) { return a.login.toLowerCase() === login.toLowerCase(); })[0];
  if (!admin) { ui.alert('Такого логина нет.'); return; }
  var pass = ask(ui, 'Сброс пароля', 'Новый пароль (не короче 8 символов):');
  if (!pass || pass.length < 8) { ui.alert('Пароль должен быть не короче 8 символов.'); return; }
  admin.salt = Utilities.getUuid();
  admin.hash = hashPassword(pass, admin.salt);
  admin.active = 'да';
  writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, admin, admin._row);
  cache().remove('lock:' + admin.login.toLowerCase());
  ui.alert('Пароль для ' + admin.login + ' изменён.');
}
function menuDisableAdmin() {
  var ui = SpreadsheetApp.getUi();
  var login = ask(ui, 'Отключить администратора', 'Логин:');
  if (!login) return;
  var admin = readAdmins().filter(function (a) { return a.login.toLowerCase() === login.toLowerCase(); })[0];
  if (!admin) { ui.alert('Такого логина нет.'); return; }
  admin.active = 'нет';
  writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, admin, admin._row);
  ui.alert(admin.login + ' больше не может входить и вносить правки.');
}
