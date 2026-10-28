/**
 * Расписание · сервер правок и истории изменений (Google Apps Script).
 *
 * Скрипт живёт в отдельной Google-таблице администраторов (не в таблице расписания колледжа):
 * Расширения → Apps Script → вставить этот файл → Развернуть → Веб-приложение
 * (Выполнять от имени: «Я», Доступ: «Все»). Ссылку на веб-приложение — в config.js, поле adminApi.
 *
 * Листы создаются сами при первом запуске:
 *   «Правки»  — действующие и отменённые изменения расписания;
 *   «История» — журнал: кто, когда, что было и что стало;
 *   «Админы»  — логины, роли и хэши паролей (сами пароли нигде не хранятся).
 * Первого администратора добавьте через меню таблицы «Расписание → Добавить администратора».
 */

var VERSION = '1.0';
var SHEETS = { edits: 'Правки', history: 'История', admins: 'Админы' };
var EDIT_COLS = ['id', 'created', 'updated', 'author', 'status', 'group', 'date', 'start', 'end', 'kind', 'disc', 'teacher', 'room', 'sub', 'note', 'was', 'link', 'type'];
var EDIT_HEAD = ['ID', 'Создано', 'Изменено', 'Автор', 'Статус', 'Группа', 'Дата', 'Начало', 'Конец', 'Вид', 'Дисциплина', 'Преподаватель', 'Аудитория', 'Подгруппа', 'Комментарий', 'Было (JSON)', 'Связь', 'Тип записи'];
var HISTORY_COLS = ['id', 'time', 'login', 'name', 'action', 'editId', 'group', 'date', 'start', 'before', 'after', 'comment', 'type'];
var HISTORY_HEAD = ['ID', 'Время', 'Логин', 'Имя', 'Действие', 'ID правки', 'Группа', 'Дата', 'Начало', 'Было (JSON)', 'Стало (JSON)', 'Комментарий', 'Тип записи'];
var ADMIN_COLS = ['login', 'name', 'role', 'salt', 'hash', 'created', 'lastLogin', 'active'];
var ADMIN_HEAD = ['Логин', 'Имя', 'Роль', 'Соль', 'Хэш пароля', 'Создан', 'Последний вход', 'Активен'];
var KINDS = { replace: 'Замена', cancel: 'Отмена', add: 'Добавлена' };
var SESSION_TTL = 21600;   // 6 hours, the longest CacheService keeps a value
var MAX_FAILS = 5;         // wrong passwords in a row before the login is locked
var LOCK_SECONDS = 900;    // 15 minutes
var HASH_ROUNDS = 200;

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
      case 'edits': return listEdits(req);
      case 'history': return history(req);
      case 'login': return login(req);
      case 'logout': return logout(req);
      case 'me': return { ok: true, user: needSession(req) };
      case 'save': return save(req);
      case 'remove': return remove(req);
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
function writeRow(name, cols, head, obj, row) {
  var sh = sheet(name, cols, head), vals = [cols.map(function (c) { return obj[c] == null ? '' : String(obj[c]); })];
  if (row) sh.getRange(row, 1, 1, cols.length).setValues(vals);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length).setValues(vals);
}
function readEdits() {
  return readRows(SHEETS.edits, EDIT_COLS, EDIT_HEAD).map(function (e) { e.sub = +e.sub || 0; return e; });
}
function readAdmins() { return readRows(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD); }
function isActive(a) { return a.login && String(a.active).toLowerCase() !== 'нет' && String(a.active).toLowerCase() !== 'false'; }
function nowIso() { return new Date().toISOString(); }
function newId(prefix) { return prefix + Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

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
function changePassword(req) {
  var u = needSession(req), pass = String(req.next || '');
  if (pass.length < 8) return fail('Новый пароль — не короче 8 символов');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var admin = readAdmins().filter(function (a) { return a.login === u.login; })[0];
    if (!admin || !sameText(hashPassword(String(req.old || ''), admin.salt), admin.hash)) return fail('Текущий пароль неверный');
    admin.salt = Utilities.getUuid();
    admin.hash = hashPassword(pass, admin.salt);
    writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, admin, admin._row);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// ---------- edits ----------
function clean(e) {
  var s = function (v, max) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max); };
  var out = {
    id: s(e.id, 40), group: s(e.group, 40), date: s(e.date, 10), start: s(e.start, 5), end: s(e.end, 5),
    kind: s(e.kind, 10), disc: s(e.disc, 300), teacher: s(e.teacher, 120), room: s(e.room, 60),
    sub: Math.max(0, Math.min(3, parseInt(e.sub, 10) || 0)), note: s(e.note, 300),
    was: String(e.was == null ? '' : typeof e.was === 'string' ? e.was : JSON.stringify(e.was)).slice(0, 3000), link: s(e.link, 40)
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
  return e ? JSON.stringify({ kind: e.kind, disc: e.disc, teacher: e.teacher, room: e.room, sub: e.sub, note: e.note, status: e.status, link: e.link || '' }) : '';
}
// Тип записи: «замена» — пару заменили другой; «ручная правка» — пару добавили, отменили или перенесли.
// Записи основного расписания («основная») живут в таблице колледжа и сюда не попадают.
function recordType(e) { return e.kind === 'replace' && !e.link ? 'замена' : 'ручная правка'; }
function logChange(user, action, edit, before, after, comment) {
  writeRow(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, {
    id: newId('h'), time: nowIso(), login: user.login, name: user.name, action: action, editId: edit.id,
    group: edit.group, date: edit.date, start: edit.start, before: before, after: after, comment: comment || '', type: recordType(edit)
  });
}
// Everyone gets the active edits; who made each one only goes to signed-in admins.
function listEdits(req) {
  var u = session(req), names = {};
  if (u) readAdmins().forEach(function (a) { names[a.login] = a.name || a.login; });
  var edits = readEdits().filter(function (e) { return e.status === 'active'; }).map(function (e) {
    var o = publicEdit(e);
    if (u) o.authorName = names[o.author] || o.author;
    else delete o.author;
    return o;
  });
  return { ok: true, edits: edits, at: Date.now(), full: !!u };
}
// One active edit per slot: saving onto a slot that already has one updates it.
function save(req) {
  var user = needSession(req);
  var list = Array.isArray(req.edits) ? req.edits : req.edit ? [req.edit] : [];
  if (!list.length || list.length > 20) return fail('Передайте от 1 до 20 правок');
  var items = list.map(clean);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var saved = [];
  try {
    var all = readEdits();
    var link = items.length > 1 ? (items[0].link || newId('l')) : items[0].link;
    items.forEach(function (it) {
      if (items.length > 1) it.link = link;
      var cur = it.id ? all.filter(function (e) { return e.id === it.id; })[0] : null;
      if (!cur) cur = all.filter(function (e) { return e.status === 'active' && slotKey(e) === slotKey(it); })[0] || null;
      var now = nowIso();
      if (cur) {
        var before = snapshot(cur);
        if (cur.status !== 'active') before = '';
        var row = cur._row;
        var next = publicEdit(cur);
        ['group', 'date', 'start', 'end', 'kind', 'disc', 'teacher', 'room', 'sub', 'note', 'link'].forEach(function (k) { next[k] = it[k]; });
        if (it.was) next.was = it.was;
        next.status = 'active'; next.updated = now; next.author = user.login; next.type = recordType(next);
        writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, next, row);
        logChange(user, before ? 'update' : 'create', next, before, snapshot(next), req.comment);
        saved.push(next);
      } else {
        var e = publicEdit(it);
        e.id = newId('e'); e.created = now; e.updated = now; e.author = user.login; e.status = 'active'; e.type = recordType(e);
        writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, e);
        all.push(e);
        logChange(user, 'create', e, '', snapshot(e), req.comment);
        saved.push(e);
      }
    });
  } finally { lock.releaseLock(); }
  return { ok: true, edits: saved.map(function (e) { e = publicEdit(e); e.authorName = user.name; return e; }) };
}
// Removing an edit brings back the lesson from the main timetable; the row stays as «deleted».
function remove(req) {
  var user = needSession(req);
  var ids = Array.isArray(req.ids) ? req.ids.map(String) : req.id ? [String(req.id)] : [];
  if (!ids.length) return fail('Не указаны правки');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var removed = [];
  try {
    var all = readEdits();
    // removing one half of a move removes the other half too
    var links = {};
    all.forEach(function (e) { if (ids.indexOf(e.id) >= 0 && e.link) links[e.link] = true; });
    all.forEach(function (e) {
      if (e.status !== 'active' || (ids.indexOf(e.id) < 0 && !(e.link && links[e.link]))) return;
      var before = snapshot(e);
      e.status = 'deleted'; e.updated = nowIso(); e.author = user.login;
      writeRow(SHEETS.edits, EDIT_COLS, EDIT_HEAD, publicEdit(e), e._row);
      logChange(user, 'delete', e, before, '', req.comment);
      removed.push(e.id);
    });
  } finally { lock.releaseLock(); }
  return { ok: true, removed: removed };
}

// ---------- history ----------
// Everyone sees what changed; who changed it is shown only to signed-in admins.
function history(req) {
  var u = session(req), limit = Math.max(1, Math.min(1000, parseInt(req.limit, 10) || 300));
  var group = String(req.group || '');
  var items = readRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD)
    .filter(function (h) { return !group || h.group === group; })
    .slice(-limit).reverse()
    .map(function (h) {
      var o = { id: h.id, time: h.time, action: h.action, editId: h.editId, group: h.group, date: h.date, start: h.start, before: h.before, after: h.after, comment: h.comment, type: h.type };
      if (u) { o.login = h.login; o.name = h.name; }
      return o;
    });
  return { ok: true, items: items, full: !!u };
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
  sheet(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD);
}
function ask(ui, title, text) {
  var r = ui.prompt(title, text, ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return null;
  return r.getResponseText().trim();
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
  var role = readAdmins().filter(isActive).length ? 'редактор' : 'главный';
  var salt = Utilities.getUuid();
  writeRow(SHEETS.admins, ADMIN_COLS, ADMIN_HEAD, { login: login, name: name || login, role: role, salt: salt, hash: hashPassword(pass, salt), created: nowIso(), lastLogin: '', active: 'да' });
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
