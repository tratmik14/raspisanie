// Изменяемая часть учебного плана. Исходные блоки «0 Учебные графики» остаются основой.
var PLAN_COLS = ['id', 'group', 'name', 's1', 's2', 'teacher', 'f1', 'f2', 'updated', 'author', 'status'];
var PLAN_HEAD = ['ID', 'Группа', 'Дисциплина', 'Часы 1 семестр', 'Часы 2 семестр', 'Преподаватель', 'Форма 1 семестр', 'Форма 2 семестр', 'Обновлено', 'Автор', 'Статус'];
function publicPlan(r) {
  return { id: r.id, group: r.group, name: r.name, s1: +r.s1, s2: +r.s2, teacher: r.teacher, f1: r.f1, f2: r.f2, updated: r.updated };
}
function activePlans() {
  return readRows(SHEETS.plans, PLAN_COLS, PLAN_HEAD).filter(function (r) { return r.status === 'active'; }).map(publicPlan);
}
function planSave(req) {
  var user = needSession(req), input = req.record || {}, forms = ['', 'Э', 'З', 'ДЗ', 'ЗаО', 'К/Р', 'ДР', 'Проект'];
  var record = { group: str(input.group, 40), name: str(input.name, 300), s1: Number(input.s1), s2: Number(input.s2), teacher: str(input.teacher, 200), f1: str(input.f1, 20), f2: str(input.f2, 20) };
  if (!record.group || !record.name) return fail('Укажите группу и предмет.');
  if (!Number.isInteger(record.s1) || !Number.isInteger(record.s2) || record.s1 < 0 || record.s2 < 0 || record.s1 > 10000 || record.s2 > 10000) return fail('Часы каждого семестра — целое число от 0 до 10 000.');
  if (forms.indexOf(record.f1) < 0 || forms.indexOf(record.f2) < 0) return fail('Выберите форму аттестации из списка.');
  return withLock(function () {
    var all = readRows(SHEETS.plans, PLAN_COLS, PLAN_HEAD);
    var cur = all.filter(function (r) { return r.group === record.group && r.name === record.name; }).pop();
    var before = cur && cur.status === 'active' ? JSON.stringify(publicPlan(cur)) : '';
    record.id = cur ? cur.id : newId('p'); record.updated = nowIso(); record.author = user.login; record.status = 'active';
    writeRow(SHEETS.plans, PLAN_COLS, PLAN_HEAD, record, cur ? cur._row : null);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, {
      action: before ? 'update' : 'create', editId: record.id, group: record.group, before: before,
      after: JSON.stringify(publicPlan(record)), comment: 'изменён учебный план', target: 'plan'
    })]);
    return { ok: true, record: publicPlan(record) };
  });
}
function revertPlan(user, h, comment) {
  var cur = readRows(SHEETS.plans, PLAN_COLS, PLAN_HEAD).filter(function (r) { return r.id === h.editId; })[0];
  if (!cur) return fail('Правка учебного плана не найдена.');
  var before = cur.status === 'active' ? JSON.stringify(publicPlan(cur)) : '';
  if (h.before) { var prev = JSON.parse(h.before); ['group', 'name', 's1', 's2', 'teacher', 'f1', 'f2'].forEach(function (k) { cur[k] = prev[k]; }); cur.status = 'active'; }
  else cur.status = 'deleted';
  cur.updated = nowIso(); cur.author = user.login;
  writeRow(SHEETS.plans, PLAN_COLS, PLAN_HEAD, cur, cur._row);
  appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, {
    action: 'revert', editId: cur.id, group: cur.group, before: before,
    after: cur.status === 'active' ? JSON.stringify(publicPlan(cur)) : '', comment: comment, target: 'plan'
  })]);
  return { ok: true };
}

// Названия хранятся вместе с прежними вариантами: существующие пары и планы
// остаются связаны с одним предметом или преподавателем после переименования.
var CATALOG_COLS = ['id', 'type', 'name', 'aliases', 'updated', 'author', 'status'];
var CATALOG_HEAD = ['ID', 'Тип', 'Название', 'Прежние названия (JSON)', 'Обновлено', 'Автор', 'Статус'];
function catalogKey(s) { return str(s, 300).toLowerCase(); }
function publicCatalog(r) { return { id: r.id, type: r.type, name: r.name, aliases: JSON.parse(r.aliases || '[]'), updated: r.updated }; }
function activeCatalog() { return readRows(SHEETS.catalog, CATALOG_COLS, CATALOG_HEAD).filter(function (r) { return r.status === 'active'; }).map(publicCatalog); }
function catalogSave(req) {
  var user = needSession(req), input = req.record || {};
  var type = str(input.type, 20), name = str(input.name, 300), from = str(input.from, 300), id = str(input.id, 40);
  if (['teacher', 'subject'].indexOf(type) < 0 || !name) return fail('Укажите тип и название.');
  if (type === 'teacher' && /[,;]/.test(name)) return fail('Добавляйте каждого преподавателя отдельно, без запятых и точек с запятой.');
  return withLock(function () {
    var all = readRows(SHEETS.catalog, CATALOG_COLS, CATALOG_HEAD);
    var cur = all.filter(function (r) { return r.status === 'active' && r.type === type && (id ? r.id === id : from && [r.name].concat(JSON.parse(r.aliases || '[]')).some(function (n) { return catalogKey(n) === catalogKey(from); })); })[0];
    if (id && !cur) return fail('Запись справочника не найдена. Обновите страницу.');
    var aliases = cur ? JSON.parse(cur.aliases || '[]').concat(cur.name) : [];
    if (from) aliases.push(from);
    aliases = aliases.filter(function (v, i, a) { return v && catalogKey(v) !== catalogKey(name) && a.indexOf(v) === i; });
    var claims = [name].concat(aliases).map(catalogKey);
    if (all.some(function (r) { return r.status === 'active' && r.type === type && (!cur || r.id !== cur.id) && [r.name].concat(JSON.parse(r.aliases || '[]')).some(function (v) { return claims.indexOf(catalogKey(v)) >= 0; }); })) return fail('Такое название уже есть в справочнике. Выберите существующую запись.');
    if (aliases.length > 100) return fail('Слишком много прежних названий.');
    var before = cur ? JSON.stringify(publicCatalog(cur)) : '';
    var r = { id: cur ? cur.id : newId('c'), type: type, name: name, aliases: JSON.stringify(aliases), updated: nowIso(), author: user.login, status: 'active' };
    writeRow(SHEETS.catalog, CATALOG_COLS, CATALOG_HEAD, r, cur ? cur._row : null);
    appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, { action: before ? 'update' : 'create', editId: r.id, before: before, after: JSON.stringify(publicCatalog(r)), comment: from ? 'переименована запись справочника' : 'добавлена запись справочника', target: 'catalog' })]);
    return { ok: true, record: publicCatalog(r) };
  });
}
function revertCatalog(user, h, comment) {
  var cur = readRows(SHEETS.catalog, CATALOG_COLS, CATALOG_HEAD).filter(function (r) { return r.id === h.editId; })[0];
  if (!cur) return fail('Запись справочника не найдена.');
  var before = cur.status === 'active' ? JSON.stringify(publicCatalog(cur)) : '';
  if (h.before) {
    var prev = JSON.parse(h.before);
    var aliases = (prev.aliases || []).concat(JSON.parse(cur.aliases || '[]'), cur.name).filter(function (v, i, all) { return catalogKey(v) !== catalogKey(prev.name) && all.indexOf(v) === i; });
    cur.type = prev.type; cur.name = prev.name; cur.aliases = JSON.stringify(aliases); cur.status = 'active';
  } else cur.status = 'deleted';
  cur.updated = nowIso(); cur.author = user.login;
  writeRow(SHEETS.catalog, CATALOG_COLS, CATALOG_HEAD, cur, cur._row);
  appendRows(SHEETS.history, HISTORY_COLS, HISTORY_HEAD, [historyEntry(user, { action: 'revert', editId: cur.id, before: before, after: cur.status === 'active' ? JSON.stringify(publicCatalog(cur)) : '', comment: comment, target: 'catalog' })]);
  return { ok: true };
}
