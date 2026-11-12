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

