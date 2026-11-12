(function () {
  'use strict';
  var ctx, planGroup, examOriginal;
  var $ = function (id) { return document.getElementById(id); };
  function field(label, id, type, extra) {
    return '<div class="field"><label for="' + id + '">' + label + '</label><input id="' + id + '" type="' + (type || 'text') + '" ' + (extra || '') + '></div>';
  }
  function forms(id) {
    return '<div class="field"><label for="' + id + '">Форма аттестации, ' + (id === 'plan-f1' ? '1' : '2') + ' семестр</label><select id="' + id + '">' +
      [['', 'Не указана'], ['Э', 'Экзамен'], ['З', 'Зачёт'], ['ДЗ', 'Диф. зачёт'], ['ЗаО', 'Зачёт с оценкой'], ['К/Р', 'Контрольная работа'], ['ДР', 'Другая форма'], ['Проект', 'Защита проекта']].map(function (x) { return '<option value="' + x[0] + '">' + x[1] + '</option>'; }).join('') + '</select></div>';
  }
  function groups(select, value) {
    select.innerHTML = ctx.groups().map(function (g) { return '<option value="' + ctx.esc(g) + '">' + ctx.esc(g) + '</option>'; }).join('');
    select.value = value || ctx.group();
  }
  function planSelect() {
    var plan = ctx.plan($('plan-group').value), subjects = plan ? plan.subjects : [];
    $('plan-subject').innerHTML = subjects.map(function (s, i) { return '<option value="' + i + '">' + ctx.esc(s.name) + '</option>'; }).join('') + '<option value="new">Добавить предмет</option>';
    planFill();
  }
  function planFill() {
    var plan = ctx.plan($('plan-group').value), v = $('plan-subject').value;
    var s = plan && v !== 'new' ? plan.subjects[+v] : null;
    $('plan-name').value = s ? s.name : '';
    $('plan-name').readOnly = !!s;
    $('plan-s1').value = s ? s.s1 : 0;
    $('plan-s2').value = s ? s.s2 : 0;
    $('plan-teacher').value = s ? s.teacher || '' : '';
    $('plan-f1').value = s && s.form ? s.form[0] : '';
    $('plan-f2').value = s && s.form ? s.form[1] : '';
    $('plan-error').textContent = '';
  }
  function planOpen() {
    if (!ctx.user()) return;
    ctx.suggestions();
    groups($('plan-group'), ctx.group());
    planSelect(); $('plan-dialog').showModal();
  }
  function planSave() {
    var r = { group: $('plan-group').value, name: $('plan-name').value.trim(), s1: +$('plan-s1').value, s2: +$('plan-s2').value,
      teacher: $('plan-teacher').value.trim(), f1: $('plan-f1').value, f2: $('plan-f2').value };
    if (!r.group || !r.name || !Number.isInteger(r.s1) || !Number.isInteger(r.s2) || r.s1 < 0 || r.s2 < 0 || r.s1 > 10000 || r.s2 > 10000) {
      $('plan-error').textContent = 'Укажите предмет и целое число часов от 0 до 10 000 для каждого семестра.'; return;
    }
    $('plan-save').disabled = true; $('plan-save').textContent = 'Сохраняю…'; $('plan-error').textContent = '';
    ctx.api.planSave(r).then(function () { return ctx.refresh(); }).then(function () {
      $('plan-dialog').close(); ctx.toast('Учебный план сохранён. Часы, вычитка и формы аттестации пересчитаны.');
    }).catch(function (err) { $('plan-error').textContent = err.message; }).finally(function () { $('plan-save').disabled = false; $('plan-save').textContent = 'Сохранить учебный план'; });
  }
  function examKey(r) { return [r.group, r.discipline, r.date, r.time || ''].join('|'); }
  function examRecord(r) {
    var out = {};
    ['group', 'discipline', 'date', 'time', 'room', 'teacher', 'type', 'retakeDate', 'retakeTime'].forEach(function (k) { out[k] = r[k] || ''; });
    return out;
  }
  function examOpen(index) {
    if (!ctx.user()) return;
    var r = index == null ? null : ctx.exams()[+index];
    examOriginal = r ? examKey(r) : null;
    groups($('exam-edit-group'), r ? r.group : ctx.group());
    ['discipline', 'date', 'time', 'room', 'teacher', 'retakeDate', 'retakeTime'].forEach(function (k) { $('exam-edit-' + k).value = r ? r[k] || '' : ''; });
    $('exam-edit-type').value = r ? r.type : 'Экзамен';
    $('exam-edit-title').textContent = r ? 'Изменить экзамен или зачёт' : 'Добавить экзамен или зачёт';
    $('exam-edit-remove').hidden = !r;
    $('exam-edit-error').textContent = '';
    $('exam-edit-dialog').showModal();
  }
  function examSave(remove) {
    var r = { group: $('exam-edit-group').value, type: $('exam-edit-type').value };
    ['discipline', 'date', 'time', 'room', 'teacher', 'retakeDate', 'retakeTime'].forEach(function (k) { r[k] = $('exam-edit-' + k).value.trim(); });
    if (!remove && (!r.group || !r.discipline || !r.date || (r.retakeTime && !r.retakeDate))) { $('exam-edit-error').textContent = 'Укажите группу, дисциплину и дату. Для времени пересдачи нужна её дата.'; return; }
    var all = ctx.exams().filter(function (x) { return !examOriginal || examKey(x) !== examOriginal; }).map(examRecord);
    if (!remove) all.push(r);
    if (!remove && all.filter(function (x) { return examKey(x) === examKey(r); }).length > 1) { $('exam-edit-error').textContent = 'Такая запись уже есть. Откройте её для изменения.'; return; }
    $('exam-edit-save').disabled = true; $('exam-edit-remove').disabled = true;
    $('exam-edit-error').textContent = 'Сохраняю…';
    ctx.api.examsImport(all, 'Редактирование на сайте').then(function () { return ctx.refresh(); }).then(function () {
      $('exam-edit-dialog').close(); ctx.toast(remove ? 'Запись убрана. Предыдущий список можно вернуть в журнале изменений.' : 'Экзамены сохранены. Зачётка и расписание преподавателя обновлены.');
    }).catch(function (err) { $('exam-edit-error').textContent = err.message; }).finally(function () { $('exam-edit-save').disabled = false; $('exam-edit-remove').disabled = false; });
  }
  function init(context) {
    ctx = context;
    window.RaspScheduleTools.init(context);
    ['hours', 'balance', 'exams'].forEach(function (view) {
      var bar = document.createElement('div'); bar.className = 'admin-tools'; bar.id = 'admin-tools-' + view; bar.hidden = true;
      bar.innerHTML = view === 'exams' ? '<button type="button" class="btn btn-small" data-exam-new>Добавить экзамен или зачёт</button><span class="hint">Изменить запись можно кнопкой с карандашом рядом с дисциплиной.</span>' : '<button type="button" class="btn btn-small" data-plan-open>Изменить учебный план</button><span class="hint">Проведённые часы считаются по занятиям. План, преподавателей и формы аттестации можно изменить здесь.</span>';
      $('view-' + view).prepend(bar);
    });
    var holder = document.createElement('div');
    holder.innerHTML = '<dialog class="dlg dlg-wide" id="plan-dialog" aria-labelledby="plan-title"><form method="dialog" onsubmit="return false"><h2 id="plan-title">Учебный план группы</h2><p class="hint">Изменения сохраняются в общей базе и учитываются в часах, вычитке и зачётке. Проведённые часы меняются при добавлении и отмене занятий.</p>' +
      '<div class="row2"><div class="field"><label for="plan-group">Группа</label><select id="plan-group"></select></div><div class="field"><label for="plan-subject">Предмет</label><select id="plan-subject"></select></div></div>' + field('Название предмета', 'plan-name', 'text', 'list="dl-disc"') +
      '<div class="row2">' + field('План часов, 1 семестр', 'plan-s1', 'number', 'min="0" max="10000" step="1"') + field('План часов, 2 семестр', 'plan-s2', 'number', 'min="0" max="10000" step="1"') + '</div>' +
      '<div class="row2">' + forms('plan-f1') + forms('plan-f2') + '</div>' + field('Преподаватель', 'plan-teacher', 'text', 'list="dl-teacher"') +
      '<p class="error" id="plan-error" role="alert"></p><div class="dlg-actions"><button type="button" class="btn btn-quiet" data-close-dialog="plan-dialog">Закрыть</button><button type="button" class="btn btn-primary" id="plan-save">Сохранить учебный план</button></div></form></dialog>' +
      '<dialog class="dlg dlg-wide" id="exam-edit-dialog" aria-labelledby="exam-edit-title"><form method="dialog" onsubmit="return false"><h2 id="exam-edit-title">Экзамен или зачёт</h2>' +
      '<div class="row2"><div class="field"><label for="exam-edit-group">Группа</label><select id="exam-edit-group"></select></div><div class="field"><label for="exam-edit-type">Тип аттестации</label><select id="exam-edit-type"><option>Экзамен</option><option>Зачёт</option><option>Диф. зачёт</option></select></div></div>' +
      field('Дисциплина', 'exam-edit-discipline', 'text', 'list="dl-disc"') + '<div class="row2">' + field('Дата', 'exam-edit-date', 'date') + field('Время', 'exam-edit-time', 'time') + '</div>' +
      '<div class="row2">' + field('Преподаватель', 'exam-edit-teacher', 'text', 'list="dl-teacher"') + field('Аудитория', 'exam-edit-room', 'text', 'list="dl-room"') + '</div>' +
      '<div class="row2">' + field('Дата пересдачи', 'exam-edit-retakeDate', 'date') + field('Время пересдачи', 'exam-edit-retakeTime', 'time') + '</div>' +
      '<p class="hint">Занятия в день экзамена проверяйте в разделе «Занятия». При ручном редактировании экзамена они не отменяются автоматически.</p><p class="error" id="exam-edit-error" role="alert"></p><div class="dlg-actions"><button type="button" class="btn btn-quiet btn-danger" id="exam-edit-remove">Убрать запись</button><span class="spacer"></span><button type="button" class="btn btn-quiet" data-close-dialog="exam-edit-dialog">Закрыть</button><button type="button" class="btn btn-primary" id="exam-edit-save">Сохранить</button></div></form></dialog>';
    while (holder.firstChild) document.body.append(holder.firstChild);
    var catalogButton = document.createElement('button'); catalogButton.type = 'button'; catalogButton.className = 'btn'; catalogButton.setAttribute('data-catalog-open',''); catalogButton.textContent = 'Преподаватели и предметы';
    $('account-dialog').querySelector('.dlg-actions').prepend(catalogButton);
    $('plan-group').addEventListener('change', planSelect);
    $('plan-subject').addEventListener('change', planFill);
    $('plan-save').addEventListener('click', planSave);
    $('exam-edit-save').addEventListener('click', function () { examSave(false); });
    $('exam-edit-remove').addEventListener('click', function () { examSave(true); });
    document.addEventListener('click', function (event) {
      var target = event.target.closest('button'); if (!target) return;
      if (target.hasAttribute('data-plan-open')) planOpen();
      if (target.hasAttribute('data-exam-new')) examOpen(null);
      if (target.hasAttribute('data-exam-edit')) examOpen(target.getAttribute('data-exam-edit'));
      if (target.hasAttribute('data-close-dialog')) $(target.getAttribute('data-close-dialog')).close();
    });
  }
  window.RaspAdminTools = { init: init, render: function () {
    if (!ctx) return;
    ['hours', 'balance', 'exams'].forEach(function (view) { $('admin-tools-' + view).hidden = !ctx.user(); });
  } };
})();
