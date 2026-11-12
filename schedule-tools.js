(function () {
  'use strict';
  var c, batch = '', dates = [], loaded = '', loading = 0;
  function $(id) { return document.getElementById(id); }
  function input(id, label, list) { return '<div class="field"><label for="' + id + '">' + label + '</label><input type="text" id="' + id + '" list="' + list + '"></div>'; }
  function options(values) { return values.map(function (v) { return '<option value="' + c.esc(v) + '">' + c.esc(v) + '</option>'; }).join(''); }
  function catalogList(want) {
    var type = $('catalog-type').value;
    $('catalog-entry').innerHTML = '<option value="">Добавить новую запись</option>' + options(c.catalogNames(type));
    if (want) $('catalog-entry').value = want;
    catalogFill();
  }
  function catalogFill() {
    $('catalog-name').value = $('catalog-entry').value;
    $('catalog-save').textContent = $('catalog-entry').value ? 'Сохранить название' : 'Добавить';
    $('catalog-error').textContent = '';
  }
  function catalogOpen() {
    if (!c.user()) return;
    c.prepare(); catalogList(); $('account-dialog').close(); $('catalog-dialog').showModal();
  }
  function catalogSave() {
    var type = $('catalog-type').value, from = $('catalog-entry').value, name = $('catalog-name').value.trim();
    if (!name) { $('catalog-error').textContent = 'Введите название.'; return; }
    if (name !== from && c.catalogNames(type).some(function (v) { return window.RaspCatalog.key(v) === window.RaspCatalog.key(name); })) {
      $('catalog-error').textContent = 'Это название уже используется. Выберите его в списке.'; return;
    }
    var current = c.catalog().find(function (r) { return r.type === type && r.name === from; });
    $('catalog-save').disabled = true; $('catalog-error').textContent = 'Сохраняю…';
    c.api.catalogSave({type:type, name:name, from:from, id:current ? current.id : ''}).then(function () { return c.refresh(); }).then(function () {
      catalogList(name); c.toast(from ? 'Название обновлено в занятиях, плане и экзаменах.' : 'Запись добавлена в общий справочник.');
    }).catch(function (err) { $('catalog-error').textContent = err.message; }).finally(function () { $('catalog-save').disabled = false; });
  }
  function rowElements() { return Array.from($('multi-rows').querySelectorAll('[data-multi-row]')); }
  function multiRows() {
    var chosen = $('multi-date').value;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(chosen)) return;
    dates = $('multi-range').value === 'day' ? [chosen] : [0,1,2,3,4,5].map(function (i) { return c.shift(c.monday(chosen), i); });
    var sel = c.selection(), teacher = sel && sel.mode === 'teacher' ? sel.name : '', room = sel && sel.mode === 'room' ? sel.name : '';
    $('multi-rows').innerHTML = dates.map(function (day) {
      return '<fieldset class="multi-day"><legend>' + c.esc(c.weekday(day) + ', ' + c.fmt(day)) + '</legend><button type="button" class="btn btn-quiet btn-small" data-multi-check-day="' + day + '">Отметить весь день</button>' + c.slots().map(function (s) {
        var key = day + '-' + s.n, desc = c.weekday(day) + ', ' + s.n + ' пара';
        return '<div class="multi-row" data-multi-row data-day="' + day + '" data-start="' + s.start + '" data-end="' + s.end + '"><label class="multi-check"><input type="checkbox" data-multi-on><span>' + s.n + ' пара<br><span class="hint">' + s.start + '–' + s.end + '</span></span></label>' +
          '<input type="text" data-multi-disc list="dl-disc" aria-label="Предмет: ' + desc + '" placeholder="Предмет"><input type="text" data-multi-teacher list="dl-teacher" aria-label="Преподаватель: ' + desc + '" placeholder="Преподаватель" value="' + c.esc(teacher) + '"><input type="text" data-multi-room list="dl-room" aria-label="Аудитория: ' + desc + '" placeholder="Аудитория" value="' + c.esc(room) + '"><p class="hint multi-busy" data-multi-busy></p></div>';
      }).join('') + '</fieldset>';
    }).join('');
    multiPreview();
  }
  function multiOpen(mode, day) {
    if (!c.user()) return;
    c.prepare(); batch = 'r' + Date.now().toString(36) + Math.random().toString(36).slice(2,7); loaded = '';
    var sel = c.selection();
    $('multi-group').innerHTML = '<option value="">Выберите группу</option>' + options(c.groups());
    $('multi-group').value = sel && sel.mode === 'group' ? sel.name : '';
    $('multi-range').value = mode || 'week'; $('multi-date').value = day || c.day();
    $('multi-repeat').value = 'once'; $('multi-until').value = c.until(); $('multi-until').max = c.until();
    $('multi-study').checked = true; $('multi-confirm').checked = false;
    $('multi-common-disc').value = ''; $('multi-common-teacher').value = sel && sel.mode === 'teacher' ? sel.name : '';
    $('multi-common-room').value = sel && sel.mode === 'room' ? sel.name : '';
    $('multi-error').textContent = ''; multiRows(); $('multi-dialog').showModal(); loadGroup();
  }
  function loadGroup() {
    var g = $('multi-group').value, version = ++loading; loaded = '';
    multiPreview(); if (!g) return;
    $('multi-error').textContent = 'Загружаю расписание группы…';
    c.ready(g).then(function () { if (version !== loading) return; loaded = g; $('multi-error').textContent = ''; multiPreview(); })
      .catch(function (err) { if (version === loading) { $('multi-error').textContent = err.message; multiPreview(); } });
  }
  function getList(validate) {
    var group = $('multi-group').value, list = [];
    rowElements().forEach(function (el) {
      if (!el.querySelector('[data-multi-on]').checked) return;
      var disc = el.querySelector('[data-multi-disc]').value.trim(), teacher = el.querySelector('[data-multi-teacher]').value.trim(), room = el.querySelector('[data-multi-room]').value.trim();
      if (validate && !disc) throw new Error('Укажите предмет для каждой отмеченной пары.');
      list.push({group:group,date:el.dataset.day,start:el.dataset.start,end:el.dataset.end,kind:'add',disc:disc,teacher:teacher,room:room,sub:0,was:''});
    });
    return list;
  }
  function multiPreview() {
    var g = $('multi-group').value, selected = getList(false), conflicts = 0;
    rowElements().forEach(function (el) {
      var busy = loaded === g && g ? c.rows(g,el.dataset.day,el.dataset.start).filter(function (r) { return !r.ghost; }) : [];
      el.querySelector('[data-multi-busy]').textContent = busy.length ? 'У группы: ' + busy.map(function (r) { return r.disc; }).join('; ') : '';
      if (el.querySelector('[data-multi-on]').checked && busy.length) conflicts++;
    });
    $('multi-conflict-field').hidden = !conflicts;
    var text = 'Отмечено пар: ' + selected.length + '.';
    if (g && loaded === g && selected.length) try { text += ' Будет сохранено записей: ' + c.expand(selected,$('multi-repeat').value,$('multi-until').value,$('multi-study').checked).length + '.'; } catch (err) { text += ' ' + err.message; }
    if (conflicts) text += ' Занятых слотов у группы: ' + conflicts + '. Для добавления параллельных пар подтвердите ниже.';
    $('multi-preview').textContent = text;
    $('multi-save').disabled = !g || loaded !== g || !selected.length || (conflicts > 0 && !$('multi-confirm').checked);
    $('multi-until').disabled = $('multi-repeat').value === 'once';
    $('multi-study').disabled = $('multi-repeat').value === 'once';
  }
  function multiSave() {
    var list;
    try {
      if (!loaded || loaded !== $('multi-group').value) throw new Error('Выберите группу и дождитесь загрузки.');
      list = getList(true); if (!list.length) throw new Error('Отметьте пары для добавления.');
      list = c.expand(list,$('multi-repeat').value,$('multi-until').value,$('multi-study').checked);
      var busy = list.some(function (r) { return c.rows(r.group,r.date,r.start).some(function (x) { return !x.ghost && (!x.edit || x.edit.id !== r.id); }); });
      if (busy && !$('multi-confirm').checked) { $('multi-conflict-field').hidden = false; throw new Error('Часть выбранных или повторяющихся слотов уже занята. Подтвердите добавление параллельных пар или измените выбор.'); }
    } catch (err) { $('multi-error').textContent = err.message; return; }
    $('multi-save').disabled = true; $('multi-error').textContent = 'Сохраняю ' + list.length + ' записей…';
    c.saveMany(list,batch).then(function () { $('multi-dialog').close(); c.toast('Пары сохранены: ' + list.length + '. Часы и вычитка пересчитаны.'); })
      .catch(function (err) { $('multi-error').textContent = err.message; }).finally(multiPreview);
  }
  function init(ctx) {
    c = ctx;
    var holder = document.createElement('div');
    holder.innerHTML = '<dialog class="dlg dlg-wide" id="catalog-dialog" aria-labelledby="catalog-title"><form onsubmit="return false"><h2 id="catalog-title">Преподаватели и предметы</h2><p class="hint">Добавляйте новые записи или выбирайте существующую для переименования. Новое название применяется к занятиям, учебным планам и экзаменам. Часы сохраняются.</p><div class="row2"><div class="field"><label for="catalog-type">Справочник</label><select id="catalog-type"><option value="teacher">Преподаватели</option><option value="subject">Предметы</option></select></div><div class="field"><label for="catalog-entry">Запись</label><select id="catalog-entry"></select></div></div>' + input('catalog-name','Новое название','') + '<p class="error" id="catalog-error" role="alert"></p><div class="dlg-actions"><button type="button" class="btn btn-quiet" data-close-dialog="catalog-dialog">Закрыть</button><button type="button" class="btn btn-primary" id="catalog-save">Добавить</button></div></form></dialog>' +
      '<dialog class="dlg multi-dialog" id="multi-dialog" aria-labelledby="multi-title"><form onsubmit="return false"><h2 id="multi-title">Добавить пары на день или неделю</h2><p class="hint">Выберите группу, отметьте пары и заполните их. Существующие занятия сохранятся. Новые пары будут видны в расписании группы, преподавателя и аудитории.</p><div class="row2"><div class="field"><label for="multi-group">Группа</label><select id="multi-group"></select></div><div class="field"><label for="multi-date">Дата</label><input type="date" id="multi-date"></div></div><div class="field"><label for="multi-range">Заполнить</label><select id="multi-range"><option value="day">Выбранный день</option><option value="week">Неделю с понедельника по субботу</option></select></div><fieldset class="repeat-fields"><legend>Общие значения для отмеченных пар</legend>' + input('multi-common-disc','Предмет','dl-disc') + '<div class="row2">' + input('multi-common-teacher','Преподаватель','dl-teacher') + input('multi-common-room','Аудитория','dl-room') + '</div><button type="button" class="btn btn-small" id="multi-apply">Заполнить отмеченные пары</button><button type="button" class="btn btn-quiet btn-small" id="multi-check-all">Отметить все пары</button></fieldset><div id="multi-rows"></div><fieldset class="repeat-fields"><legend>Повторение</legend><div class="row2"><div class="field"><label for="multi-repeat">На какие недели</label><select id="multi-repeat"><option value="once">Только выбранный день или неделя</option><option value="term">До конца семестра</option><option value="year">Следующие недели учебного года</option><option value="alternate">Через неделю</option></select></div><div class="field"><label for="multi-until">Повторять по</label><input type="date" id="multi-until"></div></div><label class="repeat-check"><input type="checkbox" id="multi-study" checked>Пропускать праздники, каникулы, практику и сессии</label></fieldset><p id="multi-preview" class="hint" aria-live="polite"></p><label class="repeat-check" id="multi-conflict-field" hidden><input type="checkbox" id="multi-confirm">Добавить параллельные пары в занятые слоты</label><p class="error" id="multi-error" role="alert"></p><div class="dlg-actions"><button type="button" class="btn btn-quiet" data-close-dialog="multi-dialog">Закрыть</button><button type="button" class="btn btn-primary" id="multi-save">Сохранить пары</button></div></form></dialog>';
    while (holder.firstChild) document.body.append(holder.firstChild);
    $('catalog-type').addEventListener('change', function () { catalogList(); }); $('catalog-entry').addEventListener('change',catalogFill); $('catalog-save').addEventListener('click',catalogSave);
    $('multi-group').addEventListener('change',loadGroup);
    ['multi-date','multi-range'].forEach(function (id) { $(id).addEventListener('change',function () { multiRows(); $('multi-confirm').checked = false; }); });
    ['multi-repeat','multi-until','multi-study','multi-confirm'].forEach(function (id) { $(id).addEventListener('change',multiPreview); });
    $('multi-rows').addEventListener('input',multiPreview);
    $('multi-check-all').addEventListener('click',function () { rowElements().forEach(function (el) { el.querySelector('[data-multi-on]').checked = true; }); multiPreview(); });
    $('multi-apply').addEventListener('click',function () { rowElements().forEach(function (el) { if (!el.querySelector('[data-multi-on]').checked) return; ['disc','teacher','room'].forEach(function (key) { el.querySelector('[data-multi-' + key + ']').value = $('multi-common-' + key).value; }); }); multiPreview(); });
    $('multi-save').addEventListener('click',multiSave);
    document.addEventListener('click',function (event) {
      var b = event.target.closest('button'); if (!b) return;
      if (b.hasAttribute('data-catalog-open')) catalogOpen();
      if (b.hasAttribute('data-multi-open')) multiOpen(b.getAttribute('data-multi-open'));
      if (b.hasAttribute('data-multi-day')) multiOpen('day',b.getAttribute('data-multi-day'));
      if (b.hasAttribute('data-multi-check-day')) { rowElements().forEach(function (el) { if (el.dataset.day === b.getAttribute('data-multi-check-day')) el.querySelector('[data-multi-on]').checked = true; }); multiPreview(); }
    });
  }
  window.RaspScheduleTools = {init:init};
})();
