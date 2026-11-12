(function (root) {
  'use strict';
  function key(s) { return String(s || '').replace(/\s+/g, ' ').trim().toLocaleLowerCase('ru'); }
  function name(items, type, value) {
    var found = items.find(function (r) { return r.type === type && [r.name].concat(r.aliases || []).some(function (v) { return key(v) === key(value); }); });
    return found ? found.name : value;
  }
  function splitTeachers(value) { return String(value || '').split(/\s*[,;]\s*|\s+(?=[А-ЯЁ][а-яё]+(?:-[А-ЯЁ][а-яё]+)?\s+[А-ЯЁ]\.)/).map(function (v) { return v.trim(); }).filter(Boolean); }
  function teachers(items, value) { return splitTeachers(value).map(function (v) { return name(items, 'teacher', v); }).join(', '); }
  var api = { key: key, name: name, teachers: teachers, splitTeachers: splitTeachers };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.RaspCatalog = api;
})(typeof window === 'object' ? window : this);
