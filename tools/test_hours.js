// Проверки подсчёта часов и вычитки: node tools/test_hours.js
'use strict';
const assert = require('assert');
const H = require('../hours.js');
const opt = { today: '2026-11-02', nowMin: 0, half2: '2027-01-11', hoursPerPair: 2 };
const row = (day, start, disc, sub) => ({ day, start, end: '10:00', disc, sub: sub || 0, teacher: 'Т.' });

// подгруппы в одно время — одна пара, в разное время — тоже не две
let c = H.countHours([row('2026-10-01', '8:30', 'Англ', 1), row('2026-10-01', '8:30', 'Англ', 2)], opt);
assert.strictEqual(c['Англ'].h[1].all, 2);
c = H.countHours([row('2026-10-01', '8:30', 'Англ', 1), row('2026-10-02', '8:30', 'Англ', 2), row('2026-10-03', '8:30', 'Англ', 0)], opt);
assert.strictEqual(c['Англ'].h[1].all, 4);
assert.strictEqual(c['Англ'].h[1].done, 4);

// годовые часы в колонке 1-го семестра не считаются семестровыми
const plan = H.parsePlan([
  ['Группа СПм - 25 Курс 2'], ['№', 'Наименование дисциплины', 'Кол-во часов в год', 'в 1 с.', 'Форма', 'во 2 с.', 'ФИО преподавателя'],
  ['', 'МДК 03.01 Технология', '', '144', 'К.р.', '9/144 4с', 'Истратова С.М.'],
  ['', 'Физика', '108', '4,5/50', '-/ДЗ', '4/58', 'Меринова Т.С.'],
  ['', 'Физика', '108', '4,5/50', '-/ДЗ', '4/58', 'Меринова Т.С.']
]);
const g = plan.groups['СПм-25'];
assert.strictEqual(g.subjects.length, 2, 'повтор предмета в блоке плана убран');
assert.strictEqual(g.subjects[0].s1, 0);
assert.strictEqual(g.subjects[0].s2, 144);

// вычитка: добавить, снять, практика
const sum = H.summarize({ subjects: [
  { name: 'Физика', s1: 50, s2: 0, form: ['', ''] }, { name: 'История', s1: 20, s2: 0, form: ['', ''] },
  { name: 'УП.01 Учебная практика', s1: 36, s2: 0, form: ['', ''] }
] }, H.countHours([
  ...Array.from({ length: 20 }, (_, i) => row('2026-09-' + String(i + 1).padStart(2, '0'), '8:30', 'Физика')),
  ...Array.from({ length: 12 }, (_, i) => row('2026-09-' + String(i + 1).padStart(2, '0'), '10:10', 'История'))
], opt), 1);
const b = H.balance(sum, 2);
const by = n => b.lines.find(l => l.name === n);
assert.strictEqual(by('Физика').status, 'add'); assert.strictEqual(by('Физика').pairs, 5);
assert.strictEqual(by('История').status, 'cut'); assert.strictEqual(by('История').pairs, 2);
assert.strictEqual(by('УП.01 Учебная практика').status, 'practice');
assert.deepStrictEqual([b.total.add, b.total.cut], [5, 2]);

// срок: конец теории и последний день перед сессией, одиночный день аттестации посреди практики не в счёт
const days = 'UUTTT UUTTT APPPP PPPPP PAAAA KKKKK'.split(' ').flatMap((w, i) => w.split('').map((c2, d) => ({ iso: 'w' + i + 'd' + d, c: c2 })));
assert.deepStrictEqual(H.balanceDeadline(days), { theoryEnd: 'w1d4', until: 'w4d0' });
console.log('hours.js: все проверки прошли');
