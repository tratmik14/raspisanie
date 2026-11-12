'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const code = name => fs.readFileSync(path.join(root, name), 'utf8');
const sheets = {
  'Список групп': [['ИС2-25']],
  'ИС2-25': [['Предмет, с запятой', 'Текст "в кавычках"', 'Строка\nвторая']],
  'Админы': [['Секрет']],
  '0 Учебные графики': [['План']]
};
const ranges = [];
const ctx = vm.createContext({
  SpreadsheetApp: { openById: id => {
    assert.equal(id, '1SqSkidIPZ9eFIgfINub6y3kWxBgQB8tAd3JrxXtlxw8');
    return { getSheetByName: name => sheets[name] && {
    getLastRow: () => sheets[name].length,
    getLastColumn: () => 20,
    getRange: (r, c, height, width) => {
      ranges.push({name, width});
      return { getDisplayValues: () => sheets[name] };
    }
  } }; } }
});
vm.runInContext(code('apps-script/Code.gs'), ctx);
assert.equal(ctx.route({action:'table',sheet:'Админы'}).ok, false);
assert.equal(ctx.route({action:'table',sheet:'История'}).ok, false);
sheets['Список групп'].push(['Админы']);
assert.equal(ctx.route({action:'table',sheet:'Админы'}).ok, false);
assert.equal(ctx.route({action:'table',sheet:'Нет такой группы'}).ok, false);
const res = ctx.route({action:'table',sheet:'ИС2-25'});
assert.equal(res.ok, true);
assert.equal(res.database, '1SqSkidIPZ9eFIgfINub6y3kWxBgQB8tAd3JrxXtlxw8');
assert.equal(res.csv, '"Предмет, с запятой","Текст ""в кавычках""","Строка\nвторая"');
assert.equal(ranges.find(x => x.name === 'ИС2-25').width, 9);
assert.equal(ctx.route({action:'save',edits:[]}).ok, false);
const config = {};
vm.runInNewContext(code('config.js'), {window:config});
assert.equal(config.SCHEDULE_CONFIG.dataSource, 'database');
assert.equal(config.SCHEDULE_CONFIG.sheetId, res.database);
async function checkClient() {
  let reply = {ok:true,csv:'new data',database:res.database};
  let requested;
  const window = {SCHEDULE_CONFIG:{...config.SCHEDULE_CONFIG,adminApi:'https://example.test/exec'}};
  vm.runInNewContext(code('api.js'), {
    window, localStorage:{getItem:()=>null},
    fetch:async url=>{requested=url;return{ok:true,json:async()=>reply};}
  });
  assert.equal(await window.RaspApi.loadTable('ИС2-25'), 'new data');
  assert.ok(requested.includes('action=table'));
  assert.ok(!requested.includes('token='));
  reply = {...reply,database:'old-database'};
  await assert.rejects(window.RaspApi.loadTable('ИС2-25'), /другой базы/);
  console.log('database: чтение, изоляция источника и запрет служебных листов проверены');
}
checkClient().catch(err=>{console.error(err);process.exitCode=1;});

