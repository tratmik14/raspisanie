'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
const path=require('node:path');
const {webcrypto}=require('node:crypto');
const root=path.resolve(__dirname,'..');
const storage=new Map();
const window={crypto:webcrypto};
vm.runInNewContext(fs.readFileSync(path.join(root,'gas-sim.js'),'utf8'),{
  window,crypto:webcrypto,TextEncoder,Uint8Array,Array,Date,JSON,Math,Promise,Number,
  localStorage:{getItem:k=>storage.get(k)||null,setItem:(k,v)=>storage.set(k,v),removeItem:k=>storage.delete(k)},
  fetch:async url=>({ok:true,text:async()=>fs.readFileSync(path.join(root,url.split('?')[0]),'utf8')})
});
async function run(){
  const call=x=>window.RaspSim.call(x);
  assert.equal((await call({action:'planSave',record:{}})).ok,false);
  const login=await call({action:'login',login:'demo',password:'demo'});assert.equal(login.ok,true);
  const token=login.token, record={group:'ИС2-25',name:'Физика',s1:44,s2:56,teacher:'Тест А.А.',f1:'З',f2:'Э'};
  assert.equal((await call({action:'planSave',token,record:{...record,s1:-1}})).ok,false);
  assert.equal((await call({action:'planSave',token,record:{...record,f1:'invalid'}})).ok,false);
  assert.equal((await call({action:'planSave',token,record})).ok,true);
  let state=await call({action:'state'});assert.equal(state.plans[0].s1,44);assert.equal(state.plans[0].author,undefined);
  const h=await call({action:'history',token});assert.equal(h.items[0].target,'plan');
  assert.equal((await call({action:'planSave',token,record:{...record,s1:66}})).ok,true);
  const update=await call({action:'history',token});
  assert.equal((await call({action:'revert',token,id:update.items[0].id})).ok,true);
  assert.equal((await call({action:'state'})).plans[0].s1,44);
  assert.equal((await call({action:'revert',token,id:h.items[0].id})).ok,true);
  assert.equal((await call({action:'state'})).plans.length,0);
  assert.equal((await call({action:'table',sheet:'Планы'})).ok,false);
  assert.equal((await call({action:'catalogSave',record:{type:'teacher',name:'Тест'}})).ok,false);
  let saved=await call({action:'catalogSave',token,record:{type:'teacher',name:'Тест А.А.'}}); assert.equal(saved.ok,true);
  const catalogId=saved.record.id;
  saved=await call({action:'catalogSave',token,record:{id:catalogId,type:'teacher',from:'Тест А.А.',name:'Тест Б.Б.'}});assert.equal(saved.ok,true);
  assert.deepEqual(Array.from(saved.record.aliases),['Тест А.А.']);
  saved=await call({action:'catalogSave',token,record:{id:catalogId,type:'teacher',from:'Тест Б.Б.',name:'Тест В.В.'}});assert.equal(saved.ok,true);
  assert.ok(saved.record.aliases.includes('Тест А.А.'));assert.ok(saved.record.aliases.includes('Тест Б.Б.'));
  assert.equal((await call({action:'catalogSave',token,record:{type:'teacher',name:'Тест А.А.'}})).ok,false);
  assert.equal((await call({action:'catalogSave',token,record:{type:'subject',name:'Тест А.А.'}})).ok,true);
  assert.equal((await call({action:'catalogSave',token,record:{type:'teacher',name:'А., Б.'}})).ok,false);
  const catHistory=await call({action:'history',token});
  const rename=catHistory.items.find(x=>x.target==='catalog' && JSON.parse(x.after||'{}').name==='Тест В.В.');
  assert.equal((await call({action:'revert',token,id:rename.id})).ok,true);
  assert.equal((await call({action:'state'})).catalog.find(x=>x.id===catalogId).name,'Тест Б.Б.');
  assert.equal((await call({action:'table',sheet:'Справочники'})).ok,false);
  const edits=Array.from({length:30},(_,i)=>({group:'ИС2-25',date:'2026-11-'+(16+Math.floor(i/5)),start:['8:30','10:10','12:00','13:40','15:20'][i%5],end:'16:50',disc:'Предмет',teacher:'Тест Б.Б.',room:'302',kind:'add',sub:0}));
  const bulk=await call({action:'save',token,edits,batch:'rtestweek'});assert.equal(bulk.ok,true);assert.equal(bulk.edits.length,30);
  assert.equal((await call({action:'state'})).edits.length,30);
  console.log('Справочники: авторизация, добавление, цепочка переименований, дубликаты, публичное чтение и откат; пакет 30 пар проверены');
  const attestations=['Экзамен','Зачёт','Диф. зачёт'].map((type,i)=>({group:'ИС2-25',discipline:'Аттестация '+i,date:'2027-01-20',time:'10:00',room:'302',teacher:'А.',type,retakeDate:'',retakeTime:''}));
  assert.equal((await call({action:'examsImport',token,records:attestations})).ok,true);
  let exams=(await call({action:'state'})).exams;assert.equal(exams.length,3);assert.equal(exams[1].type,'Зачёт');
  attestations[1].retakeDate='2027-01-25';attestations[1].retakeTime='12:00';
  assert.equal((await call({action:'examsImport',token,records:attestations})).ok,true);
  exams=(await call({action:'state'})).exams;assert.equal(exams[1].date,'2027-01-20');assert.equal(exams[1].retakeDate,'2027-01-25');assert.equal(exams.length,3);
  console.log('Сервер: экзамены, зачёты, дифзачёты и даты пересдач сохраняются и читаются из общей базы');
  console.log('Сервер: авторизация, сохранение плана, проверка полей, публичное чтение и оба вида отката проверены');
}
run().catch(e=>{console.error(e);process.exitCode=1;});
