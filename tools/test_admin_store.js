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
  console.log('Сервер: авторизация, сохранение плана, проверка полей, публичное чтение и оба вида отката проверены');
}
run().catch(e=>{console.error(e);process.exitCode=1;});

