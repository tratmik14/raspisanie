'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),vm=require('node:vm'),path=require('node:path');
const elements={};
const element=id=>elements[id]||(elements[id]={value:'',innerHTML:'',textContent:'',disabled:false,hidden:false,showModal(){this.open=true;},close(){this.open=false;}});
const window={};
let source=fs.readFileSync(path.join(__dirname,'../admin-tools.js'),'utf8');
source=source.replace('  window.RaspAdminTools = {',`  window.TestExam={setContext:function(c){ctx=c;},open:examOpen,save:examSave,retakeOpen:retakeOpen,retakeSave:retakeSave};
  window.RaspAdminTools = {`);
vm.runInNewContext(source,{window,document:{getElementById:element},console});
let records=[{group:'ИС2-25',discipline:'Физика',date:'2027-06-10',time:'08:30',room:'208',teacher:'А.',type:'Экзамен',retakeDate:'',retakeTime:''}];
let shown=null,calls=0;
const context={user:()=>true,suggestions:()=>{},groups:()=>['ИС2-25','ГД-26'],group:()=> 'ГД-26',examGroup:()=> 'ИС2-25',exams:()=>records,esc:v=>v,
  refresh:async()=>{},toast:()=>{},showExam:(r,retake)=>{shown={r,retake};},api:{examsImport:async list=>{calls++;records=list;}}};
window.TestExam.setContext(context);
const ui=window.TestExam;
async function run(){
  for(const type of ['Экзамен','Зачёт','Диф. зачёт']){
    ui.open(null,type);assert.equal(element('exam-edit-type').value,type);assert.equal(element('exam-edit-group').value,'ИС2-25');
  }
  ui.open(null,'Зачёт');
  element('exam-edit-discipline').value='История';element('exam-edit-date').value='2027-01-20';element('exam-edit-time').value='10:10';
  await ui.save(false);assert.equal(records.length,2);assert.equal(records[1].type,'Зачёт');assert.equal(shown.r.type,'Зачёт');assert.equal(shown.retake,false);
  ui.open(1);element('exam-edit-discipline').value='Новая история';await ui.save(false);
  assert.equal(records.length,2);assert.equal(records[1].discipline,'Новая история');
  ui.retakeOpen(0);element('retake-date').value='2027-06-09';await ui.retakeSave(false);assert.equal(calls,2);assert.match(element('retake-error').textContent,/не раньше/);
  element('retake-date').value='2027-06-20';element('retake-time').value='12:00';records.reverse();await ui.retakeSave(false);
  const exam=records.find(r=>r.type==='Экзамен');assert.equal(exam.date,'2027-06-10');assert.equal(exam.retakeDate,'2027-06-20');assert.equal(shown.retake,true);
  ui.retakeOpen(records.indexOf(exam));await ui.retakeSave(true);assert.equal(records.find(r=>r.type==='Экзамен').retakeDate,'');assert.equal(records.length,2);
  ui.open(0);await ui.save(true);assert.equal(records.length,1);
  ui.open(null,null,{group:'ГД-26',discipline:'Математика',teacher:'Б.',type:'Диф. зачёт'});
  assert.equal(element('exam-edit-group').value,'ГД-26');assert.equal(element('exam-edit-type').value,'Диф. зачёт');assert.equal(element('exam-edit-discipline').value,'Математика');
  console.log('Редактор аттестаций: типы, контекст группы, создание, изменение, удаление, назначение из зачётки и сохранение/снятие пересдачи без потери основной записи проверены');
}
run().catch(err=>{console.error(err);process.exitCode=1;});
