'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const root = path.resolve(__dirname, '..');
const Rules = require('../schedule-rules.js');
const Hours = require('../hours.js');
assert.deepEqual(Rules.dates('2026-12-21', '2027-01-18', 7, d => !d.startsWith('2027-01-0')), ['2026-12-21','2026-12-28','2027-01-11','2027-01-18']);
assert.deepEqual(Rules.dates('2026-12-21','2027-01-18',14), ['2026-12-21','2027-01-04','2027-01-18']);
assert.throws(() => Rules.dates('2026-11-12','2026-11-01',7), /даты/);
assert.throws(() => Rules.dates('2026-11-12','2028-11-01',7), /год/);
const series = Rules.expand([{id:'old',group:'Г-26',date:'2026-12-21',kind:'cancel',link:'move'}, {group:'Г-26',date:'2026-12-22',kind:'add',link:'move'}], ['2026-12-21','2026-12-28'], '2026-12-21', r => ({id:r.kind==='cancel'?'existing':'',was:'actual'}));
assert.equal(series[0].id,'old'); assert.equal(series[2].id,'existing'); assert.equal(series[3].date,'2026-12-29');
assert.equal(series[2].link,series[3].link); assert.notEqual(series[0].link,series[2].link); assert.equal(series[3].was,'actual');
const localStorage = { getItem:()=>null };
const elements = {};
const document = {
  querySelectorAll:()=>[],
  getElementById:id=>elements[id]||(elements[id]={value:'',innerHTML:'',textContent:'',querySelectorAll:()=>[],querySelector:()=>({}),showModal(){this.open=true;}})
};
const window = {SCHEDULE_CONFIG:{slots:[['8:30','10:00'],['10:10','11:40'],['12:00','13:30'],['13:40','15:20'],['15:20','16:50']],hoursPerPair:2,holidays:[]},RaspApi:{user:()=>({login:'admin'}),mode:'live'},RaspHours:Hours,RaspRules:Rules};
let source = fs.readFileSync(path.join(root,'app.js'),'utf8');
const marker='  window.RaspAdminTools.init({';
assert.ok(source.includes(marker));
source=source.replace(marker,`  bal={open:'',until:{},wantAll:false}; balMemo={}; PENCIL=''; PLUS='';
  window.testApp={DB:DB,state:state,planFor:planFor,balanceCompute:balanceCompute,dayCardHTML:dayCardHTML,countFor:countFor,openEditor:openEditor,edLoadSlot:edLoadSlot,
    setPlan:function(p){PLAN=p;},setOverrides:function(p){PLAN_OVERRIDES=p;},setGrafik:function(g){GRAFIK={data:g};},
    setEdits:function(e){EDITS={list:e,at:0};}};
  return;
`+marker);
vm.runInNewContext(source,{window,localStorage,document,Date,Intl,URLSearchParams,console});
const a=window.testApp, year=new Date().getMonth()>=6?new Date().getFullYear():new Date().getFullYear()-1;
const first=year+'-11-10', second=(year+1)+'-03-10';
a.setPlan({groups:{'Г-26':{group:'Г-26',subjects:[{name:'Физика',s1:4,s2:6,teacher:'А.',form:['З','Э']}] }},order:['Г-26']});
a.setOverrides([{group:'Г-26',name:'Физика',s1:8,s2:10,teacher:'Б.',f1:'ДЗ',f2:'Э'}]);
assert.equal(a.planFor('Г-26').subjects[0].s1,8);assert.equal(a.planFor('Г-26').subjects[0].teacher,'Б.');
const rows=[{group:'Г-26',day:first,start:'8:30',end:'10:00',disc:'Физика',sub:0},{group:'Г-26',day:second,start:'8:30',end:'10:00',disc:'Физика',sub:0}];
const sheet={rows,byDay:{[first]:[rows[0]],[second]:[rows[1]]},scan:{},at:1};a.DB.sheets['Г-26']=sheet;
a.setEdits([]);
a.state.half=1; const one=a.balanceCompute('Г-26',sheet);
a.state.half=2; const two=a.balanceCompute('Г-26',sheet);
a.state.half=0; const all=a.balanceCompute('Г-26',sheet);
assert.equal(one.total.plan,8);assert.equal(two.total.plan,10);assert.equal(all.total.plan,18);
assert.equal(one.total.all,2);assert.equal(two.total.all,2);assert.equal(all.total.all,4);
assert.equal(all.total.addH,one.total.addH+two.total.addH);
a.DB.slots=window.SCHEDULE_CONFIG.slots.map((s,i)=>({n:i+1,start:s[0],end:s[1]}));a.DB.slotByStart={};a.DB.slots.forEach(s=>a.DB.slotByStart[s.start]=s);
const html=a.dayCardHTML(first,1,[],{mode:'group',name:'Г-26'},first);
assert.equal((html.match(/data-add-start=/g)||[]).length,5,'Все свободные пары доступны даже в пустой день');
const teacherHtml=a.dayCardHTML(first,1,[],{mode:'teacher',name:'Бахаева Г.Р.'},first);
assert.equal((teacherHtml.match(/data-add-start=/g)||[]).length,5,'Свободные слоты доступны и в режиме преподавателя');
sheet.source='live';
a.openEditor({day:first,start:'8:30',groups:['Г-26','Д-26'],kind:'add',chooseGroup:true,teacher:'Бахаева Г.Р.'});
assert.equal(elements['ed-group'].value,'');assert.ok(!elements['ed-group'].innerHTML.includes('value="*"'));
assert.equal(elements['ed-teacher'].value,'Бахаева Г.Р.');assert.equal(elements['ed-save'].disabled,true);
elements['ed-group'].value='Г-26';a.edLoadSlot();
assert.equal(elements['ed-teacher'].value,'Бахаева Г.Р.');assert.equal(elements['ed-save'].disabled,false);
console.log('Расписание: повторение, переносы через границу года, учебный план, оба семестра, год и свободные пары проверены');

