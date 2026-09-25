const fs=require('node:fs'),assert=require('node:assert/strict');
const {JSDOM,VirtualConsole}=require('jsdom');
const errors=[];const vc=new VirtualConsole();vc.on('jsdomError',e=>errors.push(e.message));
const html=fs.readFileSync('index.html','utf8').replace('<script src="solar-terms.js"></script>','<script>'+fs.readFileSync('solar-terms.js','utf8')+'</script>');
const dom=new JSDOM(html,{url:'https://example.test',runScripts:'dangerously',virtualConsole:vc,beforeParse(w){w.TextEncoder=TextEncoder;w.TextDecoder=TextDecoder;w.localStorage.setItem('ug_welcomed','true');}});
const w=dom.window,r=s=>w.eval(s);
setTimeout(()=>{
try{
 assert.ok(w.document.querySelector('.dash'));
 assert.equal(r('dateKey(koreaDate(Date.parse("2026-09-24T15:00:00Z")))'),'2026-09-25');
 assert.equal(r('solarTermStatus(Date.parse("2026-09-25T00:00:00Z")).current.name'),'추분');
 assert.equal(r('calendarTerms(2026).length'),24);
 assert.equal(r('solarTermStatus(calendarTerms(2026)[17].at-1).current.name'),'백로');
 assert.equal(r('solarTermStatus(calendarTerms(2026)[17].at).current.name'),'추분');
 r('go("calendar");setCalendarMonth("2024-02")');assert.equal(w.document.querySelectorAll('.calendar-day').length,29);
 r('setCalendarMonth("2023-02")');assert.equal(w.document.querySelectorAll('.calendar-day').length,28);
 r('setCalendarMonth("2026-12");moveCalendar(1)');assert.equal(w.document.querySelector('input[type=month]').value,'2027-01');
 r('setCalendarMonth("1891-01");moveCalendar(-1)');assert.equal(w.document.querySelector('input[type=month]').value,'1891-01');
 r('people=[{id:"dtest",name:"<img src=x>",gender:"M",cal:"S",y:1983,mo:12,d:12,h:9,min:0,memo:"<script>secret</script>",yearMemos:[{id:"old",year:2026,text:"기존 기록"}]}];currentId="dtest";homeDepth={};view="home";render()');
 const before=r('JSON.stringify(people)');
 assert.equal(w.document.querySelectorAll('.fortune-col.is-filled').length,0);
 assert.equal(r('JSON.stringify(window._chartRelations)'),r('JSON.stringify(E.relations(computed(people[0]).pillars,[]))'));
 r('fillFortuneSlot(4)');assert.equal(w.document.querySelectorAll('.fortune-col.is-filled').length,1);
 r('fillFortuneSlot(1)');assert.equal(w.document.querySelectorAll('.fortune-col.is-filled').length,2);
 r('resetFortuneSlots()');assert.equal(w.document.querySelectorAll('.fortune-col.is-filled').length,0);

 assert.equal(w.document.querySelectorAll('.card-pillars .pcol').length,8);
 r('quickFortune(1)');assert.equal(w.document.querySelectorAll('.quick-fortunes .pcol').length,4);
 r('quickFortune(2)');assert.equal(w.document.querySelectorAll('.quick-fortunes .pcol').length,4);
 r('quickFortune(3)');assert.equal(w.document.querySelectorAll('.quick-fortunes .pcol').length,4);
 w.document.querySelector('.person-name').click();assert.equal(w.document.querySelectorAll('.card-pillars .pcol').length,8);
 assert.equal(w.document.querySelectorAll('.card-pillars .relbar').length,0);
 w.document.querySelector('.hid button').click();assert.match(w.document.querySelector('#infoBody').textContent,/일간 甲 기준/);
 r('closeInfo();highlightRelation(0)');assert.ok(w.document.querySelector('.relation-hit'));
 r('go("dashboard")');assert.equal(w.document.querySelectorAll('.dash img,.dash script').length,0);
 assert.equal(r('JSON.stringify(people)'),before);
 r('openCalendarChart(2027,8,1)');assert.equal(w.document.querySelector('.nb-year').value,'2027');
 r('notebookYear(2028)');assert.equal(r('window._homeYearShown'),2028);
 assert.deepEqual(errors,[]);console.log('PASS dashboard, KST/term boundaries, calendar leap years, progressive pillars, hidden stems, safe notes');
}finally{dom.window.close();}
},50);
