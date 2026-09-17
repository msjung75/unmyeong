const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { JSDOM, VirtualConsole } = require('jsdom');

const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const scripts = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)];
scripts.forEach((m, i) => new vm.Script(m[1], { filename: `inline-${i}.js` }));
const errors = [];
const virtualConsole = new VirtualConsole();
virtualConsole.on('jsdomError', error => errors.push(error.message));
const dom = new JSDOM(html, {
  url: 'https://example.test/unmyeong/', runScripts: 'dangerously',
  pretendToBeVisual: true, virtualConsole,
  beforeParse(w) {
    w.TextEncoder = TextEncoder; w.TextDecoder = TextDecoder;
    w.URL.createObjectURL = () => 'blob:test'; w.URL.revokeObjectURL = () => {};
    w.localStorage.setItem('ug_welcomed', 'true');
    w.print = () => {};
  }
});
const w = dom.window;
const run = code => w.eval(code);
const fixture = (overrides = {}) => ({id:'regression',name:'검토용',gender:'M',cal:'S',y:1990,mo:1,d:1,h:9,min:30,noTime:false,memo:'비공개 메모 검증',...overrides});
const use = record => {
  w.testRecord = record;
  run('people=[window.testRecord];currentId=window.testRecord.id;selLuck={};selYear={};settings.correction=true;settings.lateNight=false;view="home";render();');
};

(async () => {
  await new Promise(resolve => setTimeout(resolve, 30));
  assert.equal(w.SajuEngine !== undefined, true);
  use(fixture());
  assert.equal(run('computed(people[0]).yearNum'), 1989, '입춘 연주는 유지');
  assert.equal(run('computed(people[0]).solar.y'), 1990, '실제 출생연도 분리');
  run('goYear(2026)');
  assert.equal(w.document.querySelector('.badge.eq').textContent, '37세');
  assert.equal(run('luckAt(computed(people[0]),2026).age'), 37);
  assert.equal(run('luckAt(computed(people[0]),2026).luck.age'), 28, '대운 전환을 1년 앞당기지 않음');
  assert.match(run('renderMemoView()'), /2026년 37세/);
  assert.match(run('buildAiPrompt(people[0],2026)'), /조회 2026년 37세/);
  run('goYear(2027)');
  assert.equal(run('luckAt(computed(people[0]),2027).luck.age'), 38);
  assert.equal(w.document.querySelector('.badge.eq').textContent, '38세');
  run('window._nowPid=currentId;window._nowDT={y:2027,mo:1,d:1,h:12,min:0};renderNow();');
  assert.match(w.document.querySelector('#nowBody .pillars').textContent,/38세~/);

  for (const rec of [fixture({y:1983,mo:12,d:12,h:9,min:0}),fixture({y:1990,mo:2,d:5}),fixture({cal:'L',y:1989,mo:12,d:5}),fixture({y:1990,h:0,min:10}),fixture({noTime:true,h:null,min:null})]) {
    use(rec); run('goYear(2027)');
    const age = run('2027-computed(people[0]).solar.y+1');
    assert.equal(w.document.querySelector('.badge.eq').textContent, `${age}세`);
    assert.equal(run('luckAt(computed(people[0]),2027).age'), age);
    assert.match(run('renderMemoView()'), new RegExp(`2027년 ${age}세`));
    const years = JSON.parse(run('JSON.stringify(E.yearlyLuck(computed(people[0]).solar.y,computed(people[0]).luck[0].age))'));
    assert.equal(years[0].year - run('computed(people[0]).solar.y') + 1, years[0].age);
  }
  use(fixture({y:1983,mo:12,d:12,h:9,min:0}));
  assert.equal(run('E.STEMS[computed(people[0]).pillars.hour.stem]+E.BRANCHES[computed(people[0]).pillars.hour.branch]'), '戊辰');
  run('settings.correction=false');
  assert.equal(run('E.STEMS[computed(people[0]).pillars.hour.stem]+E.BRANCHES[computed(people[0]).pillars.hour.branch]'), '己巳');
  run('settings.correction=true;goYear(2027);showCalculationBasis()');
  assert.match(w.document.querySelector('#infoBody').textContent, /08:30/);

  for(let month=1; month<=12; month++) {
    const period=JSON.parse(run(`JSON.stringify(monthPeriod(2027,${month}))`));
    assert.ok(period.start && period.end);
    const start=Date.UTC(period.start.year,period.start.m-1,period.start.d,period.start.h-9,period.start.min);
    const end=Date.UTC(period.end.year,period.end.m-1,period.end.d,period.end.h-9,period.end.min);
    assert.equal(run(`isCurrentMonthPeriod(2027,${month},new Date(${start}))`),true);
    assert.equal(run(`isCurrentMonthPeriod(2027,${month},new Date(${end}))`),false);
    assert.equal(run(`isCurrentMonthPeriod(2027,${month},new Date(${start-1}))`),false);
  }
  assert.equal(run('monthPeriod(2027,12).end.year'),2028);
  run('closeInfo();render();');
  w.document.querySelector('[aria-label="2027년 8월 상세"]').click();
  assert.match(w.document.querySelector('#infoBody').textContent,/戊申/);
  assert.match(w.document.querySelector('#infoBody').textContent,/KST/);
  run('closeInfo();setChartMode("simple")');
  assert.ok(w.document.querySelector('.card-pillars.is-simple'));
  assert.equal(JSON.parse(w.localStorage.getItem('ug_settings')).chartMode,'simple');

  run('showReportOptions()');
  assert.equal(w.document.querySelector('#report-memo').checked,false);
  assert.equal(w.document.querySelector('#report-reading').checked,false);
  const publicReport=run('reportHtml(people[0],2027,{luck:true,private:true})');
  assert.doesNotMatch(publicReport,/검토용|1983|09:00|08:30|비공개 메모 검증/);
  assert.match(publicReport,/2027年|2027년/);
  assert.match(publicReport,/2028\.01/);
  assert.match(run('reportHtml(people[0],2027,{memo:true})'),/비공개 메모 검증/);
  run('people[0].aiReadings={2027:{text:"저장된 풀이",at:Date.now(),basis:calculationSummary(people[0],computed(people[0]))}}');
  assert.doesNotMatch(run('reportHtml(people[0],2027,{private:true,reading:true})'),/1983\.12\.12|09:00|08:30/);
  run('printReport()');
  assert.ok(w.document.querySelector('#print-report'));
  assert.doesNotMatch(w.document.querySelector('#print-report').textContent,/비공개 메모 검증/);
  w.dispatchEvent(new w.Event('afterprint'));
  assert.equal(w.document.querySelector('#print-report'),null);

  // 공유 링크는 열기만 해서는 사람을 저장하지 않는다.
  run('window.testLink=buildPersonLink({...people[0],name:"공유검증"});');
  w.location.hash = new URL(w.testLink).hash;
  const before=run('people.length');
  assert.equal(run('checkIncomingLink()'),true);
  assert.equal(run('people.length'),before);
  assert.match(w.document.querySelector('#infoBody').textContent,/내 목록에 추가/);

  // 공유 취소·다운로드 준비 실패를 백업 성공으로 표시하지 않는다.
  run('store.set("ug_lastFileBackup",0);');
  Object.defineProperty(w.navigator,'canShare',{value:()=>true,configurable:true});
  Object.defineProperty(w.navigator,'share',{value:async()=>{throw new w.DOMException('취소','AbortError');},configurable:true});
  await run('shareExport()');
  assert.equal(run('store.get("ug_lastFileBackup",0)'),0);
  w.URL.createObjectURL = () => {throw new Error('저장 실패');};
  await run('exportData()');
  assert.equal(run('store.get("ug_lastFileBackup",0)'),0);
  assert.match(run('buildAiPrompt(people[0],2027)'),/12월/);
  assert.match(run('buildAiPrompt(people[0],2027)'),/원국 네 기둥/);

  // 전문가 전환은 명식 데이터와 조회 중인 연도·대운을 그대로 사용한다.
  run('closeInfo();goYear(2027)');
  const peopleBeforeMode=run('JSON.stringify(people)');
  const selectionBeforeMode=run('JSON.stringify([selLuck,selYear])');
  run('setDisplayMode("expert")');
  assert.equal(run('JSON.stringify(people)'),peopleBeforeMode);
  assert.equal(run('JSON.stringify([selLuck,selYear])'),selectionBeforeMode);
  assert.equal(w.document.querySelector('#main').dataset.mode,'expert');
  assert.equal(JSON.parse(w.localStorage.getItem('ug_settings')).displayMode,'expert');
  assert.equal(new URL(w.location.href).searchParams.get('mode'),'expert');
  assert.ok(w.document.querySelector('.x-chart'));
  assert.equal(w.document.querySelectorAll('.x-chart tbody tr').length,6);
  const expertStems=[...w.document.querySelectorAll('.x-chart tbody tr:nth-child(2) td')].map(el=>el.textContent);
  const expertBranches=[...w.document.querySelectorAll('.x-chart tbody tr:nth-child(3) td')].map(el=>el.textContent);
  assert.deepEqual(expertStems,['丁','己','戊','甲','甲','癸']);
  assert.deepEqual(expertBranches,['未','未','辰','戌','子','亥']);
  w.document.querySelector('[aria-label="2027년 8월 상세"]').click();
  assert.match(w.document.querySelector('#infoBody').textContent,/戊申/);
  run('closeInfo()');
  w.document.querySelector('[aria-label="다음 연도"]').click();
  assert.equal(run('window._homeYearShown'),2028);
  w.document.querySelector('[aria-label="51세 대운 戊午"]').click();
  assert.equal(run('window._homeYearShown'),2033);
  run('goYear(2027);showReportOptions();printReport()');
  assert.match(w.document.querySelector('#print-report').textContent,/2027년 명식 보고서/);
  assert.match(w.document.querySelector('#print-report').textContent,/戊辰/);
  run('closeInfo();setDisplayMode("standard")');
  assert.ok(w.document.querySelector('.card-pillars.is-simple'));
  assert.equal(w.document.querySelector('#main').dataset.mode,'standard');
  assert.equal(run('JSON.stringify(people)'),peopleBeforeMode);
  assert.equal(w.document.querySelector('.badge.eq').textContent,'45세');
  use(fixture({noTime:true,h:null,min:null}));
  run('setDisplayMode("expert")');
  assert.equal(w.document.querySelector('.x-chart tbody tr:nth-child(2) td:nth-of-type(3)').textContent,'—');
  run('people=[];currentId=null;render()');
  assert.match(w.document.querySelector('#main').textContent,/첫 명식 등록/);
  assert.deepEqual(errors,[]);
  console.log('PASS: existing regression suite; expert/standard switch preserves people and selection, six pillars match, year/luck/month controls, PDF, unknown time and empty state');
  w.close();
})().catch(error=>{console.error(error);w.close();process.exitCode=1;});
