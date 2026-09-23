const test=require('node:test');
const assert=require('node:assert/strict');
const C=require('../sync-core.js');
const p=(memo='처음',id='p1')=>({id,name:'검증용',y:1980,mo:1,d:1,memo});
const change=(state,people,actor)=>C.capture(state,C.projected(state).base,people,actor).state;

test('sequential phone/tablet edits converge and stale snapshots do not resurrect deletes',()=>{
  const a=change(C.empty(),[p()],'phone');
  const b=change(a,[p('패드 수정')],'tablet');
  assert.equal(C.projected(C.merge(a,b)).people[0].memo,'패드 수정');
  const deleted=change(b,[],'phone');
  assert.equal(C.projected(C.merge(deleted,a)).people.length,0);
  assert.equal(C.projected(C.merge(deleted,b)).people.length,0);
});
test('offline concurrent edits retain both copies regardless of arrival order',()=>{
  const base=change(C.empty(),[p()],'seed');
  const a=change(base,[p('휴대폰 메모')],'phone');
  const b=change(base,[p('패드 강의')],'tablet');
  const merged=C.merge(a,b);
  assert.equal(C.stable(merged),C.stable(C.merge(b,a)));
  assert.equal(C.conflicts(merged)[0].versions.length,2);
  const further=change(merged,[p('현재 사본 이어 쓰기')],'phone');
  assert.equal(C.conflicts(further)[0].versions.length,2,'계속 쓰기만으로 다른 사본을 폐기하지 않는다');
  const chosen=C.resolve(further,'p1',0,'phone');
  assert.equal(C.conflicts(C.merge(chosen,merged)).length,0);
});
test('independent people merge; deletion concurrent with editing remains recoverable',()=>{
  const base=change(C.empty(),[p(),p('둘째','p2')],'seed');
  const a=change(base,[p('수정'),p('둘째','p2')],'phone');
  const b=change(base,[p(),p('수정2','p2')],'tablet');
  assert.equal(C.conflicts(C.merge(a,b)).length,0);
  const deleted=change(base,[p('둘째','p2')],'tablet');
  const merged=C.merge(a,deleted);
  assert.equal(C.conflicts(merged).length,1);
  assert.equal(C.projected(merged).people.find(p=>p.id==='p1').memo,'수정');
});
test('identical imports are not shown as conflicts; common subsequent edits supersede duplicates',()=>{
  const a=change(C.empty(),[p()],'phone'),b=change(C.empty(),[p()],'tablet');
  const same=C.merge(a,b);assert.equal(C.conflicts(same).length,0);
  const updated=change(same,[p('업데이트')],'phone');
  assert.equal(updated.records.p1.versions.length,1);
});
test('incoming unseen data must not be acknowledged by a draft being typed',()=>{
  const initial=change(C.empty(),[p()],'phone'),shown=C.projected(initial).base;
  const remote=change(initial,[p('패드 변경')],'tablet');
  const typed=C.capture(C.merge(initial,remote),shown,[p('아직 쓰는 내용')],'phone').state;
  assert.equal(C.conflicts(typed).length,1);
  assert.deepEqual(new Set(typed.records.p1.versions.map(v=>v.value.memo)),new Set(['패드 변경','아직 쓰는 내용']));
});
test('merge is associative and idempotent, audio/view-only metadata stays local',()=>{
  const base=change(C.empty(),[p()],'seed');
  const [a,b,c]=['a','b','c'].map(actor=>change(base,[p(actor)],actor));
  assert.equal(C.stable(C.merge(C.merge(a,b),c)),C.stable(C.merge(a,C.merge(b,c))));
  assert.equal(C.stable(C.merge(a,a)),C.stable(a));
  const noChange=C.capture(base,C.projected(base).base,[{...p(),viewedAt:100,recs:[{blob:'local'}]}],'phone');
  assert.equal(noChange.changed,false);
  assert.equal(C.projected(change(C.empty(),[{...p(),_temp:true}],'a')).people.length,0);
});
test('rejects malformed remote registers and unsafe IDs',()=>{
  assert.throws(()=>C.validate({schema:9,records:{}}));
  assert.throws(()=>change(C.empty(),[p('x',"bad' onclick='")],'phone'));
  const s=change(C.empty(),[p()],'phone');s.records.p1.versions[0].seq=999;
  assert.throws(()=>C.validate(s));
});
