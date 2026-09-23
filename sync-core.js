/* Multi-value registers: an unseen concurrent edit is retained, never last-write-wins. */
(function(root,factory){if(typeof module==='object'&&module.exports)module.exports=factory();else root.SajuSyncCore=factory();})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const copy=x=>JSON.parse(JSON.stringify(x));
  function stable(x){if(x===null||typeof x!=='object')return JSON.stringify(x);if(Array.isArray(x))return '['+x.map(stable).join(',')+']';return '{'+Object.keys(x).sort().map(k=>JSON.stringify(k)+':'+stable(x[k])).join(',')+'}';}
  const empty=()=>({schema:1,records:{}});
  const keyOK=k=>typeof k==='string'&&/^[A-Za-z0-9_.:-]{1,150}$/.test(k)&&!['__proto__','prototype','constructor'].includes(k);
  const dot=v=>v.actor+':'+v.seq;
  function safePerson(p){
    const walk=value=>{if(value&&typeof value==='object')for(const [k,v] of Object.entries(value)){if(['__proto__','constructor','prototype'].includes(k))throw Error('허용하지 않은 기록 속성');walk(v);}};
    walk(p);
    for(const n of [...(p.yearMemos||[]),...(p.studyNotes||[])])if(n.id!=null&&!keyOK(n.id))throw Error('메모 식별자 오류');
    for(const r of p.relations||[])if(r.pid!=null&&!keyOK(r.pid))throw Error('관계 식별자 오류');
  }
  function clockMerge(a,b){const c={...a};for(const k of Object.keys(b))c[k]=Math.max(c[k]||0,b[k]);return c;}
  function validate(s){
    if(!s||s.schema!==1||!s.records||typeof s.records!=='object'||Array.isArray(s.records))throw Error('지원하지 않는 동기화 파일입니다');
    if(Object.keys(s.records).length>10000)throw Error('동기화 기록 수가 너무 많습니다');
    for(const [pid,r] of Object.entries(s.records)){
      if(!keyOK(pid)||!r||!r.clock||!Array.isArray(r.versions)||r.versions.length>100)throw Error('동기화 기록 형식 오류');
      for(const [actor,n] of Object.entries(r.clock))if(!keyOK(actor)||!Number.isSafeInteger(n)||n<1)throw Error('동기화 순서 오류');
      const seen=new Set();
      for(const v of r.versions){
        if(!keyOK(v.actor)||!Number.isSafeInteger(v.seq)||v.seq<1||r.clock[v.actor]<v.seq||!v.context||seen.has(dot(v)))throw Error('동기화 버전 오류');
        seen.add(dot(v));
        for(const [actor,n] of Object.entries(v.context))if(!keyOK(actor)||!Number.isSafeInteger(n)||n<1||!(r.clock[actor]>=n))throw Error('동기화 문맥 오류');
        if(v.value!==null&&(!v.value||v.value.id!==pid||typeof v.value.name!=='string'||!Number.isInteger(+v.value.y)))throw Error('명식 형식 오류');
        if(v.value!==null)safePerson(v.value);
      }
    }
    return s;
  }
  function mergeRegister(a={clock:{},versions:[]},b={clock:{},versions:[]}){
    const av=new Map(a.versions.map(v=>[dot(v),v])),bv=new Map(b.versions.map(v=>[dot(v),v])), versions=[];
    for(const [id,v] of av){
      if(bv.has(id)&&stable(v)!==stable(bv.get(id)))throw Error('같은 동기화 버전의 내용이 다릅니다');
      if(bv.has(id)||(b.clock[v.actor]||0)<v.seq)versions.push(copy(v));
    }
    for(const [id,v] of bv)if(!av.has(id)&&(a.clock[v.actor]||0)<v.seq)versions.push(copy(v));
    versions.sort((a,b)=>dot(a).localeCompare(dot(b)));
    return {clock:clockMerge(a.clock,b.clock),versions};
  }
  function merge(a,b){
    validate(a);validate(b);const s=empty();
    for(const pid of new Set([...Object.keys(a.records),...Object.keys(b.records)]))s.records[pid]=mergeRegister(a.records[pid],b.records[pid]);
    return s;
  }
  function clean(p){const out=copy(p);delete out.recs;delete out.viewedAt;delete out._temp;return out;}
  function personMap(people){const out={};for(const p of people){if(p._temp||p.id==='sample1')continue;if(!keyOK(p.id))throw Error('명식 식별자 오류');out[p.id]=clean(p);}return out;}
  function winner(r){return r.versions.filter(v=>v.value!==null).at(-1)||r.versions.at(-1);}
  function projected(s){
    const people=[],base={};
    for(const [pid,r] of Object.entries(s.records)){
      const v=winner(r);if(!v)continue;if(v.value!==null)people.push(copy(v.value));
      // Editing the shown copy acknowledges only that copy, not unresolved siblings.
      const same=r.versions.filter(n=>stable(n.value)===stable(v.value));
      base[pid]={value:copy(v.value),clock:same.reduce((clock,n)=>clockMerge(clock,{...n.context,[n.actor]:n.seq}),{})};
    }
    return {people,base};
  }
  function write(s,pid,value,context,actor){
    if(!keyOK(pid)||!keyOK(actor))throw Error('동기화 식별자 오류');
    // Reusing a writer's counter would implicitly acknowledge its unseen sibling.
    // Fork only in that case; normal sequential edits keep a bounded writer clock.
    if(s.records[pid]?.versions.some(v=>v.actor===actor&&(context[actor]||0)<v.seq))actor=actor+'_'+crypto.randomUUID();
    const seq=Math.max(s.records[pid]?.clock[actor]||0,context[actor]||0)+1;
    const next={clock:{...context,[actor]:seq},versions:[{actor,seq,context:copy(context),value:copy(value),at:Date.now()}]};
    s.records[pid]=mergeRegister(s.records[pid],next);return {value:copy(value),clock:next.clock};
  }
  function capture(state,base,people,actor){
    const s=copy(state),b=copy(base),map=personMap(people);let changed=false;
    for(const pid of new Set([...Object.keys(b),...Object.keys(map)])){
      const value=map[pid]||null,old=b[pid]?.value??null;
      if(stable(value)===stable(old))continue;
      b[pid]=write(s,pid,value,b[pid]?.clock||{},actor);changed=true;
    }
    return {state:s,base:b,changed};
  }
  function conflicts(s){return Object.entries(s.records).filter(([,r])=>new Set(r.versions.map(v=>stable(v.value))).size>1).map(([pid,r])=>({pid,versions:copy(r.versions)}));}
  function resolve(s,pid,index,actor){const out=copy(s),r=out.records[pid];if(!r||!r.versions[index])throw Error('기록을 다시 확인해 주세요');write(out,pid,r.versions[index].value,r.clock,actor);return out;}
  return {empty,validate,merge,stable,clean,personMap,projected,capture,conflicts,resolve};
});
