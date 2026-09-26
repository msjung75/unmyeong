const test=require('node:test'),assert=require('node:assert/strict'),fs=require('node:fs');
const {JSDOM}=require('jsdom');const {webcrypto:crypto}=require('node:crypto');
const C=require('../sync-core.js');
const source=fs.readFileSync(require('node:path').join(__dirname,'../onedrive.js'),'utf8');
const person=(memo='처음')=>({id:'p1',name:'검증용',y:1980,mo:1,d:1,memo});
const clone=x=>JSON.parse(JSON.stringify(x));
const encode=x=>Buffer.from(x).toString('base64');
async function key(pass,salt){return crypto.subtle.importKey('raw',await crypto.subtle.digest('SHA-256',new TextEncoder().encode(pass+salt)),'AES-GCM',false,['encrypt','decrypt']);}
async function encrypt(data,pass,salt){const iv=crypto.getRandomValues(new Uint8Array(12));return {iv:encode(iv),ct:encode(await crypto.subtle.encrypt({name:'AES-GCM',iv},await key(pass,salt),new TextEncoder().encode(JSON.stringify(data))))};}
async function decrypt(data,pass,salt){return JSON.parse(new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:Buffer.from(data.iv,'base64')},await key(pass,salt),Buffer.from(data.ct,'base64'))));}
function cloud(){
  const files=new Map();let serial=0;const requests=[];
  const service={files,requests,offline:false,nextLink:null,drive:'drive1',status:0};
  service.fetch=async(url,opts={})=>{
    requests.push({url,opts});if(service.offline)throw Error('network offline');
    const response=value=>({ok:true,json:async()=>clone(value),text:async()=>JSON.stringify(value),headers:{get:()=>null}});
    if(url.startsWith('https://download.test/')){
      assert.equal(opts.headers,undefined,'download URLs must never receive a bearer token');
      return response(files.get(url.split('/').at(-1)).payload);
    }
    assert.ok(url.startsWith('https://graph.microsoft.com/v1.0/'));
    assert.equal(opts.headers.Authorization,'Bearer test-token');
    if(service.status)return {ok:false,status:service.status,headers:{get:()=> '30'}};
    if(url.endsWith('/special/approot'))return response({id:'folder1',parentReference:{driveId:service.drive}});
    if(url.includes('/children'))return response({value:[...files.values()].map(f=>({id:f.id,name:f.name,eTag:f.eTag,size:JSON.stringify(f.payload).length})),'@odata.nextLink':service.nextLink});
    if(opts.method==='PUT'){
      const name=url.match(/:\/((?:link-[0-9a-f]+-)?device-[^/]+):\/content$/)[1];let f=[...files.values()].find(f=>f.name===name);
      if(!f)f={id:'file'+(++serial),name,eTag:0};f.eTag=String(+f.eTag+1);f.payload=JSON.parse(opts.body);files.set(f.id,f);return response({id:f.id,eTag:f.eTag});
    }
    const id=url.match(/\/items\/([^?]+)/)[1],f=files.get(id);if(!f)throw Error('unknown file '+id);
    return response({id,eTag:f.eTag,...(!service.omitSelectedUrl||!url.includes('$select')?{'@microsoft.graph.downloadUrl':'https://download.test/'+id}:{})});
  };
  return service;
}
async function client(service,initial=[],pass='shared-pass',options={}){
  const dom=new JSDOM('<input id="od-pass" type="password"><textarea id="draft"></textarea><div data-od-status></div>',{url:'https://example.test/unmyeong/',runScripts:'outside-only',pretendToBeVisual:true});
  const w=dom.window,stored=options.stored||new Map(),idb=options.idb||new Map();
  Object.defineProperty(w,'crypto',{value:crypto});
  Object.defineProperty(w.navigator,'locks',{value:{request:async(name,fn)=>fn()}});
  w.setTimeout=()=>1;w.clearTimeout=()=>{};w.fetch=service.fetch;w.AbortSignal=AbortSignal;
  Object.assign(w,{SajuSyncCore:C,UNMYEONG_ONEDRIVE_CLIENT_ID:'11111111-1111-1111-1111-111111111111',settings:{},sessionPw:null,people:clone(initial),store:{get:(k,d)=>stored.has(k)?clone(stored.get(k)):d,set:(k,v)=>stored.set(k,clone(v))},idbGet:async k=>idb.get(k),idbSet:async(k,v)=>{idb.set(k,structuredClone(v));return true;},idbDel:async k=>{idb.delete(k);return true;},TextEncoder,TextDecoder,b64:encode,encryptData:encrypt,decryptData:decrypt,esc:x=>String(x).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('"','&quot;'),render:()=>{},view:'home',toast:x=>{w.lastToast=x;},closeInfo:()=>{},_openInfo:(title,body)=>{w.infoBody=body;},flushNotebook:async()=>{},savePeople:async()=>{w.saved=clone(w.people);w.ugCloud?.changed();return true;}});
  w.msal={PublicClientApplication:class{async initialize(){}async handleRedirectPromise(){return null;}getActiveAccount(){return {username:'same@example.test'};}setActiveAccount(){}async acquireTokenSilent(){return {accessToken:'test-token'};}async clearCache(){w.cacheCleared=true;}async loginRedirect(){w.redirected=true;}},InteractionRequiredAuthError:class extends Error{}};
  if(options.msal)w.msal.PublicClientApplication=options.msal;
  w.eval(source);await Promise.resolve();await Promise.resolve();
  w.document.querySelector('#od-pass').value=pass;await w.ugCloud.start();
  return {w,stored,idb,close:()=>dom.window.close()};
}
test('two device mock transport: encrypted sync, offline edits, preserved conflicts and safe resolution',async()=>{
  const service=cloud(),a=await client(service,[person()]),b=await client(service);
  try{
    assert.equal(b.w.people[0].memo,'처음');
    assert.ok(service.requests.filter(r=>r.opts.method==='PUT').every(r=>!r.opts.body.includes('검증용')));
    a.w.people[0].memo='휴대폰 기록';await a.w.ugCloud.sync();
    b.w.people[0].memo='패드에서 오프라인 작성';service.offline=true;await b.w.ugCloud.sync();
    assert.equal(b.w.people[0].memo,'패드에서 오프라인 작성');service.offline=false;await b.w.ugCloud.sync();await a.w.ugCloud.sync();
    assert.match(a.w.ugCloud.status(),/두 기록 보존/);
    assert.equal(a.w.people[0].memo,'휴대폰 기록');assert.equal(b.w.people[0].memo,'패드에서 오프라인 작성');
    const states=await Promise.all([...service.files.values()].map(f=>decrypt(f.payload,'shared-pass',f.payload.salt)));
    const state=states.reduce((a,b)=>C.merge(a,b),C.empty()),versions=C.conflicts(state)[0].versions;
    const chosen=versions.find(v=>v.value.memo==='패드에서 오프라인 작성');
    await b.w.ugCloud.resolve('p1',chosen.actor+':'+chosen.seq,true);await b.w.ugCloud.sync();await a.w.ugCloud.sync();
    assert.equal(a.w.people.length,2);assert.equal(a.w.people.find(p=>p.id==='p1').memo,'패드에서 오프라인 작성');
    assert.ok(a.w.people.some(p=>p.memo==='휴대폰 기록'));
  }finally{a.close();b.close();}
});
test('wrong encryption password cannot upload or replace either local or cloud data',async()=>{
  const service=cloud(),a=await client(service,[person()]);
  const before=JSON.stringify([...service.files]),puts=service.requests.filter(r=>r.opts.method==='PUT').length;
  const b=await client(service,[person('다른 기기 보존')],'wrong-pass');
  try{assert.match(b.w.ugCloud.status(),/암호가 다릅니다/);assert.equal(b.w.people[0].memo,'다른 기기 보존');assert.equal(JSON.stringify([...service.files]),before);assert.equal(service.requests.filter(r=>r.opts.method==='PUT').length,puts);}finally{a.close();b.close();}
});
test('remote changes defer during active typing, then apply on a later poll',async()=>{
  const service=cloud(),a=await client(service,[person()]),b=await client(service);
  try{
    b.w.document.querySelector('#draft').focus();a.w.people[0].memo='새 내용';await a.w.ugCloud.sync();await b.w.ugCloud.sync();
    assert.equal(b.w.people[0].memo,'처음');assert.match(b.w.ugCloud.status(),/편집을 마치면/);
    b.w.document.querySelector('#draft').blur();await b.w.ugCloud.sync();assert.equal(b.w.people[0].memo,'새 내용');
  }finally{a.close();b.close();}
});
test('untrusted pagination URL and throttling do not leak tokens or overwrite data',async()=>{
  const service=cloud(),a=await client(service,[person()]);
  try{
    const puts=service.requests.filter(r=>r.opts.method==='PUT').length;
    service.nextLink='https://evil.test/next';a.w.people[0].memo='보관';await a.w.ugCloud.sync();
    assert.match(a.w.ugCloud.status(),/허용하지 않은/);assert.equal(service.requests.some(r=>r.url.includes('evil.test')),false);
    assert.equal(service.requests.filter(r=>r.opts.method==='PUT').length,puts);
    service.nextLink=null;service.status=429;await a.w.ugCloud.sync();assert.match(a.w.ugCloud.status(),/잠시 후/);
    const calls=service.requests.length;await a.w.ugCloud.sync();assert.equal(service.requests.length,calls,'Retry-After 전에 재요청하지 않는다');
  }finally{a.close();}
});

test('remember is opt-in, resumes without password entry and disconnect erases credentials only',async()=>{
 const service=cloud(),a=await client(service,[person()]);
 try{
  assert.equal(a.idb.has('onedrive-remembered-device'),false);
  await a.w.ugCloud.remember();
  const stored=a.idb.get('onedrive-remembered-device');
  assert.equal(stored.key.extractable,false);assert.equal(stored.account,'same@example.test');
  await a.w.ugCloud.pause();
  const n=service.requests.length;await a.w.ugCloud.resume();assert.equal(service.requests.length,n,'pause stays paused');
  a.stored.set('ug_od_paused',false);
  assert.equal(a.w.document.querySelector('#od-pass').value,'');
  await a.w.ugCloud.resume();assert.match(a.w.ugCloud.status(),/동기화됨/);
  await a.w.ugCloud.disconnect();
  assert.equal(a.idb.has('onedrive-remembered-device'),false);assert.equal(a.w.cacheCleared,true);
  assert.equal(a.w.people[0].memo,'처음');assert.equal(a.stored.get('ug_od_remember'),false);
 }finally{a.close();}
});

test('restart recovers matching cached account without account picker',async()=>{
 const service=cloud(),a=await client(service,[person()]);
 let redirects=0;
 await a.w.ugCloud.remember();
 const b=await client(service,[person()],'shared-pass',{stored:a.stored,idb:a.idb,msal:class{
  async initialize(){}async handleRedirectPromise(){return null;}
  getActiveAccount(){return null;}getAllAccounts(){return [{username:'same@example.test'}];}setActiveAccount(){}
  async acquireTokenSilent(){return {accessToken:'test-token'};}async loginRedirect(){redirects++;}
 }});
 try{await b.w.ugCloud.login();assert.equal(redirects,0);assert.match(b.w.ugCloud.status(),/동기화됨|전송 대기/);}finally{a.close();b.close();}
});
test('missing browser cache recovers silently and never redirects on startup',async()=>{
 const service=cloud(),a=await client(service,[person()]);await a.w.ugCloud.remember();
 let silent=0,redirects=0;
 const b=await client(service,[person()],'shared-pass',{stored:a.stored,idb:a.idb,msal:class{
  async initialize(){}async handleRedirectPromise(){return null;}getActiveAccount(){return null;}getAllAccounts(){return [];}setActiveAccount(){}
  async ssoSilent(request){silent++;assert.equal(request.loginHint,'same@example.test');return {account:{username:'same@example.test'},accessToken:'test-token'};}
  async acquireTokenSilent(){return {accessToken:'test-token'};}async loginRedirect(){redirects++;}
 }});
 try{assert.equal(silent,1);assert.equal(redirects,0);assert.match(b.w.ugCloud.status(),/동기화됨|전송 대기/);}finally{a.close();b.close();}
});
test('silent SSO cannot switch to a different OneDrive account',async()=>{
 const service=cloud(),a=await client(service,[person()]);await a.w.ugCloud.remember();const calls=service.requests.length;
 const b=await client(service,[person()],'shared-pass',{stored:a.stored,idb:a.idb,msal:class{
  async initialize(){}async handleRedirectPromise(){return null;}getActiveAccount(){return null;}getAllAccounts(){return [];}setActiveAccount(){throw Error('wrong account selected');}
  async ssoSilent(){return {account:{username:'other@example.test'},accessToken:'wrong-token'};}
  async loginRedirect(){throw Error('must not redirect automatically');}
 }});
 try{assert.equal(service.requests.length,calls);assert.match(b.w.ugCloud.status(),/로그인/);}finally{a.close();b.close();}
});

test('personal drive metadata without selected download annotation still imports records',async()=>{
 const service=cloud();service.omitSelectedUrl=true;
 const a=await client(service,[person('휴대폰 메모')]),b=await client(service,[]);
 try{assert.equal(b.w.people[0].memo,'휴대폰 메모');assert.match(b.w.ugCloud.status(),/1개 동기화됨/);}
 finally{a.close();b.close();}
});

test('new link bypasses unreadable old files and joins another device without deleting local records',async()=>{
 const service=cloud(),old=await client(service,[person('옛 기록')],'old-password');
 const a=await client(service,[person('휴대폰 최신')],'different-password');
 const b=await client(service,[{...person('패드 메모'),id:'p2'}],'different-password');
 try{
  const oldFiles=JSON.stringify([...service.files]);
  await a.w.ugCloud.createLink();
  const code=a.w.infoBody.match(/[0-9a-f]{6}(?:-[0-9a-f]{6}){3}/)[0];
  assert.equal(JSON.stringify([...service.files].slice(0,1)),oldFiles);
  const input=b.w.document.createElement('input');input.id='od-link-code';input.value='000000-000000-000000-000000';b.w.document.body.append(input);
  const count=service.files.size;await b.w.ugCloud.joinLink();assert.equal(service.files.size,count);assert.match(b.w.ugCloud.status(),/기록이 없습니다/);
  input.value=code;await b.w.ugCloud.joinLink();await a.w.ugCloud.sync();
  assert.equal(b.w.people.find(p=>p.id==='p1').memo,'휴대폰 최신');
  assert.equal(a.w.people.find(p=>p.id==='p2').memo,'패드 메모');
  assert.equal(a.stored.get('ug_od_room'),b.stored.get('ug_od_room'));
  assert.equal(b.stored.get('ug_od_remember'),true);
 }finally{old.close();a.close();b.close();}
});
