/* OneDrive app-folder sync. Authentication is handled by Microsoft's pinned MSAL build. */
(function(){
  'use strict';
  const C=window.SajuSyncCore, GRAPH='https://graph.microsoft.com/v1.0', SCOPES=['Files.ReadWrite.AppFolder'];
  const OD={account:null,client:null,pass:'',binding:null,folder:null,state:C.empty(),base:null,etags:{},busy:false,active:false,timer:null,dirty:false,retryAt:0,last:0,message:'원드라이브 연결 전',pending:false,epoch:0};
  const configId=()=>String(window.UNMYEONG_ONEDRIVE_CLIENT_ID||store.get('ug_od_client','')||'').trim().toLowerCase();
  const validId=id=>/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id);
  const redirectUri=()=>new URL('onedrive-return.html',location.href).href;
  const clone=x=>JSON.parse(JSON.stringify(x));
  let authPromise=null;
  function status(message){OD.message=message;document.querySelectorAll('[data-od-status]').forEach(e=>{e.textContent=message;});}
  function busyEditor(){const e=document.activeElement;return !!(e&&e.matches('input,textarea,[contenteditable=true]'))||!!document.querySelector('#modalBg.show,#cmpBg.show,#nowBg.show,#lockScreen.show');}
  function available(){if(!navigator.locks||!crypto.subtle)throw Error('최신 Chrome·Safari에서 다시 열어 주세요');if(settings.lock&&!sessionPw)throw Error('먼저 앱 잠금을 해제해 주세요');}
  function deviceId(){let id=store.get('ug_od_device','');if(!validId(id)){id=crypto.randomUUID();store.set('ug_od_device',id);}return id;}
  async function auth(){
    if(authPromise)return authPromise;
    authPromise=(async()=>{
      if(!validId(configId()))throw Error('최초 Microsoft 앱 등록이 필요합니다');
      if(!window.msal)throw Error('Microsoft 로그인 모듈을 불러오지 못했습니다');
      const client=new msal.PublicClientApplication({auth:{clientId:configId(),authority:'https://login.microsoftonline.com/common',redirectUri:redirectUri(),navigateToLoginRequestUrl:true},cache:{cacheLocation:'sessionStorage'}});
      await client.initialize();const response=await client.handleRedirectPromise();
      OD.client=client;OD.account=response?.account||client.getActiveAccount();
      if(OD.account)client.setActiveAccount(OD.account);
      if(response){view='settings';render();}
      return client;
    })().catch(e=>{authPromise=null;throw e;});
    return authPromise;
  }
  async function token(){
    const client=await auth();if(!OD.account)throw Error('Microsoft 로그인이 필요합니다');
    try{return (await client.acquireTokenSilent({account:OD.account,scopes:SCOPES,redirectUri:redirectUri()})).accessToken;}
    catch(e){if(e instanceof msal.InteractionRequiredAuthError||e.errorCode==='no_account_error'){OD.active=false;throw Error('로그인 기간이 끝났습니다. Microsoft에 다시 로그인해 주세요');}throw Error('Microsoft 로그인 연결을 확인해 주세요');}
  }
  async function graph(path,options={}){
    const url=path.startsWith('/')?GRAPH+path:path;
    if(!url.startsWith(GRAPH+'/'))throw Error('허용하지 않은 원드라이브 주소입니다');
    const access=await token();
    const r=await fetch(url,{...options,headers:{...options.headers,Authorization:'Bearer '+access},redirect:'error',signal:AbortSignal.timeout(25000)});
    if(r.status===429||r.status===503){const delay=Math.max(10,Number(r.headers.get('Retry-After'))||30);OD.retryAt=Date.now()+delay*1000;throw Error('원드라이브 요청이 많아 잠시 후 다시 시도합니다');}
    if(r.status===401){OD.active=false;throw Error('Microsoft에 다시 로그인해 주세요');}
    if(!r.ok)throw Error(r.status===403?'원드라이브 앱 폴더 권한을 확인해 주세요':r.status===507?'원드라이브 용량이 부족합니다':'원드라이브 응답 오류 ('+r.status+')');
    return r.json();
  }
  async function seal(value){
    const salt=b64(crypto.getRandomValues(new Uint8Array(16)));
    return {app:'unmyeong-sync',v:1,salt,...await encryptData(value,OD.pass,salt)};
  }
  async function unseal(payload){
    if(payload?.app!=='unmyeong-sync'||payload.v!==1||!payload.salt||!payload.iv||!payload.ct||payload.ct.length>16000000)throw Error('동기화 파일 형식을 확인해 주세요');
    try{return await decryptData(payload,OD.pass,payload.salt);}catch(e){throw Error('동기화 암호가 다릅니다. 두 기기에 같은 암호를 입력해 주세요');}
  }
  function cacheKey(){return 'onedrive-sync:'+OD.binding;}
  async function loadCache(){const encrypted=await idbGet(cacheKey());return encrypted?await unseal(encrypted):null;}
  async function cache(){
    const payload=await seal({state:OD.state,base:OD.base,etags:OD.etags});
    if(!await idbSet(cacheKey(),payload))throw Error('동기화 대기 기록을 이 기기에 저장하지 못했습니다');
  }
  function captureLocal(){
    // base holds only versions shown to this tab; remote changes do not acknowledge themselves.
    const result=C.capture(OD.state,OD.base||{},people,deviceId());
    OD.state=result.state;OD.base=result.base;if(result.changed)OD.dirty=true;
  }
  async function readCloud(){
    let next='/me/drive/items/'+encodeURIComponent(OD.folder)+'/children?$top=200';const updates=[];
    while(next){
      const page=await graph(next);
      for(const item of page.value||[]){
        if(!/^device-[0-9a-f-]{36}\.json$/i.test(item.name||''))continue;
        if(item.size>12000000)throw Error('동기화 파일이 너무 큽니다. 백업 후 확인해 주세요');
        if(OD.etags[item.id]===item.eTag)continue;
        const meta=await graph('/me/drive/items/'+encodeURIComponent(item.id)+'?$select=id,eTag,size,@microsoft.graph.downloadUrl');
        const download=meta['@microsoft.graph.downloadUrl'];
        if(!download||new URL(download).protocol!=='https:')throw Error('안전한 다운로드 주소를 받지 못했습니다');
        // Preauthenticated download URL must NOT receive the Graph bearer token.
        const r=await fetch(download,{credentials:'omit',referrerPolicy:'no-referrer',signal:AbortSignal.timeout(25000)});
        if(!r.ok)throw Error('다른 기기의 기록을 읽지 못했습니다');
        const text=await r.text();if(text.length>12000000)throw Error('동기화 파일 크기 초과');
        const state=await unseal(JSON.parse(text));C.validate(state);updates.push({state,id:item.id,etag:meta.eTag});
      }
      next=page['@odata.nextLink']||null;
    }
    // Commit only after every changed file has been decoded successfully.
    for(const u of updates){OD.state=C.merge(OD.state,u.state);OD.etags[u.id]=u.etag;}
  }
  async function applyCloud(){
    if(busyEditor()){OD.pending=true;return false;}
    const projection=C.projected(OD.state),old=new Map(people.map(p=>[p.id,p]));
    const disputed=new Set(C.conflicts(OD.state).map(c=>c.pid));
    const next=projection.people.map(p=>({...((disputed.has(p.id)&&old.has(p.id))?old.get(p.id):p),...(old.get(p.id)?.recs?{recs:old.get(p.id).recs}:{}),...(old.get(p.id)?.viewedAt?{viewedAt:old.get(p.id).viewedAt}:{})}));
    next.push(...people.filter(p=>p._temp||p.id==='sample1'));
    for(const pid of disputed)if(old.has(pid)&&OD.base?.[pid])projection.base[pid]=OD.base[pid];
    OD.base=projection.base;OD.pending=false;
    if(C.stable(C.personMap(next))!==C.stable(C.personMap(people))){
      people=next;OD.applying=true;
      try{if(!await savePeople())throw Error('받은 기록을 기기에 저장하지 못했습니다');}finally{OD.applying=false;}
      render();
    }
    return true;
  }
  function schedule(delay=8000){clearTimeout(OD.timer);if(OD.active)OD.timer=setTimeout(cycle,Math.max(delay,OD.retryAt-Date.now()));}
  async function cycle(){
    if(!OD.active||OD.busy||document.visibilityState==='hidden')return;
    if(Date.now()<OD.retryAt){schedule();return;}
    OD.busy=true;const epoch=OD.epoch;
    try{
      available();
      await navigator.locks.request('unmyeong-onedrive-sync',async()=>{
        if(epoch!==OD.epoch||!OD.active)return;
        const local=await loadCache();
        if(local){OD.state=C.merge(OD.state,local.state);if(OD.base===null)OD.base=local.base||{};OD.etags={...local.etags,...OD.etags};}
        if(OD.initializing){
          // A settings tab may have been left open while another tab saved new records.
          // Read the latest durable local copy under the same lock before the first capture.
          if(typeof _persistQueue!=='undefined')await _persistQueue;
          let saved=store.get('ug_people',null);
          if(settings.lock&&sessionPw){const encrypted=store.get('ug_people_enc',null);if(encrypted?.ct)saved=(await decryptData(encrypted,sessionPw,encrypted.salt)).people;}
          if(Array.isArray(saved))people=saved.concat(people.filter(p=>p._temp));
          OD.base=local?.base||{};OD.initializing=false;
        }
        captureLocal();await cache();
        await readCloud();
        if(epoch!==OD.epoch||!OD.active)return;
        captureLocal();await cache();
        const before=C.stable(OD.state);
        // Each browser installation owns a separate file. Web Locks serialize its tabs.
        const blob=await seal(OD.state);const body=JSON.stringify(blob);
        if(body.length>12000000)throw Error('동기화 용량이 커졌습니다. 백업 후 확인이 필요합니다');
        if(OD.dirty||before!==OD.uploaded){
          const item=await graph('/me/drive/items/'+encodeURIComponent(OD.folder)+':/device-'+deviceId()+'.json:/content',{method:'PUT',headers:{'Content-Type':'application/json'},body});
          OD.uploaded=before;OD.dirty=false;if(item.id&&item.eTag)OD.etags[item.id]=item.eTag;
        }
        if(epoch!==OD.epoch||!OD.active)return;
        // Include keystrokes made while the upload was in flight before projecting remote state.
        captureLocal();await applyCloud();await cache();OD.last=Date.now();
        const count=C.conflicts(OD.state).length;
        status(count?'동시 수정 '+count+'건 · 두 기록 보존됨':OD.pending?'편집을 마치면 다른 기기 기록 반영':OD.dirty?'추가 변경 전송 대기':'원드라이브 동기화됨 · '+new Date(OD.last).toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'}));
      });
    }catch(e){status(navigator.onLine===false?'오프라인 · 기기에 저장 후 연결되면 전송':e.message||'동기화 연결을 확인해 주세요');}
    finally{OD.busy=false;schedule(OD.dirty?2500:8000);}
  }
  async function login(){
    try{
      available();
      const entered=document.getElementById('od-client-id')?.value.trim();
      if(entered){if(!validId(entered))throw Error('Microsoft 앱 ID 형식을 확인해 주세요');if(entered!==configId()){store.set('ug_od_client',entered);authPromise=null;}}
      await flushNotebook();await savePeople();
      const client=await auth();await client.loginRedirect({scopes:SCOPES,prompt:'select_account',redirectUri:redirectUri(),redirectStartPage:new URL('./',location.href).href});
    }catch(e){status(e.message||'로그인을 시작하지 못했습니다');toast(OD.message);}
  }
  async function start(){
    if(OD.busy||OD.starting)return;OD.starting=true;
    try{
      available();await auth();if(!OD.account)throw Error('Microsoft에 먼저 로그인해 주세요');
      const field=document.getElementById('od-pass');const pass=field?.value||'';
      if(pass.length<8)throw Error('두 기기에서 사용할 같은 동기화 암호를 8자 이상 입력해 주세요');
      OD.pass=pass;if(field)field.value='';
      const folder=await graph('/me/drive/special/approot');if(!folder.id||!folder.parentReference?.driveId)throw Error('원드라이브 앱 폴더를 확인하지 못했습니다');
      const binding=configId()+':'+folder.parentReference.driveId+':'+folder.id;
      const previous=store.get('ug_od_binding',null);
      if(previous&&previous!==binding)throw Error('기존 연결과 다른 원드라이브입니다. 기록 보호를 위해 원래 계정으로 로그인해 주세요');
      OD.folder=folder.id;OD.binding=binding;OD.state=C.empty();OD.base=null;OD.etags={};OD.uploaded=null;
      await flushNotebook();OD.initializing=true;
      const local=await loadCache();if(local){C.validate(local.state);OD.state=local.state;OD.base=local.base||{};OD.etags=local.etags||{};}
      // Verify the password against all remote changes before any new cloud write.
      await readCloud();
      OD.active=true;OD.epoch++;store.set('ug_od_binding',binding);store.set('ug_od_linked',true);
      status('연결됨 · 첫 동기화 중');render();await cycle();
    }catch(e){OD.active=false;OD.pass='';status(e.message||'연결하지 못했습니다');toast(OD.message);refreshSettings();}
    finally{OD.starting=false;}
  }
  async function pause(){
    OD.active=false;OD.epoch++;clearTimeout(OD.timer);
    // An already sent request may finish; its completion cannot apply to the local UI.
    OD.pass='';status('동기화 일시 정지 · 기록은 기기에 저장');refreshSettings();
  }
  async function resolve(pid,versionId,all){
    if(!OD.active||OD.busy){toast('동기화를 연결하고 잠시 뒤 다시 눌러 주세요');return;}
    try{
      await navigator.locks.request('unmyeong-onedrive-sync',async()=>{
        const cacheValue=await loadCache();if(cacheValue)OD.state=C.merge(OD.state,cacheValue.state);
        captureLocal();const versions=OD.state.records[pid]?.versions,index=versions?.findIndex(v=>v.actor+':'+v.seq===versionId);if(index==null||index<0)throw Error('기록이 바뀌었습니다. 다시 확인해 주세요');
        if(all){
          const copies=versions.filter((v,i)=>i!==index&&v.value!==null).map(v=>({...clone(v.value),id:'p'+crypto.randomUUID(),name:v.value.name+' (동시 수정본)'}));
          const added=C.capture(OD.state,{},copies,deviceId());OD.state=added.state;
        }
        OD.state=C.resolve(OD.state,pid,index,deviceId());OD.dirty=true;await applyCloud();await cache();
      });closeInfo();schedule(0);render();
    }catch(e){toast(e.message);}
  }
  function conflictView(){
    const items=C.conflicts(OD.state);
    const text=p=>!p?'이 명식 삭제':p.name+' · '+p.y+'/'+p.mo+'/'+p.d+'\n\n일반 메모\n'+(p.memo||'')+'\n\n연도 기록\n'+(p.yearMemos||[]).map(n=>n.year+'년\n'+n.text).join('\n\n')+'\n\n강의 기록\n'+(p.studyNotes||[]).map(n=>(n.title||'제목 없음')+'\n'+n.text).join('\n\n')+'\n\nGPT 풀이\n'+Object.entries(p.aiReadings||{}).map(([y,n])=>y+'년\n'+n.text).join('\n\n');
    _openInfo('<h2>동시 수정 기록</h2>','<p class="section-hint">같은 명식을 두 기기에서 수정한 경우입니다. 선택하기 전까지 양쪽 기록을 보존합니다.</p>'+(items.map(({pid,versions})=>'<div class="card">'+versions.map((v,i)=>'<details class="memo-review-item"><summary>기록 '+(i+1)+' · '+esc(v.value?.name||'삭제 기록')+'</summary><div class="rd-text">'+esc(text(v.value))+'</div><button class="btn-primary" onclick="ugCloud.resolve('+esc(JSON.stringify(pid))+','+esc(JSON.stringify(v.actor+':'+v.seq))+',true)">이 기록 사용 · 다른 기록도 사본 보관</button></details>').join('')+'</div>').join('')||'<p>확인할 동시 수정 기록이 없습니다.</p>'));
  }
  function settingsHtml(){
    const id=configId(),linked=!!OD.account;
    return '<div class="card"><div class="memo-heading">원드라이브 · 휴대폰/패드 공유</div><p class="section-hint" data-od-status>'+esc(OD.message)+'</p>'
      +'<p class="section-hint">두 기기에서 같은 Microsoft 계정과 같은 동기화 암호를 사용하세요. 화면이 열려 있을 때 약 8초마다 확인합니다. 명식·메모·저장 풀이를 공유하며 녹음 원본은 각 기기에 남습니다.</p>'
      +(linked?'<p class="section-hint">로그인: '+esc(OD.account.username||'Microsoft 계정')+'</p>':'')
      +(!window.UNMYEONG_ONEDRIVE_CLIENT_ID?'<details '+(!id?'open':'')+'><summary>최초 앱 연결 설정'+(!id?' · 등록 필요':'')+'</summary><p class="section-hint">Microsoft 앱 등록은 한 번 필요합니다. 두 기기에 같은 앱 ID를 입력하세요. 비밀 키는 사용하지 않습니다.</p><input class="nb-title" id="od-client-id" aria-label="Microsoft 앱 ID" autocomplete="off" placeholder="애플리케이션(클라이언트) ID" value="'+esc(id)+'"><p><a href="onedrive-setup.html" target="_blank" rel="noopener">앱 등록 안내 보기 ↗</a></p></details>':'')
      +'<button class="btn-ghost" onclick="ugCloud.login()">'+(linked?'Microsoft 다시 로그인':'Microsoft 로그인')+'</button>'
      +(OD.active?'<div class="recording-actions"><button class="btn-primary" onclick="ugCloud.sync()">지금 동기화</button><button class="btn-ghost" onclick="ugCloud.pause()">일시 정지</button></div>'
        :linked?'<div class="pw-row"><input id="od-pass" type="password" aria-label="동기화 암호" autocomplete="off" placeholder="두 기기에서 같은 암호 · 8자 이상"><button onclick="ugCloud.start()">동기화 시작</button></div><p class="section-hint">암호는 저장하지 않습니다. 앱을 새로 열면 다시 입력합니다. 암호를 잊으면 원드라이브의 암호화 기록을 복구할 수 없습니다.</p>':'')
      +'<button class="btn-ghost" onclick="ugCloud.conflicts()">동시 수정 기록 확인</button><p class="section-hint">원드라이브의 운명공부 전용 폴더만 사용합니다. 기존 백업 파일은 그대로 두며, 이 기능을 연결한 뒤에는 기존 파일 자동 저장은 일시 중지됩니다.</p></div>';
  }
  function refreshSettings(){const e=document.getElementById('cloudSettings');if(e)e.innerHTML=settingsHtml();}
  window.ugCloud={login,start,pause,resolve,conflicts:conflictView,sync:()=>{if(!OD.active){toast('원드라이브를 먼저 연결해 주세요');return;}return cycle();},settingsHtml,status:()=>OD.message,
    changed(){if(OD.applying)return;if(OD.active){OD.dirty=true;status('이 기기 저장됨 · 원드라이브 전송 대기');schedule(Math.min(2500,Math.max(0,8000-(Date.now()-OD.last))));}},
    badge(){return store.get('ug_od_linked',false)?'<button class="save-status" onclick="go(\'settings\')"><strong>원드라이브</strong><span data-od-status>'+esc(OD.message)+'</span></button>':'';}
  };
  window.addEventListener('online',()=>schedule(0));
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible')schedule(0);});
  document.addEventListener('focusout',()=>{if(OD.pending)schedule(0);});
  if(validId(configId()))auth().then(()=>{status(OD.account?'동기화 암호를 입력하면 연결됩니다':'Microsoft 로그인이 필요합니다');refreshSettings();}).catch(e=>{status(e.message);refreshSettings();});
})();
