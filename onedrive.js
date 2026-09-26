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
  OD.room=store.get('ug_od_room','');
  const filePrefix=()=>OD.room?'link-'+OD.room+'-device-':'device-';
  function status(message){OD.message=message;const dash=document.querySelector('.dash');if(dash&&!dash.querySelector('[data-od-status]')&&window.ugCloud)dash.querySelector('.today-card')?.insertAdjacentHTML('beforebegin',window.ugCloud.badge());document.querySelectorAll('[data-od-status]').forEach(e=>{e.textContent=message;});}
  function busyEditor(){const e=document.activeElement;return !!(e&&e.matches('input,textarea,[contenteditable=true]'))||!!document.querySelector('#modalBg.show,#cmpBg.show,#nowBg.show,#lockScreen.show');}
  function available(){if(!navigator.locks||!crypto.subtle)throw Error('최신 Chrome·Safari에서 다시 열어 주세요');if(settings.lock&&!sessionPw)throw Error('먼저 앱 잠금을 해제해 주세요');}
  function deviceId(){let id=store.get('ug_od_device','');if(!validId(id)){id=crypto.randomUUID();store.set('ug_od_device',id);}return id;}
  function accountHint(){const h=store.get('ug_od_account_hint',null);return h&&h.client===configId()?h:null;}
  function useAccount(client,account){
    OD.account=account;if(!account)return;
    client.setActiveAccount(account);
    store.set('ug_od_account_hint',{client:configId(),id:account.homeAccountId||'',username:account.username||''});
  }
  async function silentSession(client){
    const hint=accountHint();
    if(!hint?.username||!store.get('ug_od_linked',false)||store.get('ug_od_paused',false))return null;
    const response=await client.ssoSilent({scopes:SCOPES,loginHint:hint.username,redirectUri:redirectUri()});
    const a=response?.account;
    if(!a||(hint.id?a.homeAccountId!==hint.id:a.username?.toLowerCase()!==hint.username.toLowerCase()))throw Error('기존 원드라이브 계정으로 다시 연결해 주세요');
    useAccount(client,a);return response;
  }
  async function auth(){
    if(authPromise)return authPromise;
    authPromise=(async()=>{
      if(!validId(configId()))throw Error('최초 Microsoft 앱 등록이 필요합니다');
      if(!window.msal)throw Error('Microsoft 로그인 모듈을 불러오지 못했습니다');
      const client=new msal.PublicClientApplication({auth:{clientId:configId(),authority:'https://login.microsoftonline.com/common',redirectUri:redirectUri(),navigateToLoginRequestUrl:true},cache:{cacheLocation:'localStorage'}});
      await client.initialize();const response=await client.handleRedirectPromise();
      OD.client=client;
      const hint=accountHint(),accounts=client.getAllAccounts?.()||[client.getActiveAccount()].filter(Boolean);
      const cached=hint?accounts.find(a=>hint.id?a.homeAccountId===hint.id:a.username?.toLowerCase()===hint.username.toLowerCase()):(accounts.length===1?accounts[0]:null);
      useAccount(client,response?.account||cached||(!hint?client.getActiveAccount():null));
      if(!OD.account){try{await silentSession(client);}catch(e){/* A blocked silent session never opens an interactive login on startup. */}}
      if(response){view='settings';render();}
      return client;
    })().catch(e=>{authPromise=null;throw e;});
    return authPromise;
  }
  async function token(){
    const client=await auth();if(!OD.account)throw Error('Microsoft 로그인이 필요합니다');
    try{return (await client.acquireTokenSilent({account:OD.account,scopes:SCOPES,redirectUri:redirectUri()})).accessToken;}
    catch(e){if(e instanceof msal.InteractionRequiredAuthError||e.errorCode==='no_account_error'){try{const recovered=await silentSession(client);if(recovered?.accessToken)return recovered.accessToken;}catch(ignore){}OD.active=false;throw Error('로그인 기간이 끝났습니다. Microsoft에 다시 로그인해 주세요');}throw Error('Microsoft 로그인 연결을 확인해 주세요');}
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
  function cacheKey(){return 'onedrive-sync:'+OD.binding+(OD.room?':'+OD.room:'');}
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
    let next='/me/drive/items/'+encodeURIComponent(OD.folder)+'/children?$top=200';const updates=[];let matched=0;
    while(next){
      const page=await graph(next);
      for(const item of page.value||[]){
        if(!new RegExp('^'+filePrefix()+'[0-9a-f-]{36}\\.json$','i').test(item.name||''))continue;matched++;
        if(item.size>12000000)throw Error('동기화 파일이 너무 큽니다. 백업 후 확인해 주세요');
        if(OD.etags[item.id]===item.eTag)continue;
        let meta=await graph('/me/drive/items/'+encodeURIComponent(item.id));
        let download=meta['@microsoft.graph.downloadUrl']||item['@microsoft.graph.downloadUrl'];
        if(!download){
          const extra=await graph('/me/drive/items/'+encodeURIComponent(item.id)+'?$select='+encodeURIComponent('id,eTag,size,@microsoft.graph.downloadUrl'));
          download=extra['@microsoft.graph.downloadUrl'];meta={...meta,...extra};
        }
        if(!download||new URL(download).protocol!=='https:')throw Error('안전한 다운로드 주소를 받지 못했습니다');
        // Preauthenticated download URL must NOT receive the Graph bearer token.
        const r=await fetch(download,{credentials:'omit',referrerPolicy:'no-referrer',signal:AbortSignal.timeout(25000)});
        if(!r.ok)throw Error('다른 기기의 기록을 읽지 못했습니다');
        const text=await r.text();if(text.length>12000000)throw Error('동기화 파일 크기 초과');
        const state=await unseal(JSON.parse(text));C.validate(state);updates.push({state,id:item.id,etag:meta.eTag});
      }
      next=page['@odata.nextLink']||null;
    }
    OD.remoteFiles=matched;
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
          const item=await graph('/me/drive/items/'+encodeURIComponent(OD.folder)+':/'+filePrefix()+deviceId()+'.json:/content',{method:'PUT',headers:{'Content-Type':'application/json'},body});
          OD.uploaded=before;OD.dirty=false;if(item.id&&item.eTag)OD.etags[item.id]=item.eTag;
        }
        if(epoch!==OD.epoch||!OD.active)return;
        // Include keystrokes made while the upload was in flight before projecting remote state.
        captureLocal();await applyCloud();await cache();OD.last=Date.now();
        const count=C.conflicts(OD.state).length;
        status(count?'동시 수정 '+count+'건 · 두 기록 보존됨':OD.pending?'편집을 마치면 다른 기기 기록 반영':OD.dirty?'추가 변경 전송 대기':'명식 '+Object.keys(C.personMap(people)).length+'개 동기화됨 · '+new Date(OD.last).toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'}));
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
      const client=await auth();
      if(OD.account){try{await token();await resume();if(!OD.active)status('로그인 유지 중 · 동기화 암호로 연결해 주세요');refreshSettings();return;}catch(e){/* Explicit reconnect may now use Microsoft's existing session. */}}
      const hint=accountHint();
      await client.loginRedirect({scopes:SCOPES,...(hint?.username?{loginHint:hint.username}:{}),redirectUri:redirectUri(),redirectStartPage:new URL('./',location.href).href});
    }catch(e){status(e.message||'로그인을 시작하지 못했습니다');toast(OD.message);}
  }
  const REMEMBER='onedrive-remembered-device';
  const accountId=()=>OD.account?.homeAccountId||OD.account?.username||'';
  async function remember(){
    if(!OD.active||!OD.pass)throw Error('먼저 동기화를 연결해 주세요');
    const key=await crypto.subtle.generateKey({name:'AES-GCM',length:256},false,['encrypt','decrypt']);
    const iv=crypto.getRandomValues(new Uint8Array(12));
    const ct=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,new TextEncoder().encode(OD.pass));
    if(!await idbSet(REMEMBER,{key,iv,ct,account:accountId(),client:configId(),binding:OD.binding}))throw Error('이 기기에 암호를 기억하지 못했습니다');
    store.set('ug_od_remember',true);store.set('ug_od_paused',false);
  }
  async function rememberCurrent(){try{await remember();status('이 기기에서 기억함 · 다음부터 자동 연결');refreshSettings();}catch(e){toast(e.message);}}
  async function resume(){
    if(OD.active||OD.starting||OD.busy||!store.get('ug_od_remember',false)||store.get('ug_od_paused',false)||(settings.lock&&!sessionPw))return;
    try{
      await auth();if(!OD.account)return;
      const saved=await idbGet(REMEMBER);
      if(!saved||saved.client!==configId()||saved.account!==accountId()||saved.binding!==store.get('ug_od_binding',null)){status('저장된 연결을 확인해 주세요. 암호로 다시 연결할 수 있습니다');return;}
      const pass=new TextDecoder().decode(await crypto.subtle.decrypt({name:'AES-GCM',iv:saved.iv},saved.key,saved.ct));
      await start(pass,true);
    }catch(e){status('자동 연결하지 못했습니다. 암호를 입력해 다시 연결해 주세요');refreshSettings();}
  }
  async function disconnect(){
    await pause();
    store.set('ug_od_remember',false);
    if(!await idbDel(REMEMBER)){status('저장된 암호 삭제 실패 · 다시 연결 해제를 눌러 주세요');return;}
    if(OD.client)await OD.client.clearCache();
    store.set('ug_od_account_hint',null);
    OD.account=null;OD.pass='';
    // Retain the account binding and old-file guard to prevent stale data restoration.
    status('연결 해제됨 · 저장된 동기화 암호 삭제됨');refreshSettings();
  }
  async function start(savedPass=null,automatic=false){
    if(OD.busy||OD.starting)return;OD.starting=true;
    try{
      available();await auth();if(!OD.account)throw Error('Microsoft에 먼저 로그인해 주세요');
      const field=document.getElementById('od-pass');const pass=typeof savedPass==='string'?savedPass:field?.value||'';
      const keep=!!OD.pairing||automatic||!!document.getElementById('od-remember')?.checked;
      if(pass.length<8)throw Error('두 기기에서 사용할 같은 동기화 암호를 8자 이상 입력해 주세요');
      OD.pass=pass;if(field){field.value='';field.blur();}
      const folder=await graph('/me/drive/special/approot');if(!folder.id||!folder.parentReference?.driveId)throw Error('원드라이브 앱 폴더를 확인하지 못했습니다');
      const binding=configId()+':'+folder.parentReference.driveId+':'+folder.id;
      const previous=store.get('ug_od_binding',null);
      if(previous&&previous!==binding)throw Error('기존 연결과 다른 원드라이브입니다. 기록 보호를 위해 원래 계정으로 로그인해 주세요');
      OD.folder=folder.id;OD.binding=binding;OD.state=C.empty();OD.base=null;OD.etags={};OD.uploaded=null;OD.last=0;
      await flushNotebook();OD.initializing=true;
      const local=await loadCache();if(local){C.validate(local.state);OD.state=local.state;OD.base=local.base||{};OD.etags=local.etags||{};}
      // Verify the password against all remote changes before any new cloud write.
      await readCloud();
      if(OD.joining&&!OD.remoteFiles)throw Error('이 연결 코드의 기록이 없습니다. 휴대폰에서 연결 완료 후 코드를 다시 확인해 주세요');
      OD.active=true;OD.epoch++;store.set('ug_od_binding',binding);store.set('ug_od_linked',true);
      store.set('ug_od_paused',false);
      if(keep)await remember();else {store.set('ug_od_remember',false);await idbDel(REMEMBER);}
      status('연결됨 · 첫 동기화 중');render();await cycle();store.set('ug_od_room',OD.room);return true;
    }catch(e){OD.active=false;OD.pass='';status(e.message||'연결하지 못했습니다');if(!automatic)toast(OD.message);refreshSettings();return false;}
    finally{OD.starting=false;}
  }
  function shortLinkCode(){const alphabet='23456789abcdefghjkmnpqrstuvwxyz';let code='';while(code.length<12){for(const n of crypto.getRandomValues(new Uint8Array(24))){if(n<240&&code.length<12)code+=alphabet[n%30];}}return code;}
  async function pair(create){
    if(OD.busy||OD.starting){toast('진행 중인 동기화가 끝난 뒤 눌러 주세요');return;}
    try{
      await auth();if(!OD.account){toast('먼저 Microsoft 계정을 연결해 주세요');return;}
      const raw=create?shortLinkCode():(document.getElementById('od-link-code')?.value||'').replace(/[\s-]/g,'').toLowerCase();
      if(!/^[0-9a-f]{24}$/.test(raw)&&!/^[2-9a-hjkmnp-z]{12}$/.test(raw))throw Error('휴대폰의 12자리 연결 코드를 입력해 주세요. 이전 24자리 코드도 사용할 수 있습니다');
      const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(raw));
      const room=Array.from(new Uint8Array(digest),n=>n.toString(16).padStart(2,'0')).join('').slice(0,24);
      await flushNotebook();if(!await savePeople())throw Error('먼저 이 기기의 기록 저장을 확인해 주세요');
      // Preserve the entire old connection cache and all old cloud files.
      const previousRoom=OD.room;OD.active=false;clearTimeout(OD.timer);OD.epoch++;
      OD.room=room;OD.pairing=true;OD.joining=!create;
      const ok=await start(raw,false);
      if(!ok){OD.room=previousRoom;return;}
      refreshSettings();if(create&&OD.last)showCode();
    }catch(e){toast(e.message);status(e.message);}
    finally{OD.pairing=false;OD.joining=false;}
  }
  function showCode(){
    if(!OD.active||!OD.room||!OD.pass){toast('연결을 완료한 뒤 확인해 주세요');return;}
    const code=OD.pass.length===12?OD.pass.toUpperCase().match(/.{1,4}/g).join('-'):OD.pass.match(/.{1,6}/g).join('-');
    _openInfo('패드 연결 코드','<p>패드에서 같은 Microsoft 계정으로 로그인하고 아래 코드를 한 번 입력하세요.</p><p style="font-size:22px;letter-spacing:1px;word-break:break-all;user-select:all;padding:16px;background:#eef3f7;border-radius:12px">'+esc(code)+'</p><p class="section-hint">이 코드는 기록을 여는 열쇠입니다. 본인 기기에만 입력하세요. 연결 후에는 자동으로 기억합니다.</p>');
  }
  async function pause(){
    OD.active=false;OD.epoch++;clearTimeout(OD.timer);store.set('ug_od_paused',true);
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
      +'<p class="section-hint">휴대폰에서 먼저 연결한 뒤, 패드에서 같은 Microsoft 계정·동기화 암호로 한 번 연결하세요. 화면이 열려 있을 때 약 8초마다 확인합니다. 명식·메모·저장 풀이를 공유하며 녹음 원본은 각 기기에 남습니다.</p>'
      +'<p class="section-hint">이 기기 명식 '+Object.keys(C.personMap(people)).length+'개 · 기본 샘플은 공유하지 않습니다.</p>'
      +(linked?'<p class="section-hint">로그인: '+esc(OD.account.username||'Microsoft 계정')+'</p>':'')
      +(!window.UNMYEONG_ONEDRIVE_CLIENT_ID?'<details '+(!id?'open':'')+'><summary>최초 앱 연결 설정'+(!id?' · 등록 필요':'')+'</summary><p class="section-hint">Microsoft 앱 등록은 한 번 필요합니다. 두 기기에 같은 앱 ID를 입력하세요. 비밀 키는 사용하지 않습니다.</p><input class="nb-title" id="od-client-id" aria-label="Microsoft 앱 ID" autocomplete="off" placeholder="애플리케이션(클라이언트) ID" value="'+esc(id)+'"><p><a href="onedrive-setup.html" target="_blank" rel="noopener">앱 등록 안내 보기 ↗</a></p></details>':'')
      +(!linked?'<button class="btn-primary" onclick="ugCloud.login()">① Microsoft 계정 연결</button>':'')
      +(linked?'<div class="card" style="background:#f4f7fa;margin:12px 0"><b>휴대폰 · 패드 간편 연결</b><p class="section-hint">기록이 많이 있는 휴대폰에서 새 연결을 만든 뒤, 패드에는 그 코드를 입력하세요. 기존 명식·메모와 이전 동기화 파일은 보존됩니다.</p><button class="btn-primary" onclick="ugCloud.createLink()">휴대폰에서 새 연결 만들기</button><div class="pw-row" style="margin-top:12px"><input id="od-link-code" aria-label="휴대폰 연결 코드" autocomplete="off" autocapitalize="off" spellcheck="false" placeholder="예: ABCD-EFGH-JKMP"><button onclick="ugCloud.joinLink()">패드 연결</button></div></div>':'')
      +(OD.active&&OD.room?'<button class="btn-ghost" onclick="ugCloud.showCode()">패드 연결 코드 보기</button>':'')
      +(OD.active?'<div class="recording-actions"><button class="btn-primary" onclick="ugCloud.sync()">지금 동기화</button><button class="btn-ghost" onclick="ugCloud.pause()">일시 정지</button></div>'
        :linked?'<details><summary>이전 암호 연결 사용</summary><div class="pw-row"><input id="od-pass" type="password" aria-label="동기화 암호" autocomplete="off" placeholder="두 기기에서 같은 암호 · 8자 이상"><button onclick="ugCloud.start()">연결하고 기록 불러오기</button></div><label class="section-hint" style="display:flex;gap:8px;align-items:center;margin:12px 0"><input id="od-remember" type="checkbox" checked>이 기기에서 기억하기 · 다음부터 자동 연결</label><p class="section-hint">개인 휴대폰·패드에서 선택하세요. 이 기기를 사용하는 사람은 기록에 접근할 수 있습니다. 선택하지 않으면 새로 열 때 암호를 입력합니다.</p></details>':'')
      +(OD.active&&!store.get('ug_od_remember',false)?'<button class="btn-ghost" onclick="ugCloud.remember()">이 기기에서 기억하기 · 다음부터 자동 연결</button><p class="section-hint">이 기기를 사용하는 사람은 기록에 접근할 수 있습니다.</p>':'')
      +(store.get('ug_od_remember',false)?'<p class="section-hint">이 기기에서 기억함 · 앱을 열면 자동 연결</p>':'')
      +(linked?'<details><summary>연결 관리</summary><button class="btn-ghost" onclick="ugCloud.login()">Microsoft 로그인 다시 확인</button></details>':'')
      +(linked||store.get('ug_od_remember',false)?'<button class="btn-ghost" onclick="ugCloud.disconnect()">연결 해제 · 저장된 암호 삭제</button>':'')
      +'<button class="btn-ghost" onclick="ugCloud.conflicts()">동시 수정 기록 확인</button><p class="section-hint">원드라이브의 운명공부 전용 폴더만 사용합니다. 기존 백업 파일은 그대로 두며, 이 기능을 연결한 뒤에는 기존 파일 자동 저장은 일시 중지됩니다.</p></div>';
  }
  function refreshSettings(){const e=document.getElementById('cloudSettings');if(e)e.innerHTML=settingsHtml();}
  window.ugCloud={createLink:()=>pair(true),joinLink:()=>pair(false),showCode,login,start,pause,resolve,resume,remember:rememberCurrent,disconnect,conflicts:conflictView,sync:()=>{if(!OD.active){toast('원드라이브를 먼저 연결해 주세요');return;}return cycle();},settingsHtml,status:()=>OD.message,
    changed(){if(OD.applying)return;if(OD.active){OD.dirty=true;status('이 기기 저장됨 · 원드라이브 전송 대기');schedule(Math.min(2500,Math.max(0,8000-(Date.now()-OD.last))));}},
    badge(){return '<button class="save-status" onclick="go(\'settings\')"><strong>원드라이브</strong><span data-od-status>'+esc(OD.message)+'</span></button>';}
  };
  const lock=document.getElementById('lockScreen');if(lock)new MutationObserver(()=>resume()).observe(lock,{attributes:true,attributeFilter:['class']});
  window.addEventListener('online',()=>{resume();schedule(0);});
  document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'){resume();schedule(0);}});
  document.addEventListener('focusout',()=>{if(OD.pending)schedule(0);});
  if(validId(configId()))auth().then(()=>{status(OD.account?'동기화 암호를 입력하면 연결됩니다':'Microsoft 로그인이 필요합니다');refreshSettings();return resume();}).catch(e=>{status(e.message);refreshSettings();});
})();
