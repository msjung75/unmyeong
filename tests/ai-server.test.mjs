import {test} from 'node:test';
import assert from 'node:assert/strict';
import {createAiServer} from '../server/ai-server.mjs';
test('authenticated GPT proxy: responses, audio, errors and limits',async()=>{
  const calls=[];let fail=false;
  const env={OPENAI_API_KEY:'test-server-secret',AI_ACCESS_TOKEN:'test-access-'.repeat(4),ALLOWED_ORIGIN:'https://app.test'};
  const server=createAiServer(env,async(url,init)=>{
    calls.push({url,init});
    return {ok:!fail,status:fail?429:200,json:async()=>url.endsWith('responses')?{status:'completed',model:'test-model',output:[{type:'message',content:[{type:'output_text',text:'모의 풀이'}]}]}:{text:'모의 음성'}};
  });
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const base='http://127.0.0.1:'+server.address().port;
  const headers={Origin:'https://app.test',Authorization:'Bearer '+env.AI_ACCESS_TOKEN,'Content-Type':'application/json'};
  const post=(path,body,extra={})=>fetch(base+path,{method:'POST',headers:{...headers,...extra},body});
  try{
    assert.equal((await post('/health','{}',{Authorization:'wrong'})).status,401);
    assert.equal((await post('/health','{}',{Origin:'https://evil.test'})).status,403);
    assert.equal(calls.length,0);
    assert.equal((await post('/health','{}')).status,200);
    const res=await post('/reading',JSON.stringify({prompt:'계산 명식'}));assert.equal(res.status,200);
    assert.equal((await res.json()).text,'모의 풀이');
    const sent=JSON.parse(calls[0].init.body);assert.equal(sent.store,false);assert.equal(sent.input,'계산 명식');
    assert.equal(calls[0].init.headers.Authorization,'Bearer test-server-secret');
    assert.equal((await post('/reading','{bad')).status,400);
    assert.equal((await post('/reading',JSON.stringify({prompt:'x'.repeat(20001)}))).status,400);
    assert.equal((await post('/reading',JSON.stringify({prompt:'x'.repeat(90000)}))).status,413);
    const audio=await post('/transcribe',new Uint8Array([1,2,3]),{'Content-Type':'audio/webm;codecs=opus'});
    assert.equal((await audio.json()).text,'모의 음성');assert.equal(calls[1].init.body.get('language'),'ko');assert.equal(calls[1].init.body.get('file').size,3);
    assert.equal((await post('/transcribe','audio',{'Content-Type':'text/plain'})).status,415);
    fail=true;assert.equal((await post('/reading',JSON.stringify({prompt:'재시도'}))).status,429);
  }finally{await new Promise(resolve=>server.close(resolve));}
});
