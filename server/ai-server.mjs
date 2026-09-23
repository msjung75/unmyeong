import http from 'node:http';
import { timingSafeEqual } from 'node:crypto';
import { pathToFileURL } from 'node:url';

// Deploy behind HTTPS. No API key or audio is written to disk or returned to clients.
export function createAiServer(env=process.env, upstream=fetch) {
  const key=env.OPENAI_API_KEY, token=env.AI_ACCESS_TOKEN;
  if(!key || !token || token.length<32) throw Error('Set OPENAI_API_KEY and a random AI_ACCESS_TOKEN of at least 32 characters');
  const origin=new URL(env.ALLOWED_ORIGIN||'https://msjung75.github.io').origin;
  const model=env.OPENAI_MODEL||'gpt-4.1-mini';
  const speechModel=env.OPENAI_TRANSCRIBE_MODEL||'gpt-4o-mini-transcribe';
  const hourlyLimit=Number(env.AI_HOURLY_LIMIT||30);
  if(!Number.isSafeInteger(hourlyLimit)||hourlyLimit<1)throw Error('Invalid AI_HOURLY_LIMIT');
  let running=0,used=0,windowAt=Date.now();
  const error=(status,message)=>Object.assign(Error(message),{status});
  const authenticated=req=>{const got=Buffer.from(req.headers.authorization||''),want=Buffer.from('Bearer '+token);return got.length===want.length&&timingSafeEqual(got,want);};
  async function read(req,max){
    if(Number(req.headers['content-length'])>max)throw error(413,'파일이 너무 큽니다');
    let size=0;const chunks=[];
    for await(const chunk of req){size+=chunk.length;if(size>max)throw error(413,'파일이 너무 큽니다');chunks.push(chunk);}
    return Buffer.concat(chunks);
  }
  async function openai(path,body,contentType){
    const res=await upstream('https://api.openai.com/v1/'+path,{method:'POST',headers:{Authorization:'Bearer '+key,...(contentType?{'Content-Type':contentType}:{})},body,signal:AbortSignal.timeout(100000),redirect:'error'});
    if(!res.ok)throw error(res.status===429?429:502,res.status===429?'사용량 한도에 도달했습니다. 잠시 후 다시 시도해 주세요':'OpenAI 요청 실패 — 관리자에게 API 연결을 확인해 주세요');
    return res.json();
  }
  const server=http.createServer(async(req,res)=>{
    res.setHeader('Cache-Control','no-store');res.setHeader('Content-Type','application/json; charset=utf-8');res.setHeader('X-Content-Type-Options','nosniff');res.setHeader('Vary','Origin');
    const send=(code,obj)=>{res.writeHead(code);res.end(JSON.stringify(obj));};
    let counted=false;
    try{
      if(req.headers.origin!==origin)throw error(403,'허용되지 않은 앱 주소입니다');
      res.setHeader('Access-Control-Allow-Origin',origin);
      if(req.method==='OPTIONS'){res.setHeader('Access-Control-Allow-Methods','POST');res.setHeader('Access-Control-Allow-Headers','Content-Type, Authorization');res.writeHead(204);res.end();return;}
      if(req.method!=='POST')throw error(405,'지원하지 않는 요청입니다');
      if(!authenticated(req))throw error(401,'서버 접속 코드를 확인해 주세요');
      if(!['/health','/reading','/transcribe'].includes(req.url))throw error(404,'없는 경로입니다');
      if(req.url==='/health'){await read(req,1024);send(200,{ok:true,model,speechModel});return;}
      if(Date.now()-windowAt>=3600000){used=0;windowAt=Date.now();}
      if(used>=hourlyLimit||running>=2)throw error(429,'요청이 많습니다. 잠시 후 다시 시도해 주세요');
      running++;counted=true;
      if(req.url==='/reading'){
        if(!(req.headers['content-type']||'').startsWith('application/json'))throw error(415,'JSON 형식이 필요합니다');
        let payload;try{payload=JSON.parse((await read(req,80000)).toString('utf8'));}catch(e){if(e.status)throw e;throw error(400,'잘못된 요청입니다');}
        if(typeof payload.prompt!=='string'||!payload.prompt.trim()||payload.prompt.length>20000)throw error(400,'명식 내용을 확인해 주세요');
        used++;
        const data=await openai('responses',JSON.stringify({model,store:false,max_output_tokens:4500,instructions:'한국어로 답하세요. 입력에 제공된 계산 명식을 그대로 사용하고 생년월일로 임의 재계산하지 마세요. 전통 명리의 상징 해석과 사실을 구분하세요. 외도, 질병, 사고, 수익 등 미래 사건이나 성격을 확정하지 마세요. 근거가 없는 판단은 만들지 마세요. 이름을 추정하지 마세요.',input:payload.prompt}),'application/json');
        const text=(data.output||[]).filter(x=>x.type==='message').flatMap(x=>x.content||[]).filter(x=>x.type==='output_text').map(x=>x.text).join('\n').trim();
        if(data.status!=='completed'||!text)throw error(502,'풀이가 완성되지 않았습니다. 기존 저장 내용은 유지됩니다');
        send(200,{text,model:data.model||model});
      }else{
        const type=(req.headers['content-type']||'').split(';')[0];
        const ext={'audio/webm':'webm','audio/mp4':'m4a','audio/ogg':'ogg','audio/mpeg':'mp3','audio/wav':'wav'}[type];
        if(!ext)throw error(415,'지원하지 않는 녹음 형식입니다');
        const audio=await read(req,24*1024*1024);if(!audio.length)throw error(400,'녹음 파일이 비어 있습니다');
        const form=new FormData();form.set('file',new Blob([audio],{type}),'recording.'+ext);form.set('model',speechModel);form.set('language','ko');form.set('response_format','json');
        form.set('prompt','한국어 상담 녹음. 명리 용어: 사주, 명식, 천간, 지지, 갑을병정무기경신임계, 자축인묘진사오미신유술해, 대운, 세운, 월운, 십성, 비견, 겁재, 식신, 상관, 편재, 정재, 편관, 정관, 편인, 정인.');
        used++;const data=await openai('audio/transcriptions',form);
        if(typeof data.text!=='string'||!data.text.trim())throw error(502,'인식된 음성이 없습니다');
        send(200,{text:data.text,model:speechModel});
      }
    }catch(e){if(!res.writableEnded)send(e.status||502,{error:e.status?e.message:'연결이 지연되거나 응답을 받지 못했습니다. 다시 시도해 주세요'});}
    finally{if(counted)running--;}
  });
  server.requestTimeout=120000;server.headersTimeout=15000;
  return server;
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href){
  createAiServer().listen(Number(process.env.PORT||3000),'0.0.0.0',()=>console.log('GPT server listening'));
}
