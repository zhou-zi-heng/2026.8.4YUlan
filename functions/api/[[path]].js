/* ===== 飞凡AI - 后端 (v3.5.0 批次5：Key专人专用) ===== */

/* ---------- Web Crypto ---------- */
async function sha256(t){const d=new TextEncoder().encode(t);const b=await crypto.subtle.digest('SHA-256',d);return [...new Uint8Array(b)].map(x=>x.toString(16).padStart(2,'0')).join('');}
function b64urlEnc(s){return btoa(unescape(encodeURIComponent(s))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');}
function b64urlDec(s){s=s.replace(/-/g,'+').replace(/_/g,'/');while(s.length%4)s+='=';return decodeURIComponent(escape(atob(s)));}
async function hmacSign(m,secret){const k=await crypto.subtle.importKey('raw',new TextEncoder().encode(secret),{name:'HMAC',hash:'SHA-256'},false,['sign']);const sig=await crypto.subtle.sign('HMAC',k,new TextEncoder().encode(m));return btoa(String.fromCharCode(...new Uint8Array(sig))).replace(/\+/g,'-').replace(/\//g,'_').replace(/=+$/,'');}
async function signJWT(payload,secret){const now=Math.floor(Date.now()/1000);const body=Object.assign({},payload,{iat:now,exp:now+5*24*3600});const h=b64urlEnc(JSON.stringify({alg:'HS256',typ:'JWT'}));const p=b64urlEnc(JSON.stringify(body));const sig=await hmacSign(h+'.'+p,secret);return h+'.'+p+'.'+sig;}
async function verifyJWT(token,secret){try{const parts=token.split('.');if(parts.length!==3)return null;const es=await hmacSign(parts[0]+'.'+parts[1],secret);if(es!==parts[2])return null;const pl=JSON.parse(b64urlDec(parts[1]));if(pl.exp&&Math.floor(Date.now()/1000)>pl.exp)return null;return pl;}catch(e){return null;}}
async function verifyUser(request,env){const a=request.headers.get('X-Auth-Token')||'';if(!a)return null;return await verifyJWT(a,env.JWT_SECRET);}
async function verifyAdmin(request,env){const pl=await verifyUser(request,env);if(!pl||pl.role!=='admin')return null;return pl;}

/* ---------- Key 加解密 ---------- */
async function encKey(plain,secret){if(!plain)return '';const enc=new TextEncoder();const km=await crypto.subtle.importKey('raw',enc.encode(secret),{name:'PBKDF2'},false,['deriveKey']);const salt=crypto.getRandomValues(new Uint8Array(16));const iv=crypto.getRandomValues(new Uint8Array(12));const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:50000,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['encrypt']);const cipher=await crypto.subtle.encrypt({name:'AES-GCM',iv},key,enc.encode(plain));const b=(buf)=>btoa(String.fromCharCode(...new Uint8Array(buf)));return 'ENC:'+b(salt)+':'+b(iv)+':'+b(cipher);}
async function decKey(stored,secret){if(!stored)return '';if(stored.indexOf('ENC:')!==0)return stored;try{const p=stored.split(':');const ub=(s)=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));const salt=ub(p[1]),iv=ub(p[2]),cipher=ub(p[3]);const enc=new TextEncoder();const km=await crypto.subtle.importKey('raw',enc.encode(secret),{name:'PBKDF2'},false,['deriveKey']);const key=await crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:50000,hash:'SHA-256'},km,{name:'AES-GCM',length:256},false,['decrypt']);const plain=await crypto.subtle.decrypt({name:'AES-GCM',iv},key,cipher);return new TextDecoder().decode(plain);}catch(e){return '';}}

/* ---------- ★ Key 指纹 / 哈希 ---------- */
/* 指纹：前4后4，明文存库（不足以还原原Key），列表零解密即可显示 */
function keyFp(plain){
    if(!plain)return '';
    const s=String(plain).trim();
    if(!s)return '';
    if(s.length<=8)return s.slice(0,2)+'…'+s.slice(-2);
    return s.slice(0,4)+'…'+s.slice(-4);
}
/* 哈希：sha256 前16位，仅用于"是否重复"比对，不可反推原Key */
async function keyHash(plain){
    if(!plain)return '';
    const s=String(plain).trim();
    if(!s)return '';
    return (await sha256(s)).slice(0,16);
}

function cors(){return{'Access-Control-Allow-Origin':'*','Access-Control-Allow-Methods':'GET, POST, PUT, DELETE, OPTIONS','Access-Control-Allow-Headers':'*'};}
function jr(obj,status){return new Response(JSON.stringify(obj),{status:status||200,headers:Object.assign({'Content-Type':'application/json'},cors())});}

/* ---------- 工具 ---------- */
function pj(txt,fb){
    if(txt===null||txt===undefined||txt==='')return fb;
    try{const v=JSON.parse(txt);return (v===null||v===undefined)?fb:v;}catch(e){return fb;}
}
function normParams(p){
    if(!p||typeof p!=='object')return null;
    return {
        useTemp:!!p.useTemp,
        temperature:(p.temperature!==undefined&&p.temperature!==null&&!isNaN(p.temperature))?Number(p.temperature):0.7,
        useMax:!!p.useMax,
        max_tokens:(p.max_tokens!==undefined&&p.max_tokens!==null&&!isNaN(p.max_tokens))?parseInt(p.max_tokens,10):4096,
        useTopP:!!p.useTopP,
        top_p:(p.top_p!==undefined&&p.top_p!==null&&!isNaN(p.top_p))?Number(p.top_p):1,
        useFreq:!!p.useFreq,
        frequency_penalty:(p.frequency_penalty!==undefined&&p.frequency_penalty!==null&&!isNaN(p.frequency_penalty))?Number(p.frequency_penalty):0,
    };
}
function normSlots(arr){
    if(!Array.isArray(arr))return [];
    return arr.filter(q=>q&&String(q.label||'').trim()).slice(0,12).map(q=>{
        const o={
            label:String(q.label).trim().slice(0,10),
            model:String(q.model||'').trim(),
            engineId:String(q.engineId||'').trim(),
            isImage:!!q.isImage,
            useParams:!!q.useParams,
        };
        if(o.useParams)o.params=normParams(q.params)||normParams({});
        return o;
    });
}

export async function onRequest(context){
    const{request,env}=context;
    if(request.method==='OPTIONS')return new Response(null,{status:204,headers:Object.assign({'Access-Control-Max-Age':'86400'},cors())});
    const url=new URL(request.url);
    const sub=url.pathname.replace(/^\/api\//,'');
    if(sub==='init')return await hInit(request,env,url);
    if(sub==='login')return await hLogin(request,env);
    if(sub==='verify')return await hVerify(request,env);
    if(sub==='engines')return await hEngines(request,env);
    if(sub==='engines/models')return await hEngineModels(request,env,url);
    if(sub==='engines/setmodel')return await hSetModel(request,env);
    if(sub==='presets')return await hGetPresets(request,env);
    if(sub==='log')return await hLog(request,env);
    if(sub==='config')return await hGetConfig(request,env);
    if(sub==='modelprices')return await hGetModelPrices(request,env);
    if(sub.startsWith('admin/')){const ap=await verifyAdmin(request,env);if(!ap)return jr({error:'无管理员权限'},403);return await hAdmin(request,env,sub.replace(/^admin\//,''),ap);}
    return await hProxy(request,env,url,sub);
}

async function hInit(request,env,url){
    if(!env.DB)return jr({error:'D1 未绑定（DB）'},500);
    if(url.searchParams.get('secret')!==env.JWT_SECRET)return jr({error:'初始化密钥错误'},403);
    try{const ex=await env.DB.prepare('SELECT id FROM users WHERE role=?').bind('admin').first();if(ex)return jr({error:'已存在管理员账号'},400);const h=await sha256('admin123');await env.DB.prepare('INSERT INTO users (username,password_hash,name,role,status,permissions,created_at) VALUES (?,?,?,?,?,?,?)').bind('admin',h,'超级管理员','admin','active','{}',Date.now()).run();return jr({ok:true,msg:'✅ 已创建 admin/admin123'});}catch(e){return jr({error:e.message},500);}
}

async function hLogin(request,env){
    if(!env.DB)return jr({error:'D1 未绑定'},500);
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const un=(b.username||'').trim(),pw=(b.password||'').trim();
    if(!un||!pw)return jr({error:'请输入账号密码'},400);
    try{
        const u=await env.DB.prepare('SELECT * FROM users WHERE username=?').bind(un).first();
        if(!u)return jr({error:'账号或密码错误'},401);
        if(u.status!=='active')return jr({error:'该账号已被禁用'},403);
        if(await sha256(pw)!==u.password_hash)return jr({error:'账号或密码错误'},401);
        const token=await signJWT({username:u.username,name:u.name,role:u.role},env.JWT_SECRET);
        const ip=request.headers.get('CF-Connecting-IP')||'';
        try{await env.DB.prepare('INSERT INTO sessions (username,session_id,ip,last_active,login_at) VALUES (?,?,?,?,?)').bind(u.username,token.slice(-16),ip,Date.now(),Date.now()).run();}catch(e){}
        return jr({ok:true,token,user:{
            username:u.username,name:u.name,role:u.role,
            permissions:u.permissions||'{}',
            quickModels:u.quick_models||''
        }});
    }catch(e){return jr({error:e.message},500);}
}

async function hVerify(request,env){
    const a=request.headers.get('X-Auth-Token')||'';
    if(!a)return jr({ok:false},401);
    const pl=await verifyJWT(a,env.JWT_SECRET);
    if(!pl)return jr({ok:false,error:'token无效或过期'},401);
    let perm='{}',role=pl.role,name=pl.name,status='active',qm='';
    if(env.DB){
        try{
            const u=await env.DB.prepare('SELECT name,role,status,permissions,quick_models FROM users WHERE username=?').bind(pl.username).first();
            if(u){perm=u.permissions||'{}';role=u.role;name=u.name;status=u.status;qm=u.quick_models||'';}
            const ip=request.headers.get('CF-Connecting-IP')||'';
            await env.DB.prepare('UPDATE sessions SET last_active=?,ip=? WHERE session_id=?').bind(Date.now(),ip,a.slice(-16)).run();
        }catch(e){
            try{
                const u2=await env.DB.prepare('SELECT name,role,status,permissions FROM users WHERE username=?').bind(pl.username).first();
                if(u2){perm=u2.permissions||'{}';role=u2.role;name=u2.name;status=u2.status;}
            }catch(e2){}
        }
    }
    if(status!=='active')return jr({ok:false,error:'账号已被禁用'},403);
    return jr({ok:true,user:{username:pl.username,name:name,role:role,permissions:perm,quickModels:qm}});
}

/* ==================== 用户：公有引擎（不下发任何Key信息） ==================== */
async function hEngines(request,env){
    const pl=await verifyUser(request,env);
    if(!pl)return jr({error:'未登录'},401);
    try{
        const rows=(await env.DB.prepare(
            'SELECT id,name,protocol,model,user_model,use_cache,engine_type,params,price_in,price_out,price_cache_read,price_cache_write FROM engines_public WHERE username=? ORDER BY name'
        ).bind(pl.username).all()).results||[];
        const engines=rows.map(e=>({
            id:e.id,name:e.name,protocol:e.protocol,
            model:e.user_model||e.model||'',
            useCache:!!e.use_cache,
            engineType:e.engine_type||'chat',
            params:normParams(pj(e.params,null)),
            priceIn:e.price_in,priceOut:e.price_out,
            priceCacheRead:e.price_cache_read,priceCacheWrite:e.price_cache_write,
            origin:'public'
        }));
        return jr({ok:true,engines});
    }catch(e){
        try{
            const rows=(await env.DB.prepare('SELECT id,name,protocol,model,user_model,use_cache,engine_type,price_in,price_out,price_cache_read,price_cache_write FROM engines_public WHERE username=? ORDER BY name').bind(pl.username).all()).results||[];
            const engines=rows.map(e=>({id:e.id,name:e.name,protocol:e.protocol,model:e.user_model||e.model||'',useCache:!!e.use_cache,engineType:e.engine_type||'chat',params:null,priceIn:e.price_in,priceOut:e.price_out,priceCacheRead:e.price_cache_read,priceCacheWrite:e.price_cache_write,origin:'public'}));
            return jr({ok:true,engines});
        }catch(e2){
            try{
                const rows=(await env.DB.prepare('SELECT id,name,protocol,model,price_in,price_out,price_cache_read,price_cache_write FROM engines_public WHERE username=? ORDER BY name').bind(pl.username).all()).results||[];
                const engines=rows.map(e=>({id:e.id,name:e.name,protocol:e.protocol,model:e.model||'',useCache:false,engineType:'chat',params:null,priceIn:e.price_in,priceOut:e.price_out,priceCacheRead:e.price_cache_read,priceCacheWrite:e.price_cache_write,origin:'public'}));
                return jr({ok:true,engines});
            }catch(e3){return jr({error:e3.message},500);}
        }
    }
}

async function hEngineModels(request,env,url){
    const pl=await verifyUser(request,env);if(!pl)return jr({error:'未登录'},401);
    const engId=url.searchParams.get('engineId');if(!engId)return jr({error:'缺engineId'},400);
    try{
        const e=await env.DB.prepare('SELECT * FROM engines_public WHERE id=? AND username=?').bind(engId,pl.username).first();
        if(!e)return jr({error:'引擎不存在'},404);
        const key=await decKey(e.api_key,env.KEY_SECRET);
        if(!key)return jr({error:'该引擎未配置 API Key，请联系管理员'},400);
        let path='models';
        if(e.protocol==='anthropic')path='v1/models';
        if(e.protocol==='gemini')path='v1beta/models';
        const resp=await fetch(e.base_url.replace(/\/+$/,'')+'/'+path,{headers:{'Authorization':'Bearer '+key,'anthropic-version':'2023-06-01'}});
        if(!resp.ok)return jr({error:'HTTP '+resp.status},500);
        const data=await resp.json();
        let list=[];
        if(Array.isArray(data.data))list=data.data.map(m=>m.id||m.name).filter(Boolean);
        else if(Array.isArray(data.models))list=data.models.map(m=>(m.id||m.name||'').replace(/^models\//,'')).filter(Boolean);
        else if(Array.isArray(data))list=data.map(m=>m.id||m.name||m).filter(Boolean);
        return jr({ok:true,models:list.sort()});
    }catch(e){return jr({error:e.message},500);}
}
async function hSetModel(request,env){
    const pl=await verifyUser(request,env);if(!pl)return jr({error:'未登录'},401);
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    if(!b.engineId)return jr({error:'缺engineId'},400);
    try{await env.DB.prepare('UPDATE engines_public SET user_model=? WHERE id=? AND username=?').bind(b.model||'',b.engineId,pl.username).run();return jr({ok:true});}catch(e){return jr({ok:false,error:e.message});}
}
async function hGetPresets(request,env){const pl=await verifyUser(request,env);if(!pl)return jr({error:'未登录'},401);try{const row=await env.DB.prepare('SELECT data FROM presets WHERE id=1').first();if(!row||!row.data)return jr({ok:true,presets:null});return jr({ok:true,presets:JSON.parse(row.data)});}catch(e){return jr({ok:true,presets:null});}}
async function hLog(request,env){const pl=await verifyUser(request,env);if(!pl)return jr({error:'未登录'},401);let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}try{await env.DB.prepare('INSERT INTO logs (username,chat_name,rounds,tokens,model,created_at) VALUES (?,?,?,?,?,?)').bind(pl.username,(b.chatName||'').slice(0,100),b.rounds||0,b.tokens||0,(b.model||'').slice(0,60),Date.now()).run();return jr({ok:true});}catch(e){return jr({ok:false});}}

async function hGetConfig(request,env){
    const pl=await verifyUser(request,env);
    if(!pl)return jr({error:'未登录'},401);
    if(!env.DB)return jr({ok:true,config:{}});
    try{
        const rows=(await env.DB.prepare('SELECT key,value FROM global_config').all()).results||[];
        const raw={};
        rows.forEach(r=>raw[r.key]=r.value);
        const safe={
            chunkSize:   raw.chunkSize   || '300',
            quickModels: raw.quickModels || '[]',
            quickCmds:   raw.quickCmds   || '',
        };
        return jr({ok:true,config:safe});
    }catch(e){return jr({ok:true,config:{}});}
}

async function hGetModelPrices(request,env){const pl=await verifyUser(request,env);if(!pl)return jr({error:'未登录'},401);try{const rows=(await env.DB.prepare('SELECT model_name,price_in,price_out,price_cache_read,price_cache_write FROM model_prices').all()).results||[];const map={};rows.forEach(r=>map[r.model_name]={priceIn:r.price_in,priceOut:r.price_out,priceCacheRead:r.price_cache_read,priceCacheWrite:r.price_cache_write});return jr({ok:true,prices:map});}catch(e){return jr({ok:true,prices:{}});}}

/* ==================== 管理路由 ==================== */
async function hAdmin(request,env,action,payload){
    if(!env.DB)return jr({error:'D1 未绑定'},500);
    if(action==='ping')return jr({ok:true,admin:payload.username});
    if(action==='users/list')return await aUsersList(env);
    if(action==='users/create')return await aUsersCreate(request,env);
    if(action==='users/update')return await aUsersUpdate(request,env);
    if(action==='users/delete')return await aUsersDelete(request,env);
    if(action==='users/resetpwd')return await aUsersResetPwd(request,env);
    if(action==='users/import')return await aUsersImport(request,env);
    if(action==='users/export')return await aUsersExport(request,env);
    if(action==='users/perm')return await aUsersPerm(request,env);
    if(action==='users/slots')return await aUsersSlots(request,env);
    if(action==='engines/list')return await aEnginesList(request,env,new URL(request.url));
    if(action==='engines/save')return await aEnginesSave(request,env);
    if(action==='engines/delete')return await aEnginesDelete(request,env);
    if(action==='engines/health')return await aEnginesHealth(request,env);
    if(action==='engines/keyfp')return await aEnginesKeyFp(request,env);
    if(action==='templates/list')return await aTemplatesList(env);
    if(action==='templates/save')return await aTemplatesSave(request,env);
    if(action==='templates/delete')return await aTemplatesDelete(request,env);
    if(action==='templates/deploy')return await aTemplatesDeploy(request,env);
    if(action==='templates/sync')return await aTemplatesSync(request,env);
    if(action==='slottpl/list')return await aSlotTplList(env);
    if(action==='slottpl/save')return await aSlotTplSave(request,env);
    if(action==='slottpl/delete')return await aSlotTplDelete(request,env);
    if(action==='slottpl/deploy')return await aSlotTplDeploy(request,env);
    if(action==='presets/get')return await aPresetsGet(env);
    if(action==='presets/save')return await aPresetsSave(request,env);
    if(action==='monitor')return await aMonitor(env);
    if(action==='config/get')return await aConfigGet(env);
    if(action==='config/save')return await aConfigSave(request,env);
    if(action==='models/list')return await aModelsList(env);
    if(action==='models/save')return await aModelsSave(request,env);
    if(action==='models/delete')return await aModelsDelete(request,env);
    return jr({error:'未知接口：'+action},404);
}

/* ==================== 账号 ==================== */
async function aUsersList(env){
    try{
        let users;
        try{
            users=(await env.DB.prepare('SELECT id,username,name,role,status,permissions,quick_models,created_at FROM users ORDER BY created_at DESC').all()).results||[];
        }catch(e){
            users=(await env.DB.prepare('SELECT id,username,name,role,status,permissions,created_at FROM users ORDER BY created_at DESC').all()).results||[];
        }
        const engRows=(await env.DB.prepare('SELECT username,COUNT(*) AS cnt FROM engines_public GROUP BY username').all()).results||[];
        const engMap={};engRows.forEach(r=>engMap[r.username]=r.cnt);

        /* 缺 Key 的引擎数（提示超管补齐） */
        let noKeyMap={};
        try{
            const nk=(await env.DB.prepare("SELECT username,COUNT(*) AS n FROM engines_public WHERE api_key IS NULL OR api_key='' GROUP BY username").all()).results||[];
            nk.forEach(r=>noKeyMap[r.username]=r.n);
        }catch(e){}

        const weekAgo=Date.now()-7*24*3600*1000;
        const sess=(await env.DB.prepare('SELECT username,ip,last_active FROM sessions WHERE last_active>?').bind(weekAgo).all()).results||[];
        const sm={};
        sess.forEach(s=>{if(!sm[s.username])sm[s.username]={last:0,ips:{}};if(s.last_active>sm[s.username].last)sm[s.username].last=s.last_active;if(s.ip)sm[s.username].ips[s.ip]=1;});

        const list=users.map(u=>{
            const s=sm[u.username]||{last:0,ips:{}};
            const ipc=Object.keys(s.ips).length;
            const slots=normSlots(pj(u.quick_models,[]));
            return{
                username:u.username,name:u.name,role:u.role,status:u.status,
                permissions:u.permissions||'{}',
                quickModels:u.quick_models||'',
                slotCount:slots.length,
                engineCount:engMap[u.username]||0,
                noKeyCount:noKeyMap[u.username]||0,
                lastActive:s.last,ipCount:ipc,ipAbnormal:ipc>=3
            };
        });
        return jr({ok:true,users:list});
    }catch(e){return jr({error:e.message},500);}
}
async function aUsersCreate(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const un=(b.username||'').trim(),pw=(b.password||'').trim(),nm=(b.name||'').trim(),role=b.role==='admin'?'admin':'user';if(!un||!pw)return jr({error:'账号密码必填'},400);try{if(await env.DB.prepare('SELECT id FROM users WHERE username=?').bind(un).first())return jr({error:'账号已存在'},400);const h=await sha256(pw);await env.DB.prepare('INSERT INTO users (username,password_hash,name,role,status,permissions,created_at) VALUES (?,?,?,?,?,?,?)').bind(un,h,nm,role,'active','{}',Date.now()).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aUsersUpdate(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const un=(b.username||'').trim();if(!un)return jr({error:'缺账号'},400);try{const f=[],v=[];if(b.name!==undefined){f.push('name=?');v.push(b.name);}if(b.role!==undefined){f.push('role=?');v.push(b.role==='admin'?'admin':'user');}if(b.status!==undefined){f.push('status=?');v.push(b.status==='active'?'active':'disabled');}if(!f.length)return jr({error:'无更新'},400);v.push(un);await env.DB.prepare('UPDATE users SET '+f.join(',')+' WHERE username=?').bind(...v).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aUsersDelete(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const un=(b.username||'').trim();if(!un)return jr({error:'缺账号'},400);if(un==='admin')return jr({error:'不能删admin'},400);try{await env.DB.prepare('DELETE FROM users WHERE username=?').bind(un).run();await env.DB.prepare('DELETE FROM engines_public WHERE username=?').bind(un).run();await env.DB.prepare('DELETE FROM sessions WHERE username=?').bind(un).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aUsersResetPwd(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const un=(b.username||'').trim(),pw=(b.password||'').trim();if(!un||!pw)return jr({error:'必填'},400);try{const h=await sha256(pw);await env.DB.prepare('UPDATE users SET password_hash=? WHERE username=?').bind(h,un).run();await env.DB.prepare('DELETE FROM sessions WHERE username=?').bind(un).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aUsersPerm(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const un=(b.username||'').trim();if(!un)return jr({error:'缺账号'},400);try{await env.DB.prepare('UPDATE users SET permissions=? WHERE username=?').bind(JSON.stringify(b.permissions||{}),un).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aUsersSlots(request,env){
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const un=(b.username||'').trim();
    if(!un)return jr({error:'缺账号'},400);
    try{
        const slots=normSlots(b.slots||[]);
        const txt=slots.length?JSON.stringify(slots):'';
        await env.DB.prepare('UPDATE users SET quick_models=? WHERE username=?').bind(txt,un).run();
        return jr({ok:true,count:slots.length});
    }catch(e){return jr({error:'保存失败：'+e.message+'（若提示无 quick_models 列，请先执行建表SQL）'},500);}
}

/* ★ 账号 xlsx 导入：只建/改账号，绝不碰引擎（原来会 DELETE 全部引擎，已移除） */
async function aUsersImport(request,env){
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const rows=b.rows||[];
    if(!rows.length)return jr({error:'无数据'},400);

    let created=0,updated=0;
    const errs=[];
    const seen={};

    for(const r of rows){
        const un=String(r['账号']||r['username']||'').trim();
        if(!un)continue;
        if(seen[un]){errs.push(un+'：表格中重复出现，已跳过后者');continue;}
        seen[un]=1;

        const pw=String(r['密码']||r['password']||'').trim();
        const nm=String(r['姓名']||r['name']||'').trim();
        const roleRaw=String(r['角色']||r['role']||'user').trim().toLowerCase();
        const role=(roleRaw==='admin'||roleRaw==='管理员')?'admin':'user';

        try{
            const ex=await env.DB.prepare('SELECT id FROM users WHERE username=?').bind(un).first();
            if(ex){
                /* 已存在：改姓名/角色；密码留空则不动 */
                if(pw){
                    const h=await sha256(pw);
                    await env.DB.prepare('UPDATE users SET password_hash=?,name=?,role=? WHERE username=?').bind(h,nm,role,un).run();
                }else{
                    await env.DB.prepare('UPDATE users SET name=?,role=? WHERE username=?').bind(nm,role,un).run();
                }
                updated++;
            }else{
                if(!pw){errs.push(un+'：新账号必须填密码，已跳过');continue;}
                const h=await sha256(pw);
                await env.DB.prepare('INSERT INTO users (username,password_hash,name,role,status,permissions,created_at) VALUES (?,?,?,?,?,?,?)').bind(un,h,nm,role,'active','{}',Date.now()).run();
                created++;
            }
        }catch(e){errs.push(un+'：'+e.message);}
    }
    return jr({ok:true,created,updated,errors:errs});
}

/* 导出台账（含 Key 备注） */
async function aUsersExport(request,env){
    const url=new URL(request.url);
    const wk=url.searchParams.get('withkey')==='1';
    try{
        const users=(await env.DB.prepare('SELECT username,name,role,status FROM users ORDER BY created_at').all()).results||[];
        let engs=[];
        try{engs=(await env.DB.prepare('SELECT * FROM engines_public ORDER BY username').all()).results||[];}catch(e){}
        const eb={};
        engs.forEach(e=>{if(!eb[e.username])eb[e.username]=[];eb[e.username].push(e);});

        const out=[];
        for(const u of users){
            const ue=eb[u.username]||[];
            if(!ue.length){
                out.push({姓名:u.name,账号:u.username,角色:u.role,状态:u.status,引擎名称:'',协议:'',BaseURL:'',APIKey:'',Key指纹:'',Key备注:'',模型:''});
            }else{
                for(const e of ue){
                    let ko='******';
                    if(wk)ko=await decKey(e.api_key,env.KEY_SECRET)||'（无Key）';
                    out.push({
                        姓名:u.name,账号:u.username,角色:u.role,状态:u.status,
                        引擎名称:e.name,协议:e.protocol,BaseURL:e.base_url,
                        APIKey:ko,
                        Key指纹:e.key_fp||'',
                        Key备注:e.key_note||'',
                        模型:e.model||''
                    });
                }
            }
        }
        return jr({ok:true,rows:out});
    }catch(e){return jr({error:e.message},500);}
}

/* ==================== 引擎管理（★Key指纹 + 查重） ==================== */
async function aEnginesList(request,env,url){
    const un=url.searchParams.get('username');
    try{
        const cols='id,username,name,protocol,base_url,model,use_cache,engine_type,tpl_id,params,key_note,key_fp,key_hash,api_key,price_in,price_out,price_cache_read,price_cache_write';
        let rows;
        if(un)rows=(await env.DB.prepare('SELECT '+cols+' FROM engines_public WHERE username=? ORDER BY name').bind(un).all()).results||[];
        else rows=(await env.DB.prepare('SELECT '+cols+' FROM engines_public ORDER BY username,name').all()).results||[];

        /* ★ 全局查重：哪些 key_hash 被 2 个以上引擎共用 */
        const dupMap={};
        try{
            const dups=(await env.DB.prepare(
                "SELECT key_hash, COUNT(*) AS n, GROUP_CONCAT(username) AS users FROM engines_public WHERE key_hash IS NOT NULL AND key_hash<>'' GROUP BY key_hash HAVING n>1"
            ).all()).results||[];
            dups.forEach(d=>{
                dupMap[d.key_hash]=String(d.users||'').split(',').filter(Boolean);
            });
        }catch(e){}

        const engs=rows.map(e=>{
            const hasKey=!!(e.api_key&&String(e.api_key).length);
            const dupUsers=(e.key_hash&&dupMap[e.key_hash])?dupMap[e.key_hash].filter(x=>x!==e.username):[];
            return{
                id:e.id,username:e.username,name:e.name,protocol:e.protocol,
                base:e.base_url,model:e.model,
                useCache:!!e.use_cache,
                engineType:e.engine_type||'chat',
                tplId:e.tpl_id||'',
                params:normParams(pj(e.params,null)),
                hasKey:hasKey,
                keyFp:e.key_fp||'',
                keyNote:e.key_note||'',
                keyDup:dupUsers.length>0,
                dupWith:dupUsers,
                needFp:hasKey&&!e.key_fp,     /* 需要补全指纹（历史数据） */
                priceIn:e.price_in,priceOut:e.price_out,
                priceCR:e.price_cache_read,priceCW:e.price_cache_write
            };
        });
        return jr({ok:true,engines:engs});
    }catch(e){
        /* 兜底：key_* 列不存在时降级 */
        try{
            const cols2='id,username,name,protocol,base_url,model,use_cache,engine_type,tpl_id,params,api_key,price_in,price_out,price_cache_read,price_cache_write';
            let rows2;
            if(un)rows2=(await env.DB.prepare('SELECT '+cols2+' FROM engines_public WHERE username=? ORDER BY name').bind(un).all()).results||[];
            else rows2=(await env.DB.prepare('SELECT '+cols2+' FROM engines_public ORDER BY username,name').all()).results||[];
            const engs=rows2.map(e=>({
                id:e.id,username:e.username,name:e.name,protocol:e.protocol,
                base:e.base_url,model:e.model,useCache:!!e.use_cache,
                engineType:e.engine_type||'chat',tplId:e.tpl_id||'',
                params:normParams(pj(e.params,null)),
                hasKey:!!(e.api_key&&String(e.api_key).length),
                keyFp:'',keyNote:'',keyDup:false,dupWith:[],needFp:false,
                priceIn:e.price_in,priceOut:e.price_out,priceCR:e.price_cache_read,priceCW:e.price_cache_write
            }));
            return jr({ok:true,engines:engs,warn:'未检测到 key_fp/key_hash 列，请执行建表SQL以启用Key指纹与查重'});
        }catch(e2){return jr({error:e2.message},500);}
    }
}

async function aEnginesSave(request,env){
    let b;
    try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const un=(b.username||'').trim();
    if(!un)return jr({error:'缺账号'},400);
    if(!b.name)return jr({error:'引擎名必填'},400);
    try{
        const id=b.id||('eng_'+un+'_'+Math.random().toString(36).slice(2,8));
        const ex=b.id?await env.DB.prepare('SELECT api_key,key_fp,key_hash FROM engines_public WHERE id=?').bind(b.id).first():null;

        /* Key：填了才更新；留空保留原值 */
        let keyStored,fp,kh;
        const rawKey=(b.key||'').trim();
        if(rawKey&&rawKey!=='******'){
            keyStored=await encKey(rawKey,env.KEY_SECRET);
            fp=keyFp(rawKey);
            kh=await keyHash(rawKey);
        }else if(ex){
            keyStored=ex.api_key;
            fp=ex.key_fp||'';
            kh=ex.key_hash||'';
        }else{
            keyStored='';fp='';kh='';
        }

        const uc=b.useCache?1:0;
        const etype=(b.engineType==='image')?'image':'chat';
        const tplId=(b.tplId!==undefined)?String(b.tplId||''):'';
        const np=normParams(b.params);
        const paramsTxt=np?JSON.stringify(np):'';
        const note=String(b.keyNote||'').slice(0,80);

        if(ex){
            await env.DB.prepare('UPDATE engines_public SET name=?,protocol=?,base_url=?,api_key=?,model=?,use_cache=?,engine_type=?,tpl_id=?,params=?,key_note=?,key_fp=?,key_hash=?,price_in=?,price_out=?,price_cache_read=?,price_cache_write=?,updated_at=? WHERE id=?')
                .bind(b.name,b.protocol||'openai',b.base||'',keyStored,b.model||'',uc,etype,tplId,paramsTxt,note,fp,kh,b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0,Date.now(),b.id).run();
        }else{
            await env.DB.prepare('INSERT INTO engines_public (id,username,name,protocol,base_url,api_key,model,use_cache,engine_type,tpl_id,params,key_note,key_fp,key_hash,price_in,price_out,price_cache_read,price_cache_write,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)')
                .bind(id,un,b.name,b.protocol||'openai',b.base||'',keyStored,b.model||'',uc,etype,tplId,paramsTxt,note,fp,kh,b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0,Date.now()).run();
        }
        return jr({ok:true,id});
    }catch(e){return jr({error:e.message},500);}
}

async function aEnginesDelete(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}if(!b.id)return jr({error:'缺id'},400);try{await env.DB.prepare('DELETE FROM engines_public WHERE id=?').bind(b.id).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}

/* ★ 补全历史引擎的 Key 指纹 / 哈希（分批，前端循环调） */
async function aEnginesKeyFp(request,env){
    let b;try{b=await request.json();}catch(e){b={};}
    const ids=Array.isArray(b.ids)?b.ids.slice(0,15):[];
    if(!ids.length)return jr({error:'缺 ids'},400);
    try{
        const ph=ids.map(()=>'?').join(',');
        const rows=(await env.DB.prepare('SELECT id,api_key FROM engines_public WHERE id IN ('+ph+')').bind(...ids).all()).results||[];
        let done=0,fail=0;
        for(const r of rows){
            if(!r.api_key){continue;}
            const plain=await decKey(r.api_key,env.KEY_SECRET);
            if(!plain){fail++;continue;}
            const fp=keyFp(plain);
            const kh=await keyHash(plain);
            await env.DB.prepare('UPDATE engines_public SET key_fp=?,key_hash=? WHERE id=?').bind(fp,kh,r.id).run();
            done++;
        }
        return jr({ok:true,done,fail});
    }catch(e){return jr({error:e.message},500);}
}

/* ==================== 引擎体检 ==================== */
async function aEnginesHealth(request,env){
    let b;
    try{b=await request.json();}catch(e){b={};}
    const ids=Array.isArray(b.ids)?b.ids.slice(0,20):[];
    const extraModels=Array.isArray(b.extraModels)?b.extraModels.filter(m=>m&&String(m).trim()).slice(0,30):[];
    if(!ids.length)return jr({error:'缺 ids'},400);

    try{
        const ph=ids.map(()=>'?').join(',');
        const rows=(await env.DB.prepare(
            'SELECT id,username,name,protocol,base_url,api_key,model,engine_type,key_fp FROM engines_public WHERE id IN ('+ph+')'
        ).bind(...ids).all()).results||[];

        const tasks=rows.map(async(e)=>{
            const meta={id:e.id,username:e.username,name:e.name,model:e.model||'',engineType:e.engine_type||'chat',keyFp:e.key_fp||''};
            const t0=Date.now();
            try{
                const base=(e.base_url||'').replace(/\/+$/,'');
                if(!base)return Object.assign({},meta,{ok:false,msg:'Base URL 为空',ms:0});
                if(!e.api_key)return Object.assign({},meta,{ok:false,msg:'⚠️ 未分配 API Key（请到模板下发页补齐）',ms:0});

                const key=await decKey(e.api_key,env.KEY_SECRET);
                if(!key)return Object.assign({},meta,{ok:false,msg:'密钥解密失败（检查 KEY_SECRET 是否一致）',ms:0});

                let path='models';
                if(e.protocol==='anthropic')path='v1/models';
                if(e.protocol==='gemini')path='v1beta/models';

                const ctrl=new AbortController();
                const timer=setTimeout(()=>{try{ctrl.abort();}catch(x){}},12000);
                let resp;
                try{
                    resp=await fetch(base+'/'+path,{headers:{'Authorization':'Bearer '+key,'anthropic-version':'2023-06-01'},signal:ctrl.signal});
                }finally{clearTimeout(timer);}

                const ms=Date.now()-t0;
                if(resp.ok){
                    let modelOk=null,modelCount=0,missModels=[];
                    try{
                        const data=await resp.json();
                        let list=[];
                        if(Array.isArray(data.data))list=data.data.map(m=>m.id||m.name).filter(Boolean);
                        else if(Array.isArray(data.models))list=data.models.map(m=>(m.id||m.name||'').replace(/^models\//,'')).filter(Boolean);
                        else if(Array.isArray(data))list=data.map(m=>m.id||m.name||m).filter(Boolean);
                        modelCount=list.length;
                        if(e.model&&list.length)modelOk=list.indexOf(e.model)>=0;
                        if(extraModels.length&&list.length)extraModels.forEach(m=>{if(list.indexOf(m)<0)missModels.push(m);});
                    }catch(err){}
                    return Object.assign({},meta,{ok:true,msg:'正常',ms,modelOk,modelCount,missModels});
                }

                const txt=(await resp.text()).slice(0,120);
                let hint='HTTP '+resp.status;
                if(resp.status===401)hint+=' Key 无效（该用户的Key可能过期或被撤销）';
                else if(resp.status===403)hint+=' 无权限/被封';
                else if(resp.status===402)hint+=' 余额不足（该用户的Key额度用完了）';
                else if(resp.status===404)hint+=' 路径不对（检查 Base URL）';
                else if(resp.status===429)hint+=' 触发限流';
                return Object.assign({},meta,{ok:false,msg:hint+'：'+txt,ms});
            }catch(err){
                const ms=Date.now()-t0;
                const isAbort=(err&&(err.name==='AbortError'||/abort/i.test(err.message||'')));
                return Object.assign({},meta,{ok:false,msg:isAbort?'超时（>12秒无响应）':('异常：'+(err.message||'unknown')),ms});
            }
        });

        const results=await Promise.all(tasks);
        return jr({ok:true,results});
    }catch(e){return jr({error:e.message},500);}
}

/* ==================== ★ 引擎模板库（不含 Key） ==================== */
async function aTemplatesList(env){
    try{
        const rows=(await env.DB.prepare('SELECT * FROM engine_templates ORDER BY name').all()).results||[];
        const cntRows=(await env.DB.prepare("SELECT tpl_id,COUNT(*) AS n FROM engines_public WHERE tpl_id IS NOT NULL AND tpl_id<>'' GROUP BY tpl_id").all()).results||[];
        const cntMap={};
        cntRows.forEach(r=>cntMap[r.tpl_id]=r.n);

        /* 该模板下发的引擎中，有多少个还没配 Key */
        const nkMap={};
        try{
            const nk=(await env.DB.prepare("SELECT tpl_id,COUNT(*) AS n FROM engines_public WHERE tpl_id IS NOT NULL AND tpl_id<>'' AND (api_key IS NULL OR api_key='') GROUP BY tpl_id").all()).results||[];
            nk.forEach(r=>nkMap[r.tpl_id]=r.n);
        }catch(e){}

        const list=rows.map(t=>({
            id:t.id,name:t.name,
            protocol:t.protocol||'openai',
            base:t.base_url||'',
            model:t.model||'',
            useCache:!!t.use_cache,
            engineType:t.engine_type||'chat',
            params:normParams(pj(t.params,null)),
            priceIn:t.price_in,priceOut:t.price_out,
            priceCR:t.price_cache_read,priceCW:t.price_cache_write,
            deployed:cntMap[t.id]||0,
            noKey:nkMap[t.id]||0,
            updatedAt:t.updated_at||0
        }));
        return jr({ok:true,templates:list});
    }catch(e){
        return jr({error:'模板表不存在或查询失败：'+e.message+'（请先执行建表SQL）'},500);
    }
}

/* ★ 模板保存：不再接收/写入 api_key */
async function aTemplatesSave(request,env){
    let b;
    try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const name=(b.name||'').trim();
    if(!name)return jr({error:'模板名必填'},400);
    try{
        const id=b.id||('tpl_'+Math.random().toString(36).slice(2,10));
        const ex=b.id?await env.DB.prepare('SELECT id FROM engine_templates WHERE id=?').bind(b.id).first():null;
        const uc=b.useCache?1:0;
        const etype=(b.engineType==='image')?'image':'chat';
        const np=normParams(b.params);
        const paramsTxt=np?JSON.stringify(np):'';

        if(ex){
            await env.DB.prepare('UPDATE engine_templates SET name=?,protocol=?,base_url=?,model=?,use_cache=?,engine_type=?,params=?,price_in=?,price_out=?,price_cache_read=?,price_cache_write=?,updated_at=? WHERE id=?')
                .bind(name,b.protocol||'openai',b.base||'',b.model||'',uc,etype,paramsTxt,b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0,Date.now(),b.id).run();
        }else{
            /* api_key 列留空字符串，从此不使用 */
            await env.DB.prepare('INSERT INTO engine_templates (id,name,protocol,base_url,api_key,model,use_cache,engine_type,params,price_in,price_out,price_cache_read,price_cache_write,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)')
                .bind(id,name,b.protocol||'openai',b.base||'','',b.model||'',uc,etype,paramsTxt,b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0,Date.now()).run();
        }
        return jr({ok:true,id});
    }catch(e){return jr({error:e.message},500);}
}

async function aTemplatesDelete(request,env){
    let b;
    try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    if(!b.id)return jr({error:'缺id'},400);
    try{
        await env.DB.prepare('DELETE FROM engine_templates WHERE id=?').bind(b.id).run();
        await env.DB.prepare("UPDATE engines_public SET tpl_id='' WHERE tpl_id=?").bind(b.id).run();
        return jr({ok:true});
    }catch(e){return jr({error:e.message},500);}
}

/* ★ 批量下发：逐人分配 Key
   body:{
     templateIds:[...], usernames:[...],
     keyMap:{ username:{key:'sk-..',note:'月$50'} , ... },
     mode:'skip'|'replace'          // 同名引擎处理方式
   }
   注意：同一账号下发多个模板时，共用 keyMap 里该账号填的 Key（同一中转站通常如此）；
        若需不同 Key，请分两次下发。 */
async function aTemplatesDeploy(request,env){
    let b;
    try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const tplIds=Array.isArray(b.templateIds)?b.templateIds:[];
    const usernames=Array.isArray(b.usernames)?b.usernames:[];
    const keyMap=(b.keyMap&&typeof b.keyMap==='object')?b.keyMap:{};
    const mode=(b.mode==='replace')?'replace':'skip';
    if(!tplIds.length)return jr({error:'请至少选择一个模板'},400);
    if(!usernames.length)return jr({error:'请至少选择一个账号'},400);

    try{
        const ph=tplIds.map(()=>'?').join(',');
        const tpls=(await env.DB.prepare('SELECT * FROM engine_templates WHERE id IN ('+ph+')').bind(...tplIds).all()).results||[];
        if(!tpls.length)return jr({error:'所选模板不存在'},404);

        let created=0,updated=0,skipped=0,noKey=0,keySet=0,keyKept=0;
        const errs=[];

        /* 预先加密每人的 Key（同一账号只算一次，省 CPU） */
        const encCache={};
        for(const un of usernames){
            const km=keyMap[un]||{};
            const raw=String(km.key||'').trim();
            if(raw){
                encCache[un]={
                    stored:await encKey(raw,env.KEY_SECRET),
                    fp:keyFp(raw),
                    hash:await keyHash(raw),
                    note:String(km.note||'').slice(0,80)
                };
            }else{
                encCache[un]=null;
                if(km.note)encCache[un]={stored:null,fp:null,hash:null,note:String(km.note||'').slice(0,80)};
            }
        }

        for(const un of usernames){
            const kc=encCache[un];
            for(const t of tpls){
                try{
                    const ex=await env.DB.prepare('SELECT id,api_key,key_fp,key_hash,key_note FROM engines_public WHERE username=? AND name=?').bind(un,t.name).first();

                    if(ex){
                        if(mode==='skip'){skipped++;continue;}
                        /* 覆盖模式：共性配置全刷；Key 只在本次填了才换，否则保留 */
                        let ks=ex.api_key,kf=ex.key_fp||'',kh=ex.key_hash||'',kn=ex.key_note||'';
                        if(kc&&kc.stored){ks=kc.stored;kf=kc.fp;kh=kc.hash;keySet++;}
                        else keyKept++;
                        if(kc&&kc.note)kn=kc.note;
                        if(!ks)noKey++;

                        await env.DB.prepare('UPDATE engines_public SET protocol=?,base_url=?,api_key=?,model=?,use_cache=?,engine_type=?,tpl_id=?,params=?,key_note=?,key_fp=?,key_hash=?,price_in=?,price_out=?,price_cache_read=?,price_cache_write=?,updated_at=? WHERE id=?')
                            .bind(t.protocol||'openai',t.base_url||'',ks||'',t.model||'',t.use_cache?1:0,t.engine_type||'chat',t.id,t.params||'',kn,kf,kh,t.price_in||0,t.price_out||0,t.price_cache_read||0,t.price_cache_write||0,Date.now(),ex.id).run();
                        updated++;
                    }else{
                        const ks=(kc&&kc.stored)?kc.stored:'';
                        const kf=(kc&&kc.fp)?kc.fp:'';
                        const kh=(kc&&kc.hash)?kc.hash:'';
                        const kn=(kc&&kc.note)?kc.note:'';
                        if(!ks)noKey++;else keySet++;

                        const eid='eng_'+un+'_'+Math.random().toString(36).slice(2,8);
                        await env.DB.prepare('INSERT INTO engines_public (id,username,name,protocol,base_url,api_key,model,use_cache,engine_type,tpl_id,params,key_note,key_fp,key_hash,price_in,price_out,price_cache_read,price_cache_write,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)')
                            .bind(eid,un,t.name,t.protocol||'openai',t.base_url||'',ks,t.model||'',t.use_cache?1:0,t.engine_type||'chat',t.id,t.params||'',kn,kf,kh,t.price_in||0,t.price_out||0,t.price_cache_read||0,t.price_cache_write||0,Date.now()).run();
                        created++;
                    }
                }catch(e){errs.push(un+' / '+t.name+'：'+e.message);}
            }
        }
        return jr({ok:true,created,updated,skipped,noKey,keySet,keyKept,errors:errs});
    }catch(e){return jr({error:e.message},500);}
}

/* ★ 同步：绝不触碰 api_key / key_fp / key_hash / key_note
   （模板本身已无 Key，物理上不可能覆盖任何人的 Key） */
async function aTemplatesSync(request,env){
    let b;
    try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const tplId=(b.templateId||'').trim();
    if(!tplId)return jr({error:'缺 templateId'},400);
    const syncName=!!b.syncName;

    try{
        const t=await env.DB.prepare('SELECT * FROM engine_templates WHERE id=?').bind(tplId).first();
        if(!t)return jr({error:'模板不存在'},404);
        const rows=(await env.DB.prepare('SELECT id FROM engines_public WHERE tpl_id=?').bind(tplId).all()).results||[];
        let n=0;
        for(const r of rows){
            if(syncName){
                await env.DB.prepare('UPDATE engines_public SET name=?,protocol=?,base_url=?,model=?,use_cache=?,engine_type=?,params=?,price_in=?,price_out=?,price_cache_read=?,price_cache_write=?,updated_at=? WHERE id=?')
                    .bind(t.name,t.protocol||'openai',t.base_url||'',t.model||'',t.use_cache?1:0,t.engine_type||'chat',t.params||'',t.price_in||0,t.price_out||0,t.price_cache_read||0,t.price_cache_write||0,Date.now(),r.id).run();
            }else{
                await env.DB.prepare('UPDATE engines_public SET protocol=?,base_url=?,model=?,use_cache=?,engine_type=?,params=?,price_in=?,price_out=?,price_cache_read=?,price_cache_write=?,updated_at=? WHERE id=?')
                    .bind(t.protocol||'openai',t.base_url||'',t.model||'',t.use_cache?1:0,t.engine_type||'chat',t.params||'',t.price_in||0,t.price_out||0,t.price_cache_read||0,t.price_cache_write||0,Date.now(),r.id).run();
            }
            n++;
        }
        return jr({ok:true,count:n});
    }catch(e){return jr({error:e.message},500);}
}

/* ==================== 档位模板库 ==================== */
async function aSlotTplList(env){
    try{
        const rows=(await env.DB.prepare('SELECT * FROM slot_templates ORDER BY name').all()).results||[];
        const list=rows.map(t=>{
            const slots=normSlots(pj(t.data,[]));
            return{id:t.id,name:t.name,slots:slots,slotCount:slots.length,updatedAt:t.updated_at||0};
        });
        return jr({ok:true,templates:list});
    }catch(e){return jr({error:'档位模板表不存在：'+e.message+'（请先执行建表SQL）'},500);}
}
async function aSlotTplSave(request,env){
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const name=(b.name||'').trim();
    if(!name)return jr({error:'模板名必填'},400);
    try{
        const id=b.id||('slt_'+Math.random().toString(36).slice(2,10));
        const slots=normSlots(b.slots||[]);
        const data=JSON.stringify(slots);
        const ex=b.id?await env.DB.prepare('SELECT id FROM slot_templates WHERE id=?').bind(b.id).first():null;
        if(ex)await env.DB.prepare('UPDATE slot_templates SET name=?,data=?,updated_at=? WHERE id=?').bind(name,data,Date.now(),b.id).run();
        else await env.DB.prepare('INSERT INTO slot_templates (id,name,data,updated_at) VALUES (?,?,?,?)').bind(id,name,data,Date.now()).run();
        return jr({ok:true,id,count:slots.length});
    }catch(e){return jr({error:e.message},500);}
}
async function aSlotTplDelete(request,env){
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    if(!b.id)return jr({error:'缺id'},400);
    try{await env.DB.prepare('DELETE FROM slot_templates WHERE id=?').bind(b.id).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}
}
async function aSlotTplDeploy(request,env){
    let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}
    const tplId=(b.templateId||'').trim();
    const usernames=Array.isArray(b.usernames)?b.usernames:[];
    const mode=(b.mode==='append')?'append':'replace';
    if(!tplId)return jr({error:'缺 templateId'},400);
    if(!usernames.length)return jr({error:'请至少选择一个账号'},400);
    try{
        const t=await env.DB.prepare('SELECT * FROM slot_templates WHERE id=?').bind(tplId).first();
        if(!t)return jr({error:'档位模板不存在'},404);
        const tplSlots=normSlots(pj(t.data,[]));
        if(!tplSlots.length)return jr({error:'该模板没有任何档位'},400);
        let n=0;
        const errs=[];
        for(const un of usernames){
            try{
                let finalSlots;
                if(mode==='append'){
                    const u=await env.DB.prepare('SELECT quick_models FROM users WHERE username=?').bind(un).first();
                    const cur=normSlots(pj(u&&u.quick_models,[]));
                    const exists={};
                    cur.forEach(s=>exists[s.label]=1);
                    finalSlots=cur.concat(tplSlots.filter(s=>!exists[s.label]));
                }else finalSlots=tplSlots;
                finalSlots=normSlots(finalSlots);
                await env.DB.prepare('UPDATE users SET quick_models=? WHERE username=?').bind(finalSlots.length?JSON.stringify(finalSlots):'',un).run();
                n++;
            }catch(e){errs.push(un+'：'+e.message);}
        }
        return jr({ok:true,count:n,slotCount:tplSlots.length,errors:errs});
    }catch(e){return jr({error:e.message},500);}
}

/* ==================== 预设 / 监视 / 配置 / 模型库 ==================== */
async function aPresetsGet(env){try{const row=await env.DB.prepare('SELECT data FROM presets WHERE id=1').first();return jr({ok:true,presets:row&&row.data?JSON.parse(row.data):null});}catch(e){return jr({ok:true,presets:null});}}
async function aPresetsSave(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}if(!b.presets)return jr({error:'无预设数据'},400);try{const data=JSON.stringify(b.presets);const ex=await env.DB.prepare('SELECT id FROM presets WHERE id=1').first();if(ex)await env.DB.prepare('UPDATE presets SET data=?,updated_at=? WHERE id=1').bind(data,Date.now()).run();else await env.DB.prepare('INSERT INTO presets (id,data,updated_at) VALUES (1,?,?)').bind(data,Date.now()).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aMonitor(env){try{const dayAgo=Date.now()-5*60*1000;const online=(await env.DB.prepare('SELECT COUNT(DISTINCT username) AS n FROM sessions WHERE last_active>?').bind(dayAgo).first())||{n:0};const weekAgo=Date.now()-7*24*3600*1000;const sess=(await env.DB.prepare('SELECT username,MAX(last_active) AS last,COUNT(DISTINCT ip) AS ipc FROM sessions WHERE last_active>? GROUP BY username').bind(weekAgo).all()).results||[];const logs=(await env.DB.prepare('SELECT username,COUNT(*) AS logCount,SUM(tokens) AS totalTokens FROM logs GROUP BY username').all()).results||[];const recent=(await env.DB.prepare('SELECT username,chat_name,rounds,tokens,model,created_at FROM logs ORDER BY created_at DESC LIMIT 100').all()).results||[];return jr({ok:true,onlineCount:online.n||0,sessions:sess,logs:logs,recent:recent});}catch(e){return jr({error:e.message},500);}}
async function aConfigGet(env){try{const rows=(await env.DB.prepare('SELECT key,value FROM global_config').all()).results||[];const cfg={};rows.forEach(r=>cfg[r.key]=r.value);return jr({ok:true,config:cfg});}catch(e){return jr({ok:true,config:{}});}}
async function aConfigSave(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}try{for(const k in(b.config||{})){const v=String(b.config[k]);const ex=await env.DB.prepare('SELECT key FROM global_config WHERE key=?').bind(k).first();if(ex)await env.DB.prepare('UPDATE global_config SET value=? WHERE key=?').bind(v,k).run();else await env.DB.prepare('INSERT INTO global_config (key,value) VALUES (?,?)').bind(k,v).run();}return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aModelsList(env){try{const rows=(await env.DB.prepare('SELECT * FROM model_prices ORDER BY model_name').all()).results||[];return jr({ok:true,models:rows});}catch(e){return jr({error:e.message},500);}}
async function aModelsSave(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}const mn=(b.model_name||'').trim();if(!mn)return jr({error:'模型名必填'},400);try{const ex=await env.DB.prepare('SELECT model_name FROM model_prices WHERE model_name=?').bind(mn).first();if(ex)await env.DB.prepare('UPDATE model_prices SET price_in=?,price_out=?,price_cache_read=?,price_cache_write=? WHERE model_name=?').bind(b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0,mn).run();else await env.DB.prepare('INSERT INTO model_prices (model_name,price_in,price_out,price_cache_read,price_cache_write) VALUES (?,?,?,?,?)').bind(mn,b.priceIn||0,b.priceOut||0,b.priceCR||0,b.priceCW||0).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}
async function aModelsDelete(request,env){let b;try{b=await request.json();}catch(e){return jr({error:'格式错误'},400);}if(!b.model_name)return jr({error:'缺模型名'},400);try{await env.DB.prepare('DELETE FROM model_prices WHERE model_name=?').bind(b.model_name).run();return jr({ok:true});}catch(e){return jr({error:e.message},500);}}

/* ==================== AI 请求代理 ==================== */
async function hProxy(request,env,url,sub){
    const requestId=(crypto.randomUUID?crypto.randomUUID():Math.random().toString(36).slice(2))+'';
    const pl=await verifyUser(request,env);
    if(!pl)return jr({error:'未登录或登录已过期，请重新登录'},401);
    const auth=request.headers.get('X-Auth-Token')||'';
    if(env.DB){try{await env.DB.prepare('UPDATE sessions SET last_active=? WHERE session_id=?').bind(Date.now(),auth.slice(-16)).run();}catch(e){}}

    const engineId=request.headers.get('X-Engine-Id')||'';
    let targetBase,apiKey='';
    if(engineId){
        const e=await env.DB.prepare('SELECT * FROM engines_public WHERE id=? AND username=?').bind(engineId,pl.username).first();
        if(!e)return jr({error:'公有引擎不存在或无权使用'},403);
        if(!e.api_key)return jr({error:'该引擎尚未分配 API Key，请联系管理员'},400);
        targetBase=e.base_url;
        apiKey=await decKey(e.api_key,env.KEY_SECRET);
        if(!apiKey)return jr({error:'密钥解密失败，请联系管理员检查配置'},500);
    }else{
        targetBase=request.headers.get('X-Target-Base');
        if(!targetBase)return jr({error:'Missing X-Target-Base'},400);
    }

    const targetUrl=targetBase.replace(/\/+$/,'')+'/'+sub+url.search;
    const headers=new Headers();
    const skip=['host','cf-connecting-ip','cf-ray','cf-visitor','cf-worker','cf-ipcountry','cf-ew-via','x-target-base','x-auth-token','x-engine-id','content-length','authorization','content-type'];
    for(const[k,v]of request.headers){if(!skip.includes(k.toLowerCase()))headers.set(k,v);}

    const ct=request.headers.get('Content-Type')||'';
    if(ct&&ct.indexOf('multipart/form-data')<0)headers.set('Content-Type',ct);
    if(engineId)headers.set('Authorization','Bearer '+apiKey);
    else{const oa=request.headers.get('Authorization');if(oa)headers.set('Authorization',oa);}

    const isAnthropic=/\/messages\b/.test(targetUrl)||/anthropic/i.test(targetBase);
    if(isAnthropic&&!headers.has('anthropic-version'))headers.set('anthropic-version','2023-06-01');

    try{
        const isBodyless=(request.method==='GET'||request.method==='HEAD');
        let body;
        if(!isBodyless){
            if(ct.indexOf('multipart/form-data')>=0)body=await request.formData();
            else body=request.body;
        }
        const startedAt=Date.now();
        const resp=await fetch(targetUrl,{method:request.method,headers,body:body});
        const nh=new Headers(resp.headers);
        nh.set('Access-Control-Allow-Origin','*');
        nh.set('Access-Control-Expose-Headers','*');
        nh.set('X-Feifan-Request-Id',requestId);
        nh.set('X-Feifan-Upstream-Status',String(resp.status));
        nh.set('X-Feifan-Upstream-Time',String(Date.now()-startedAt));
        return new Response(resp.body,{status:resp.status,statusText:resp.statusText,headers:nh});
    }catch(e){console.error('[Proxy '+requestId+']',e);return jr({error:'Proxy failed: '+e.message,requestId:requestId},502);}
}
