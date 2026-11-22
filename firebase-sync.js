/* Pro Sheet — Firestore sync (dashboard + prosheet)
   Put as the FIRST script in <head> of BOTH dashboard.html and prosheet.html:
     <script src="firebase-sync.js"></script>
   It mirrors everything the apps keep in localStorage to Firestore, matching the security rules:
     users/{uid}/documents/{id}            full workbook (gzip+base64 when large)
     users/{uid}/documents/{id}/versions   automatic snapshots (every 30+ min of editing, last 20 kept)
     users/{uid}/dashboard/{id}            dashboard card data (name, dates, sheets, cells, cloud flag…)
     users/{uid}/settings/prefs            view mode, theme, recent colors
     users/{uid}/trash/{id}                deleted spreadsheets (restorable)
     users/{uid}/activity/{eventId}        created / edited / opened / renamed / deleted / restored
   Sign-in: reuses your existing Google button (its ID token is exchanged for a Firebase session). */
(function(){
"use strict";
/* Shared Firebase web config (public by design; access is enforced by the Firestore rules). */
const CFG={
  apiKey:"AIzaSyAKGzMSAPluvjKf0pA_TLv2uEb0njnmmhI",
  authDomain:"pro-sheet-38194.firebaseapp.com",
  databaseURL:"https://pro-sheet-38194-default-rtdb.firebaseio.com",
  projectId:"pro-sheet-38194",
  storageBucket:"pro-sheet-38194.firebasestorage.app",
  messagingSenderId:"121904110568",
  appId:"1:121904110568:web:418c36bcf071f47199a25d"
};

const SDK="https://www.gstatic.com/firebasejs/10.12.2/";
const K={IDX:'prosheet_documents_index',DOC:'prosheet_doc_',AUTH:'prosheet_auth',PEND:'prosheet_fb_pending',UID:'prosheet_fb_uid',PAT:'prosheet_fb_prefs_at'};
const PREFS={dash_view:'prosheet_dash_view',theme:'prosheet_theme',recent_text:'prosheet_recent_text_colors',recent_fill:'prosheet_recent_fill_colors'};
let fb=null,uid=null,muted=false,lastIdx={};
const timers={},deb=(k,f,ms)=>{clearTimeout(timers[k]);timers[k]=setTimeout(f,ms)};
const $=id=>document.getElementById(id),esc=s=>String(s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const loadIdx=()=>{try{return JSON.parse(localStorage.getItem(K.IDX))||[]}catch(e){return[]}};
const snapIdx=l=>{const o={};l.forEach(d=>o[d.id]={name:d.name,fb:!!d.fb});return o};
const _set=Storage.prototype.setItem,_rm=Storage.prototype.removeItem;
const lset=(k,v)=>{muted=true;try{_set.call(localStorage,k,v)}finally{muted=false}};
const lrm=k=>{muted=true;try{_rm.call(localStorage,k)}finally{muted=false}};
const getPend=()=>{try{return JSON.parse(localStorage.getItem(K.PEND))||[]}catch(e){return[]}};
const setPend=l=>lset(K.PEND,JSON.stringify([...new Set(l)]));
const isPrs=k=>Object.values(PREFS).includes(k);
function toast(m){const t=$('toast');if(t){t.textContent=m;t.classList.add('show');clearTimeout(t._f);t._f=setTimeout(()=>t.classList.remove('show'),2200)}}
function status(s,t){const a=$('syncStatus');if(a){a.className='sync '+s;a.textContent=t}
  const b=$('saveStatus');if(b){b.className='save-status '+({syncing:'syncing',ok:'synced',err:'error'}[s]||'');b.textContent=t}}

/* ---- watch localStorage writes made by the apps ---- */
Storage.prototype.setItem=function(k,v){_set.apply(this,arguments);if(this!==localStorage||muted)return;k=String(k);
  if(k===K.AUTH)deb('auth',trySignIn,200);
  else if(k===K.IDX)deb('idx',indexChanged,700);
  else if(k.startsWith(K.DOC))deb(k,()=>pushDoc(k.slice(K.DOC.length)),3000);
  else if(isPrs(k))deb('prefs',pushPrefs,1500)};
Storage.prototype.removeItem=function(k){_rm.apply(this,arguments);if(this===localStorage&&!muted&&k===K.AUTH&&fb)fb.A.signOut(fb.auth).catch(()=>{})};

/* ---- gzip helpers (keeps big workbooks under Firestore's 1 MB limit) ---- */
async function enc(s){if(s.length<60000||typeof CompressionStream==='undefined')return{json:s,enc:'raw'};
  const buf=new Uint8Array(await new Response(new Blob([s]).stream().pipeThrough(new CompressionStream('gzip'))).arrayBuffer());let b='';
  for(let i=0;i<buf.length;i+=8192)b+=String.fromCharCode.apply(null,buf.subarray(i,i+8192));return{json:btoa(b),enc:'gz'}}
async function dec(o){if(o.enc!=='gz')return o.json;const bin=atob(o.json),u=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)u[i]=bin.charCodeAt(i);
  return new Response(new Blob([u]).stream().pipeThrough(new DecompressionStream('gzip'))).text()}

/* ---- Firestore helpers ---- */
const ref=(...p)=>fb.F.doc(fb.db,'users',uid,...p),col=(...p)=>fb.F.collection(fb.db,'users',uid,...p);
const last={};
async function log(type,id,name){if(!uid)return;try{await fb.F.addDoc(col('activity'),{type,docId:id||'',name:String(name||'').slice(0,200),at:Date.now()})}catch(e){}}
function markFb(id){const l=loadIdx(),d=l.find(x=>x.id===id);if(d&&!d.fb){d.fb=true;d.cloud=true;lset(K.IDX,JSON.stringify(l));lastIdx=snapIdx(l)}}

async function pushDoc(id){if(!uid||!fb)return;const raw=localStorage.getItem(K.DOC+id);if(!raw)return;let wb;try{wb=JSON.parse(raw)}catch(e){return}
  const meta=loadIdx().find(d=>d.id===id)||{},now=Date.now(),name=String(meta.name||wb.filename||'Untitled spreadsheet').slice(0,200),
    created=meta.createdAt||now,upd=wb.lastSavedAt||meta.updatedAt||now,e=await enc(raw);
  if(e.json.length>900000){status('err','Too large to sync');return}
  try{status('syncing','Syncing…');const F=fb.F;
    await F.setDoc(ref('documents',id),{name,createdAt:created,updatedAt:upd,json:e.json,enc:e.enc,size:raw.length});
    const cells=Object.values(wb.sheets||{}).reduce((n,s)=>n+Object.keys(s.data||{}).length,0);
    await F.setDoc(ref('dashboard',id),{name,createdAt:created,updatedAt:upd,cloud:true,sheets:(wb.order||[]).length,cells,starred:!!meta.starred,folder:meta.folder||null},{merge:true});
    markFb(id);
    if(!last['e'+id]||now-last['e'+id]>6e5){last['e'+id]=now;log('edited',id,name)}
    if(!last['v'+id]){const t=+localStorage.getItem('prosheet_fb_ver_'+id)||0;last['v'+id]=t}
    if(now-last['v'+id]>18e5){last['v'+id]=now;lset('prosheet_fb_ver_'+id,String(now));await snapshotVersion(id,name,e);}
    status('ok','Synced to cloud')}
  catch(err){status('err','Sync failed');console.warn('pushDoc',err.message)}}
async function snapshotVersion(id,name,e){const F=fb.F;await F.setDoc(ref('documents',id,'versions',String(Date.now())),{name,at:Date.now(),json:e.json,enc:e.enc});
  const s=await F.getDocs(F.query(col('documents',id,'versions'),F.orderBy('at','desc')));let i=0;
  for(const d of s.docs){if(++i>20)await F.deleteDoc(d.ref)}}
async function trashDoc(id){const F=fb.F,s=await F.getDoc(ref('documents',id));
  if(s.exists()){await F.setDoc(ref('trash',id),{...s.data(),deletedAt:Date.now()});log('deleted',id,s.data().name)}
  await F.deleteDoc(ref('documents',id));await F.deleteDoc(ref('dashboard',id))}
async function pushPrefs(){if(!uid||!fb)return;const o={updatedAt:Date.now()};Object.entries(PREFS).forEach(([f,k])=>{const v=localStorage.getItem(k);if(v!==null)o[f]=v});
  try{await fb.F.setDoc(ref('settings','prefs'),o);lset(K.PAT,String(o.updatedAt))}catch(e){console.warn('prefs',e.message)}}
async function indexChanged(){const cur=loadIdx(),now={};cur.forEach(d=>now[d.id]=d);
  for(const id in lastIdx)if(!now[id]){if(uid)trashDoc(id).catch(()=>setPend([...getPend(),id]));else if(lastIdx[id].fb)setPend([...getPend(),id])}
  for(const d of cur){const o=lastIdx[d.id];if(!o)log('created',d.id,d.name);else if(o.name!==d.name)log('renamed',d.id,d.name)}
  lastIdx=snapIdx(cur)}

/* ---- full two-way reconcile (runs after sign-in, on focus, when back online) ---- */
let syncing=false;
async function fullSync(reloadOk){if(!uid||!fb||syncing)return;syncing=true;status('syncing','Syncing…');
  try{const F=fb.F;
    const prev=localStorage.getItem(K.UID);
    if(prev&&prev!==uid){const l=loadIdx().filter(d=>{if(d.fb){lrm(K.DOC+d.id);return false}return true});lset(K.IDX,JSON.stringify(l))}
    lset(K.UID,uid);
    const pend=getPend(),left=[];for(const id of pend){try{await trashDoc(id)}catch(e){left.push(id)}}setPend(left);
    const [ds,ps]=await Promise.all([F.getDocs(col('dashboard')),F.getDoc(ref('settings','prefs'))]);
    const cloud={};ds.forEach(d=>cloud[d.id]=d.data());
    let idx=loadIdx(),changed=false,curId=new URLSearchParams(location.search).get('id'),curChanged=false;
    idx=idx.filter(d=>{if(d.fb&&!cloud[d.id]&&!left.includes(d.id)){lrm(K.DOC+d.id);changed=true;if(d.id===curId)curChanged=true;return false}return true});
    const by={};idx.forEach(d=>by[d.id]=d);
    for(const id in cloud){const c=cloud[id],l=by[id];
      if(!l||!localStorage.getItem(K.DOC+id)||(c.updatedAt||0)>(l.updatedAt||0)){
        const s=await F.getDoc(ref('documents',id));if(!s.exists())continue;
        lset(K.DOC+id,await dec(s.data()));
        const e={id,name:c.name,createdAt:c.createdAt,updatedAt:c.updatedAt,cloud:true,fb:true,starred:c.starred||false,folder:c.folder||null};
        if(l)Object.assign(l,e);else{idx.unshift(e);by[id]=e}changed=true;if(id===curId)curChanged=true}}
    lset(K.IDX,JSON.stringify(idx));lastIdx=snapIdx(idx);
    for(const d of idx){const c=cloud[d.id];if(!d.fb||!c||(d.updatedAt||0)>(c.updatedAt||0)){if(localStorage.getItem(K.DOC+d.id))await pushDoc(d.id)}}
    const p=ps.exists()?ps.data():null,lat=+localStorage.getItem(K.PAT)||0;
    if(p&&(p.updatedAt||0)>lat){Object.entries(PREFS).forEach(([f,k])=>{if(p[f]!=null)lset(k,p[f])});lset(K.PAT,String(p.updatedAt));changed=true}
    else if(!p||lat>(p.updatedAt||0))await pushPrefs();
    status('ok','Synced to cloud');
    const page=$('grid')?'sheet':'dash';
    if(changed&&((page==='dash'&&reloadOk!==false)||(page==='sheet'&&curChanged))){
      if(page==='sheet'&&reloadOk===false)status('ok','Newer version available — refresh');
      else if(!sessionStorage.getItem('fb_rl')){sessionStorage.setItem('fb_rl','1');location.reload();return}}
    sessionStorage.removeItem('fb_rl');
    if(curId&&!sessionStorage.getItem('fb_op_'+curId)){sessionStorage.setItem('fb_op_'+curId,'1');const m=by[curId];if(m)log('opened',curId,m.name)}
  }catch(e){status('err','Sync failed');console.warn('fullSync',e.message)}finally{syncing=false}}

/* ---- auth: exchange the Google ID token from your existing sign-in for a Firebase session ---- */
async function trySignIn(){if(!fb)return;let a=null;try{a=JSON.parse(localStorage.getItem(K.AUTH))}catch(e){}
  if(!a||!a.idToken)return;if(fb.auth.currentUser&&fb.auth.currentUser.email===a.email)return;
  try{await fb.A.signInWithCredential(fb.auth,fb.A.GoogleAuthProvider.credential(a.idToken))}
  catch(e){status('err','Firebase sign-in failed');console.warn('Firebase sign-in:',e.code||e.message)}}

/* ---- extra UI: Trash/Activity (dashboard) and Version history (prosheet) ---- */
function openDlg(h,w){const b=$('dialogBox'),o=$('dialogOverlay');if(!b||!o)return;b.innerHTML=h;b.style.width=w||'';o.classList.add('show');return b}
const when=t=>new Date(t).toLocaleString();
async function cloudDlg(){if(!uid)return toast('Sign in to use cloud features');
  const F=fb.F,[t,a]=await Promise.all([F.getDocs(col('trash')),F.getDocs(F.query(col('activity'),F.orderBy('at','desc'),F.limit(20)))]);
  const trash=t.docs.map(d=>({id:d.id,...d.data()})).sort((x,y)=>y.deletedAt-x.deletedAt);
  const b=openDlg(`<h3>Cloud · Trash &amp; activity</h3><div style="font:600 11px 'DM Mono',monospace;color:var(--faint);margin:6px 0">TRASH (${trash.length})</div>
  ${trash.map(d=>`<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)"><div style="flex:1;min-width:0"><div style="font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${esc(d.name)}</div><div style="font-size:11px;color:var(--faint)">Deleted ${when(d.deletedAt)}</div></div><button class="btn" data-fa="restore" data-id="${d.id}">Restore</button><button class="btn" data-fa="purge" data-id="${d.id}" style="color:var(--danger)">Delete</button></div>`).join('')||'<div style="font-size:12.5px;color:var(--dim)">Trash is empty.</div>'}
  <div style="font:600 11px 'DM Mono',monospace;color:var(--faint);margin:14px 0 6px">RECENT ACTIVITY</div>
  <div style="max-height:170px;overflow:auto;font-size:12px">${a.docs.map(d=>{const x=d.data();return`<div style="padding:3px 0;color:var(--dim)"><b style="color:var(--text)">${esc(x.type)}</b> · ${esc(x.name)} <span style="float:right;color:var(--faint)">${when(x.at)}</span></div>`}).join('')||'No activity yet.'}</div>
  <div class="dialog-actions"><button class="btn" id="fbClose">Close</button></div>`,'460px');
  $('fbClose').onclick=()=>$('dialogOverlay').classList.remove('show');
  b.querySelectorAll('[data-fa]').forEach(x=>x.onclick=async()=>{const id=x.dataset.id;x.disabled=true;
    try{if(x.dataset.fa==='purge'){await F.deleteDoc(ref('trash',id));toast('Deleted forever');cloudDlg()}
      else{const d=trash.find(q=>q.id===id),{id:_i,deletedAt,...rest}=d;await F.setDoc(ref('documents',id),rest);
        await F.setDoc(ref('dashboard',id),{name:rest.name,createdAt:rest.createdAt,updatedAt:rest.updatedAt,cloud:true},{merge:true});await F.deleteDoc(ref('trash',id));log('restored',id,rest.name);
        toast('Restored');await fullSync(true);location.reload()}}catch(e){toast('Failed: '+e.message);x.disabled=false}})}
async function versionsDlg(){const id=new URLSearchParams(location.search).get('id');if(!uid||!id)return toast('Sign in to view version history');
  const F=fb.F,s=await F.getDocs(F.query(col('documents',id,'versions'),F.orderBy('at','desc'))),vs=s.docs.map(d=>({id:d.id,...d.data()}));
  const b=openDlg(`<h3>Version history</h3><div style="font-size:11.5px;color:var(--text-dim);margin-bottom:8px">Snapshots are saved automatically while you edit (last 20 kept).</div>
  <div style="max-height:260px;overflow:auto">${vs.map(v=>`<div style="display:flex;align-items:center;gap:8px;padding:6px 0;border-bottom:1px solid var(--border)"><div style="flex:1;font-size:12.5px">${when(v.at)}</div><button class="btn" data-v="${v.id}">Restore</button></div>`).join('')||'<div style="font-size:12.5px">No snapshots yet.</div>'}</div>
  <div class="dialog-actions"><button class="btn" id="fbClose">Close</button></div>`);
  $('fbClose').onclick=()=>$('dialogOverlay').classList.remove('show');
  b.querySelectorAll('[data-v]').forEach(x=>x.onclick=async()=>{try{const v=vs.find(q=>q.id===x.dataset.v),raw=await dec(v),wb=JSON.parse(raw);
    const cur=localStorage.getItem(K.DOC+id);if(cur)await snapshotVersion(id,'Before restore',await enc(cur));
    wb.lastSavedAt=Date.now();const l=loadIdx(),m=l.find(d=>d.id===id);if(m){m.updatedAt=wb.lastSavedAt;lset(K.IDX,JSON.stringify(l))}
    lset(K.DOC+id,JSON.stringify(wb));await pushDoc(id);log('restored',id,wb.filename);location.reload()}catch(e){toast('Restore failed: '+e.message)}})}
function addButtons(){
  if($('templates')){const nr=document.querySelector('.nav-right');if(nr&&!$('fbBtn')){const b=document.createElement('button');b.className='btn';b.id='fbBtn';b.textContent='Trash';b.title='Trash & activity';b.onclick=cloudDlg;nr.insertBefore(b,nr.firstChild)}}
  if($('grid')){const tr=document.querySelector('.topbar-right');if(tr&&!$('fbBtn')){const b=document.createElement('button');b.className='btn';b.id='fbBtn';b.textContent='History';b.title='Version history';b.onclick=versionsDlg;tr.insertBefore(b,tr.firstChild)}}
  const o=$('dialogOverlay');if(o)new MutationObserver(()=>{if(!o.classList.contains('show')&&$('dialogBox'))$('dialogBox').style.width=''}).observe(o,{attributes:true,attributeFilter:['class']})}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',addButtons);else addButtons();

/* ---- boot ---- */
(async function init(){
  if(/^PASTE_/.test(CFG.apiKey)){console.warn('firebase-sync.js: paste your Firebase config at the top of the file.');return}
  try{
    const [A,F]=await Promise.all([import(SDK+'firebase-app.js'),import(SDK+'firebase-firestore.js')]),Au=await import(SDK+'firebase-auth.js');
    const app=A.initializeApp(CFG);let db;
    try{db=F.initializeFirestore(app,{localCache:F.persistentLocalCache({tabManager:F.persistentMultipleTabManager()})})}catch(e){db=F.getFirestore(app)}
    fb={A:Au,F,db,auth:Au.getAuth(app)};lastIdx=snapIdx(loadIdx());
    Au.onAuthStateChanged(fb.auth,u=>{if(u){uid=u.uid;fullSync(true)}else{uid=null;trySignIn()}});
    addEventListener('online',()=>fullSync(true));
    document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&Date.now()-(last.vs||0)>3e4){last.vs=Date.now();fullSync(false)}});
    window.ProSheetCloud={sync:()=>fullSync(true),trash:cloudDlg,history:versionsDlg,pushDoc,uid:()=>uid};
  }catch(e){console.warn('Firebase failed to load:',e.message);status('err','Cloud unavailable')}
})();
})();
