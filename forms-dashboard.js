/* Pro Sheet — Form Builder (Zoho-style) + Spreadsheet-to-Dashboard.
   Load AFTER excel-ui.js:  <script src="forms-dashboard.js"></script> */
(function(){
"use strict";
function boot(){
const API=window.ProSheetAPI,PC=window.ProSheetCharts;
if(!API||!API.getWorkbook||!PC||!document.getElementById('chartBtn'))return setTimeout(boot,250);
const $=(s,r)=>(r||document).querySelector(s),WB=()=>API.getWorkbook(),L=API.colLetter,toast=m=>API.toast(m);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const uid=p=>p+Date.now().toString(36)+Math.random().toString(36).slice(2,5);
const persist=(()=>{let t;return()=>{clearTimeout(t);t=setTimeout(()=>API.refresh(),400)}})();

const css=document.createElement('style');
css.textContent=`
#toast{z-index:1000!important}
.xo{position:fixed;inset:0;z-index:800;display:flex;flex-direction:column;font-family:'DM Sans',sans-serif;color:#222;user-select:none}
.xm{position:fixed;inset:0;z-index:900;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;font-family:'DM Sans',sans-serif;color:#222}
.xm>div{background:#fff;border-radius:8px;padding:18px;width:370px;max-height:86vh;overflow:auto;box-shadow:0 20px 50px rgba(0,0,0,.4)}
.xm h3{margin:0 0 6px}.xm label{display:block;font-size:12px;color:#555;margin:10px 0 3px}
.xm input,.xm select{width:100%;padding:7px;border:1px solid #ccc;border-radius:5px;font:inherit;font-size:13px}
.xm .ck{display:flex;flex-wrap:wrap;gap:4px 12px}.xm .ck label{display:flex;gap:4px;align-items:center;margin:0;color:#222}.xm .ck input{width:auto}
.xm .ac{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}
.xbtn{padding:6px 13px;border-radius:5px;border:1px solid #ccc;background:#fff;cursor:pointer;font:inherit;font-size:13px;color:#222}
.xbtn.g{background:#1fa37a;border-color:#1fa37a;color:#fff;font-weight:600}
.fb{background:#f0f0f0}
.fb-top{display:flex;align-items:center;gap:12px;padding:9px 16px;background:#fff;border-bottom:1px solid #ddd;font-size:13px}
.fb-top b{font-size:17px;color:#1fa37a}.sp{flex:1}
.fb-top select{padding:5px;border:1px solid #ccc;border-radius:4px}
.fb-nav{display:flex;align-items:center;gap:4px;background:#1f2a44;padding:0 10px}
.fb-nav span{padding:11px 16px;color:#c8cfdd;font-size:13px;cursor:pointer}
.fb-nav span.on{background:#1fa37a;color:#fff}
.fb-nav .xbtn{margin:6px}
.fb-main{flex:1;display:flex;min-height:0}
.fb-pal{width:290px;background:#f7f7f7;border-right:1px solid #ddd;overflow:auto;padding:12px}
.fb-pt{text-align:center;font-size:13px;padding:6px;border-bottom:2px solid #1fa37a;margin-bottom:12px}
.fb-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.fb-p{background:#fff;border:1px solid #ddd;border-radius:3px;padding:8px 9px;font-size:12.5px;cursor:grab;display:flex;gap:8px;align-items:center}
.fb-p:hover{border-color:#1fa37a}.fb-p i{font-style:normal;width:22px;color:#666;font-size:12px}
.fb-canvas{flex:1;overflow:auto;padding:24px;background:#ececec}
.fb-card{max-width:620px;margin:0 auto;background:#fff;box-shadow:0 1px 6px rgba(0,0,0,.15);min-height:300px}
.fb-title{padding:18px 20px;border-bottom:1px solid #eee;font-size:22px}
.fb-title input{border:0;font:inherit;width:100%;outline:none;background:transparent;user-select:text}
.fb-f{position:relative;padding:16px 20px;border-bottom:1px solid #eee;border-left:3px solid transparent;cursor:pointer}
.fb-f.sel{background:#f2fbf8;border-left-color:#1fa37a}.fb-f.over{box-shadow:inset 0 3px 0 #1fa37a}
.fb-l{display:block;font-size:14px;margin-bottom:7px}.fb-l s{color:#d33;text-decoration:none}
.fb-h{font-size:11.5px;color:#888;margin-top:4px}
.fb-in{padding:7px 9px;border:1px solid #ccc;border-radius:3px;font:inherit;font-size:13px;background:#fff;user-select:text;max-width:100%}
.fb-in.s{width:180px}.fb-in.m{width:300px}.fb-in.l{width:100%}
.fb-op{display:block;font-size:13px;margin:4px 0}
.fb-rate b{font-size:24px;color:#ccc;cursor:pointer}.fb-rate b.on{color:#f5a623}
.fb-act{position:absolute;right:10px;top:8px;display:none;gap:4px}.fb-f:hover .fb-act{display:flex}
.fb-act span{background:#fff;border:1px solid #ddd;border-radius:3px;padding:1px 6px;font-size:12px}
.fb-empty{padding:70px 20px;text-align:center;color:#999;font-size:13px}
.fb-props{width:300px;background:#fff;border-left:1px solid #ddd;overflow:auto}
.fb-ph{background:#1f2a44;color:#fff;padding:12px 14px;font-size:13px}
.fb-pb{padding:12px 14px;font-size:12.5px}.fb-pb label{display:block;color:#d33;margin:12px 0 4px;font-size:12px}
.fb-pb input[type=text],.fb-pb textarea{width:100%;padding:6px;border:1px solid #ccc;border-radius:3px;font:inherit;user-select:text}
.fb-seg{display:flex;border:1px solid #ccc;border-radius:3px;overflow:hidden}.fb-seg span{flex:1;text-align:center;padding:6px;cursor:pointer}.fb-seg span.on{background:#e8f7f2;color:#1fa37a;box-shadow:inset 0 0 0 1px #1fa37a}
.db{background:#eaf0ed}
.db-top{display:flex;align-items:center;gap:10px;padding:12px 20px;background:#1e7e4a;color:#fff}
.db-top b{font-size:22px;outline:none;user-select:text}
.db-top select{padding:6px;border-radius:4px;border:0}
.db-top .xbtn{background:rgba(255,255,255,.15);border-color:transparent;color:#fff}
.db-body{flex:1;overflow:auto;padding:18px 20px}
.db-g{display:grid;grid-template-columns:repeat(12,1fr);gap:14px}
.db-w{position:relative;background:#fff;border-radius:8px;box-shadow:0 1px 4px rgba(0,0,0,.1);padding:14px;min-width:0}
.db-w.kpi,.db-w.gauge{grid-column:span 4}.db-w.chart,.db-w.table{grid-column:span 6}
.db-w.kpi{color:#fff;box-shadow:none;padding:16px 18px}
.db-w .t{font-size:14px;font-weight:600;margin-bottom:6px;color:#333}
.db-w.kpi .t{color:#fff;font-size:13px}.db-w.kpi big{display:block;font-size:34px;font-weight:700;margin:4px 0}.db-w.kpi small{font-size:12px;opacity:.92}.db-w.kpi em{position:absolute;right:18px;top:34px;font-style:normal;font-size:34px}
.db-w .wt{position:absolute;right:8px;top:6px;display:none;gap:8px;font-size:13px;cursor:pointer;color:#666}.db-w:hover .wt{display:flex}.db-w.kpi .wt{color:#fff}
.db-tb{width:100%;border-collapse:collapse;font-size:13px}.db-tb th{background:#1e7e4a;color:#fff;text-align:left;padding:7px 10px}.db-tb td{padding:7px 10px;background:#e9f5ee;border-bottom:1px solid #fff}
.db-empty{padding:80px;text-align:center;color:#666;grid-column:span 12}
.db .xcb text{fill:#555}.db .xcb .gl{stroke:#e2e2e2}.db .xcb .ax{stroke:#999}
@media(max-width:900px){.db-w.kpi,.db-w.gauge,.db-w.chart,.db-w.table{grid-column:span 12}.fb-pal,.fb-props{width:190px}}`;
document.head.appendChild(css);

function overlay(cls,html){const o=document.createElement('div');o.className='xo '+cls;o.innerHTML=html;
  ['keydown','paste'].forEach(ev=>o.addEventListener(ev,e=>e.stopPropagation()));document.body.appendChild(o);return o}
function modal(html){const m=document.createElement('div');m.className='xm';m.innerHTML='<div>'+html+'</div>';
  ['keydown','paste'].forEach(ev=>m.addEventListener(ev,e=>e.stopPropagation()));m.onmousedown=e=>{if(e.target===m)m.remove()};document.body.appendChild(m);return m}

/* ================= FORM BUILDER ================= */
const forms=()=>WB().forms||(WB().forms=[]);
const FT={text:['Single Line','▭'],area:['Multi Line','▤'],number:['Number','123'],email:['Email','✉'],phone:['Phone','☎'],date:['Date','▦'],
  select:['Dropdown','▾'],radio:['Radio','◉'],check:['Checkbox','☑'],rating:['Rating','★'],section:['Section','—']};
const OPT=['select','radio','check'];
const mk=t=>({id:uid('f'),type:t,label:FT[t][0],hint:'',req:false,size:'m',opts:OPT.includes(t)?['Option 1','Option 2','Option 3']:[]});
function newForm(){const f={id:uid('form'),title:'Contact Us',fields:['text','email','phone','area'].map(mk)};
  ['Name','Email','Phone','Comments'].forEach((l,i)=>f.fields[i].label=l);forms().push(f);return f}
let F=null,selId=null,mode='build',FO=null;

function fieldHtml(f,live){const sel=!live&&f.id===selId,drag=live?'':' draggable="true"',n=` name="${f.id}"`,d=live?'':' disabled';
  const act=live?'':`<div class="fb-act"><span data-a="dup" title="Duplicate">⧉</span><span data-a="del" title="Delete">🗑</span></div>`;
  const lab=`<span class="fb-l">${esc(f.label)}${f.req?'<s> *</s>':''}</span>`,hint=f.hint&&!['text','area','number','email','phone'].includes(f.type)?`<div class="fb-h">${esc(f.hint)}</div>`:'';
  let c='';
  if(f.type==='section')return `<div class="fb-f${sel?' sel':''}" data-id="${f.id}"${drag}><div style="font-size:17px;font-weight:600">${esc(f.label)}</div>${f.hint?`<div class="fb-h">${esc(f.hint)}</div>`:''}${act}</div>`;
  if(['text','number','email','phone','date'].includes(f.type))c=`<input class="fb-in ${f.size}" type="${{text:'text',number:'number',email:'email',phone:'tel',date:'date'}[f.type]}"${n}${d} placeholder="${esc(f.hint)}">`;
  else if(f.type==='area')c=`<textarea class="fb-in ${f.size}" rows="4"${n}${d} placeholder="${esc(f.hint)}"></textarea>`;
  else if(f.type==='select')c=`<select class="fb-in ${f.size}"${n}${d}><option value="">Select…</option>${f.opts.map(o=>`<option>${esc(o)}</option>`).join('')}</select>`;
  else if(f.type==='radio'||f.type==='check')c=f.opts.map(o=>`<label class="fb-op"><input type="${f.type==='radio'?'radio':'checkbox'}"${n} value="${esc(o)}"${d}> ${esc(o)}</label>`).join('');
  else if(f.type==='rating')c=`<span class="fb-rate" data-n="${f.id}">${[1,2,3,4,5].map(i=>`<b data-v="${i}">★</b>`).join('')}</span><input type="hidden"${n}>`;
  return `<div class="fb-f${sel?' sel':''}" data-id="${f.id}"${drag}>${lab}${c}${hint}${act}</div>`}
function canvas(){const live=mode==='preview';
  $('#fbCv',FO).innerHTML=`<div class="fb-card"><div class="fb-title">${live?esc(F.title):`<input data-k="ftitle" value="${esc(F.title)}">`}</div>${live?'<form id="fbForm">':''}${F.fields.length?F.fields.map(f=>fieldHtml(f,live)).join(''):'<div class="fb-empty">Drag a field here, or click one on the left</div>'}${live?'<div style="padding:16px 20px"><button type="submit" class="xbtn g">Submit</button></div></form>':''}</div>`}
function props(){const p=$('#fbPr',FO),f=F.fields.find(x=>x.id===selId);
  if(!f||mode!=='build'){p.innerHTML=`<div class="fb-ph">Form properties</div><div class="fb-pb"><label>Form title</label><input type="text" data-k="ftitle" value="${esc(F.title)}"><label>Response sheet</label><div style="color:#555">${esc(F.sheet||'Created on the first submission')}</div><div class="fb-h" style="margin-top:12px">Each submission is added as a new row in the response sheet.</div><button class="xbtn" id="fbDel" style="margin-top:14px;color:#d33">Delete this form</button></div>`;return}
  p.innerHTML=`<div class="fb-ph">Properties — ${FT[f.type][0]}</div><div class="fb-pb"><label>Field label</label><input type="text" data-k="label" value="${esc(f.label)}">${f.type==='section'?'':`<label>Instructions / placeholder</label><textarea rows="3" data-k="hint">${esc(f.hint)}</textarea><label>Field size</label><div class="fb-seg">${[['s','Small'],['m','Medium'],['l','Large']].map(([k,n])=>`<span data-sz="${k}" class="${f.size===k?'on':''}">${n}</span>`).join('')}</div><label style="color:#222"><input type="checkbox" data-k="req"${f.req?' checked':''}> Required</label>`}${OPT.includes(f.type)?`<label>Options (one per line)</label><textarea rows="5" data-k="opts">${esc(f.opts.join('\n'))}</textarea>`:''}</div>`}
function fbSel(){$('#fbSel',FO).innerHTML=forms().map(f=>`<option value="${f.id}"${f===F?' selected':''}>${esc(f.title)}</option>`).join('')}
function fbAll(){fbSel();FO.querySelectorAll('[data-m]').forEach(x=>x.classList.toggle('on',x.dataset.m===mode));canvas();props()}

function appendRow(data){const w=WB();
  const nm=F.sheet||(F.sheet=(F.title||'Form').replace(/[\[\]:*?\/\\]/g,'').slice(0,20)+' Responses');
  if(!w.sheets[nm]){w.sheets[nm]={data:{},colWidths:{},rowHeights:{},rows:100,merges:[],frozenRow:1,frozenCol:0};w.order.push(nm)}
  const s=w.sheets[nm],hdr=[];
  for(let c=0;c<26;c++){const x=s.data[L(c)+'1'];if(x&&x.v!=="")hdr.push(x.v);else break}
  if(!hdr.length)hdr.push('Timestamp');
  F.fields.filter(f=>f.type!=='section').forEach(f=>{if(!hdr.includes(f.label)&&hdr.length<26)hdr.push(f.label)});
  hdr.forEach((h,c)=>{s.data[L(c)+'1']={v:h,f:{bold:true,bg:'#217346',color:'#ffffff'}};if(!s.colWidths[c])s.colWidths[c]=Math.max(110,h.length*9+24)});
  let row=2;while(s.data['A'+row]&&s.data['A'+row].v!=='')row++;
  if(row+5>s.rows)s.rows=row+20;
  const put=(label,v)=>{const c=hdr.indexOf(label);if(c<0)return;v=String(v);s.data[L(c)+row]={v:v[0]==='='?"'"+v:v,f:{}}};
  put('Timestamp',new Date().toLocaleString());
  F.fields.forEach(f=>{if(f.type!=='section')put(f.label,data[f.id]||'')});
  return nm}
function submitForm(form){const data={},miss=[];
  F.fields.forEach(f=>{if(f.type==='section')return;let v='';
    if(f.type==='check')v=[...form.querySelectorAll(`[name="${f.id}"]:checked`)].map(x=>x.value).join(', ');
    else if(f.type==='radio'){const c=form.querySelector(`[name="${f.id}"]:checked`);v=c?c.value:''}
    else v=(form.elements[f.id]||{}).value||'';
    if(f.req&&!String(v).trim())miss.push(f.label);data[f.id]=v});
  if(miss.length){toast('Required: '+miss.join(', '));return}
  const nm=appendRow(data);form.reset();form.querySelectorAll('.fb-rate b.on').forEach(b=>b.classList.remove('on'));
  API.refresh();toast('Response saved to "'+nm+'"')}

function openForm(){if(FO)return;if(!forms().length)newForm();if(!F||!forms().includes(F))F=forms()[0];selId=null;mode='build';
  FO=overlay('fb',`<div class="fb-top"><b>▣ Forms</b><span>|</span><select id="fbSel"></select><button class="xbtn" id="fbNew">+ New form</button><span class="sp"></span><button class="xbtn" id="fbClose">✕ Close</button></div>
  <div class="fb-nav"><span data-m="build">Builder</span><span data-m="preview">Preview &amp; fill</span><span data-m="entries">All entries</span><span class="sp"></span><button class="xbtn g" id="fbAccess">Access form</button></div>
  <div class="fb-main"><div class="fb-pal"><div class="fb-pt">Basic fields</div><div class="fb-grid">${Object.entries(FT).map(([k,v])=>`<div class="fb-p" draggable="true" data-t="${k}"><i>${v[1]}</i>${v[0]}</div>`).join('')}</div></div><div class="fb-canvas" id="fbCv"></div><div class="fb-props" id="fbPr"></div></div>`);
  const add=(t,idx)=>{const f=mk(t);F.fields.splice(idx==null?F.fields.length:idx,0,f);selId=f.id;mode='build';persist();fbAll()};
  FO.addEventListener('click',e=>{const t=e.target;let x;
    if(t.closest('#fbClose')){FO.remove();FO=null;return}
    if(t.closest('#fbNew')){F=newForm();selId=null;mode='build';persist();fbAll();return}
    if(t.closest('#fbAccess')){mode='preview';selId=null;fbAll();return}
    if(t.closest('#fbDel')){if(confirm('Delete this form? Existing responses stay on their sheet.')){forms().splice(forms().indexOf(F),1);F=forms()[0]||newForm();selId=null;persist();fbAll()}return}
    if((x=t.closest('[data-m]'))){if(x.dataset.m==='entries'){const w=WB();if(F.sheet&&w.sheets[F.sheet]){w.current=F.sheet;FO.remove();FO=null;API.refresh()}else toast('No responses yet — submit the form first');return}mode=x.dataset.m;selId=null;fbAll();return}
    if((x=t.closest('.fb-p'))){add(x.dataset.t);return}
    if((x=t.closest('[data-sz]'))){const f=F.fields.find(q=>q.id===selId);if(f){f.size=x.dataset.sz;persist();canvas();props()}return}
    if(mode==='preview'){if((x=t.closest('.fb-rate b'))){const r=x.parentNode,v=+x.dataset.v;r.nextElementSibling.value=v;r.querySelectorAll('b').forEach(b=>b.classList.toggle('on',+b.dataset.v<=v))}return}
    const fe=t.closest('.fb-f');if(!fe)return;const i=F.fields.findIndex(q=>q.id===fe.dataset.id);
    if((x=t.closest('[data-a]'))){if(x.dataset.a==='del'){F.fields.splice(i,1);selId=null}else{const c=JSON.parse(JSON.stringify(F.fields[i]));c.id=uid('f');F.fields.splice(i+1,0,c);selId=c.id}persist();canvas();props();return}
    selId=fe.dataset.id;canvas();props()});
  FO.addEventListener('input',e=>{const k=e.target.dataset.k;if(!k)return;const f=F.fields.find(q=>q.id===selId),v=e.target.type==='checkbox'?e.target.checked:e.target.value;
    if(k==='ftitle'){F.title=v;fbSel()}else if(f){f[k]=k==='opts'?v.split('\n').map(s=>s.trim()).filter(Boolean):v}
    persist();canvas()});
  FO.addEventListener('change',e=>{if(e.target.id==='fbSel'){F=forms().find(f=>f.id===e.target.value);selId=null;fbAll()}});
  FO.addEventListener('submit',e=>{e.preventDefault();submitForm(e.target)});
  FO.addEventListener('dragstart',e=>{const p=e.target.closest&&e.target.closest('.fb-p'),f=e.target.closest&&e.target.closest('.fb-f');
    if(p)e.dataTransfer.setData('text/plain','new:'+p.dataset.t);else if(f)e.dataTransfer.setData('text/plain','mv:'+f.dataset.id)});
  FO.addEventListener('dragover',e=>{if(!e.target.closest('#fbCv'))return;e.preventDefault();FO.querySelectorAll('.over').forEach(x=>x.classList.remove('over'));const f=e.target.closest('.fb-f');if(f)f.classList.add('over')});
  FO.addEventListener('drop',e=>{if(!e.target.closest('#fbCv')||mode!=='build')return;e.preventDefault();
    const d=e.dataTransfer.getData('text/plain'),tg=e.target.closest('.fb-f');let idx=F.fields.length;
    if(tg){idx=F.fields.findIndex(x=>x.id===tg.dataset.id);const r=tg.getBoundingClientRect();if(e.clientY>r.top+r.height/2)idx++}
    if(d.startsWith('new:'))add(d.slice(4),idx);
    else if(d.startsWith('mv:')){const i=F.fields.findIndex(x=>x.id===d.slice(3));if(i>=0){const[f]=F.fields.splice(i,1);if(i<idx)idx--;F.fields.splice(idx,0,f);persist();fbAll()}}});
  fbAll()}

/* ================= DASHBOARD ================= */
let D=null,DO=null;
const isNum=x=>typeof x==='number';
const numCol=(rows,c)=>{const v=rows.map(r=>r[c]).filter(x=>x!=='');return v.length>0&&v.filter(isNum).length>=v.length*.6};
const fmt=(h,v)=>{if(v===''||v==null)return '–';const s=(+v).toLocaleString(undefined,{maximumFractionDigits:2});
  return /%|rate|churn|margin|percent/i.test(h)?s+'%':/revenue|sales|profit|price|cost|amount|income|expense/i.test(h)?'₹'+s:s};
function agg(a,t){const n=a.filter(isNum);if(t==='count')return a.length;if(!n.length)return 0;const s=n.reduce((x,y)=>x+y,0);
  return t==='sum'?s:t==='avg'?s/n.length:t==='max'?Math.max(...n):t==='min'?Math.min(...n):n[n.length-1]}
function ds(){const snap=API.snapshot().sheets[D.sheet]||[],hdr=[],all=[];
  for(let c=D.c1;c<=D.c2;c++)hdr.push(String((snap[D.r1]||[])[c]==null?'':snap[D.r1][c]));
  for(let r=D.r1+1;r<snap.length;r++){const row=[];for(let c=D.c1;c<=D.c2;c++)row.push((snap[r]||[])[c]==null?'':snap[r][c]);if(row.some(x=>x!==''))all.push(row)}
  return{hdr,all,rows:D.filter&&D.cat>=0?all.filter(r=>String(r[D.cat])===D.filter):all}}
function genDash(){const R=PC.regionSel(),w=WB(),snap=API.snapshot().sheets[w.current]||[];
  if(R.r2-R.r1<1){toast('Select a table (header row + data rows) first');return false}
  const hdr=[],rows=[];
  for(let c=R.c1;c<=R.c2;c++)hdr.push(String((snap[R.r1]||[])[c]==null?'':snap[R.r1][c]));
  for(let r=R.r1+1;r<=R.r2;r++){const row=[];for(let c=R.c1;c<=R.c2;c++)row.push((snap[r]||[])[c]==null?'':snap[r][c]);rows.push(row)}
  const nums=[],cats=[];hdr.forEach((_,i)=>(numCol(rows,i)?nums:cats).push(i));
  if(!nums.length){toast('No numeric columns found in the selection');return false}
  const cat=cats.length?cats[0]:-1,W=(type,cols,x)=>Object.assign({id:uid('w'),type,cols,agg:'last',title:''},x),cn=cat>=0?hdr[cat]:'row';
  D={title:API.getName()||'Dashboard',sheet:w.current,r1:R.r1,c1:R.c1,c2:R.c2,cat,filter:'',widgets:[
    ...nums.slice(0,3).map(i=>W('kpi',[i])),W('line',nums.slice(0,3),{title:'Trend'}),W('gauge',[nums[Math.min(1,nums.length-1)]]),
    W('bar',[nums[0]],{title:hdr[nums[0]]+' by '+cn}),W('table',nums.slice(0,3),{title:'Top '+cn,n:5})]};
  w.dashboard=D;persist();return true}
function gaugeSvg(v,mx,txt){const pct=Math.max(0,Math.min(1,v/mx)),cx=130,cy=118,R=100,r=64,B=['#e74c3c','#f39c12','#f4d03f','#8bc34a','#27ae60'],P=(a,q)=>`${cx+q*Math.cos(a)},${cy-q*Math.sin(a)}`;
  let o=B.map((c,i)=>{const a0=Math.PI*(1-i/5),a1=Math.PI*(1-(i+1)/5);return `<path d="M${P(a0,R)} A${R},${R} 0 0 1 ${P(a1,R)} L${P(a1,r)} A${r},${r} 0 0 0 ${P(a0,r)} Z" fill="${c}" stroke="#fff"/>`}).join('');
  const a=Math.PI*(1-pct);
  o+=`<line x1="${cx}" y1="${cy}" x2="${cx+(R-10)*Math.cos(a)}" y2="${cy-(R-10)*Math.sin(a)}" stroke="#333" stroke-width="3" stroke-linecap="round"/><circle cx="${cx}" cy="${cy}" r="6" fill="#333"/><text x="${cx}" y="${cy+34}" text-anchor="middle" font-size="26" font-weight="700" fill="#222">${esc(txt)}</text>`;
  return `<svg viewBox="0 0 260 165" style="width:100%;max-width:320px;display:block;margin:0 auto">${o}</svg>`}
const KC=['#4caf50','#43a047','#2e86c1','#e67e22'];
function widget(w,i,hdr,rows){const col=w.cols[0],h=hdr[col]||'',t=w.title||h,a=rows.map(r=>r[col]),
  tb=`<div class="wt"><span data-e="${w.id}" title="Edit">⚙</span><span data-x="${w.id}" title="Remove">✕</span></div>`;
  if(w.type==='kpi'){const n=a.filter(isNum),v=agg(a,w.agg),p=w.agg==='last'?n[n.length-2]:(a.length>1?agg(a.slice(0,-1),w.agg):undefined);
    return `<div class="db-w kpi" style="background:${KC[i%4]}">${tb}<div class="t">${esc(t)}</div><big>${fmt(h,v)}</big><em>${p===undefined?'':v>=p?'↗':'↘'}</em><small>${p===undefined?'':'Previous: '+fmt(h,p)}</small></div>`}
  if(w.type==='gauge'){const n=a.filter(isNum),v=agg(a,w.agg),mx=+w.max||(/%|rate|churn|margin|percent/i.test(h)?100:Math.max(1,...n));
    return `<div class="db-w gauge">${tb}<div class="t">${esc(t)}</div>${gaugeSvg(v,mx,fmt(h,v))}</div>`}
  if(w.type==='table'){const srt=[...rows].sort((x,y)=>(+y[col]||0)-(+x[col]||0)).slice(0,+w.n||5),cc=D.cat>=0?[D.cat,...w.cols]:w.cols;
    return `<div class="db-w table">${tb}<div class="t">${esc(w.title||'Table')}</div><table class="db-tb"><tr>${cc.map(c=>`<th>${esc(hdr[c])}</th>`).join('')}</tr>${srt.map(r=>`<tr>${cc.map(c=>`<td>${isNum(r[c])?fmt(hdr[c],r[c]):esc(r[c])}</td>`).join('')}</tr>`).join('')}</table></div>`}
  const cats=rows.map((r,k)=>D.cat>=0?String(r[D.cat]):String(k+1)),series=w.cols.map(c=>({name:hdr[c],v:rows.map(r=>+r[c]||0)}));
  let svg='';try{svg=PC.draw(w.type,{cats,series},520,260)}catch(e){}
  return `<div class="db-w chart">${tb}<div class="t">${esc(t)}</div><div class="xcb"><svg viewBox="0 0 520 260" style="width:100%;height:auto;display:block">${svg}</svg></div></div>`}
function dash(){const{hdr,rows,all}=ds(),fs=$('#dbF',DO);
  if(D.cat>=0){fs.style.display='';const vals=[...new Set(all.map(r=>String(r[D.cat])))];fs.innerHTML=`<option value="">All ${esc(hdr[D.cat])}</option>`+vals.map(v=>`<option${v===D.filter?' selected':''}>${esc(v)}</option>`).join('')}else fs.style.display='none';
  $('#dbG',DO).innerHTML=D.widgets.map((w,i)=>widget(w,i,hdr,rows)).join('')||'<div class="db-empty">No widgets yet — click “＋ Widget”.</div>'}
const WT=[['kpi','KPI card'],['gauge','Gauge'],['table','Table'],['line','Line chart'],['bar','Bar chart'],['area','Area chart'],['pie','Pie chart'],['doughnut','Doughnut chart'],['scatter','Scatter plot'],['radar','Radar chart'],['histogram','Histogram'],['waterfall','Waterfall'],['combo','Combo chart']];
function editWidget(w){const{hdr}=ds(),nw=!w;w=w||{id:uid('w'),type:'kpi',cols:[],agg:'last',title:''};
  const m=modal(`<h3>${nw?'Add widget':'Edit widget'}</h3><label>Title (optional)</label><input id="wT" value="${esc(w.title)}">
  <label>Type</label><select id="wY">${WT.map(([k,n])=>`<option value="${k}"${w.type===k?' selected':''}>${n}</option>`).join('')}</select>
  <label>Columns (KPI and gauge use the first one)</label><div class="ck">${hdr.map((h,i)=>`<label><input type="checkbox" value="${i}"${w.cols.includes(i)?' checked':''}> ${esc(h)}</label>`).join('')}</div>
  <label>Summary (KPI / gauge)</label><select id="wA">${[['last','Latest row'],['sum','Sum'],['avg','Average'],['max','Max'],['min','Min'],['count','Count']].map(([k,n])=>`<option value="${k}"${w.agg===k?' selected':''}>${n}</option>`).join('')}</select>
  <label>Gauge maximum (optional)</label><input id="wM" type="number" value="${w.max||''}">
  <div class="ac"><button class="xbtn" id="wX">Cancel</button><button class="xbtn g" id="wO">${nw?'Add':'Save'}</button></div>`);
  $('#wX',m).onclick=()=>m.remove();
  $('#wO',m).onclick=()=>{const cols=[...m.querySelectorAll('.ck input:checked')].map(x=>+x.value);if(!cols.length){toast('Pick at least one column');return}
    Object.assign(w,{title:$('#wT',m).value.trim(),type:$('#wY',m).value,cols,agg:$('#wA',m).value,max:$('#wM',m).value});if(nw)D.widgets.push(w);
    m.remove();persist();dash()}}
function openDash(fresh){if(DO)return;const w=WB();
  if(fresh||!w.dashboard){if(!genDash())return}else D=w.dashboard;
  DO=overlay('db',`<div class="db-top"><b id="dbT" contenteditable="true">${esc(D.title)}</b><span class="sp"></span><select id="dbF"></select><button class="xbtn" id="dbAdd">＋ Widget</button><button class="xbtn" id="dbRef">⟳ Refresh</button><button class="xbtn" id="dbClose">✕ Close</button></div><div class="db-body"><div class="db-g" id="dbG"></div></div>`);
  DO.addEventListener('click',e=>{const t=e.target,x=t.closest('[data-x],[data-e]');
    if(t.closest('#dbClose')){DO.remove();DO=null;return}
    if(t.closest('#dbAdd')){editWidget();return}
    if(t.closest('#dbRef')){dash();toast('Dashboard refreshed');return}
    if(x&&x.dataset.x){D.widgets=D.widgets.filter(q=>q.id!==x.dataset.x);persist();dash()}
    else if(x)editWidget(D.widgets.find(q=>q.id===x.dataset.e))});
  DO.addEventListener('change',e=>{if(e.target.id==='dbF'){D.filter=e.target.value;persist();dash()}});
  DO.addEventListener('input',e=>{if(e.target.id==='dbT'){D.title=e.target.textContent;persist()}});
  dash()}

/* ================= ENTRY POINTS ================= */
const cb=$('#chartBtn');
[['Form',openForm],['Dashboard',()=>openDash(false)]].forEach(([n,f])=>{const b=document.createElement('button');b.className='btn';b.textContent=n;b.onclick=f;cb.before(b)});
const ins=document.querySelector('.xl-rib')&&document.querySelector('.xl-rib').children[2];
if(ins){const g=document.createElement('div');g.className='xg';g.innerHTML='<div class="xr"></div><div class="xl">Forms and dashboards</div>';
  [['📝','Form builder',openForm],['📊','Dashboard',()=>openDash(false)],['✨','New dashboard',()=>openDash(true)]].forEach(([i,l,f])=>{
    const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${i}</i><span>${l}</span>`;b.onclick=e=>{e.stopPropagation();f()};g.firstChild.appendChild(b)});
  ins.appendChild(g)}
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
