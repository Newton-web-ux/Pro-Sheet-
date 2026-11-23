/* Pro Sheet — Add-on 2: extra ribbon tabs
   Adds six Excel-style tabs: Page Layout · Review · Analyze · Data Tools · Finance · Automate
   (+ ~50 new formulas: CORREL, SLOPE, FORECAST, PERCENTILE, NPER, RATE, IPMT, EMI, SIP, INRWORDS …)
   Load AFTER excel-pro.js:   <script src="addon2.js"></script>
   Uses only window.ProSheetAPI and the ribbon DOM created by excel-ui.js / excel-pro.js. */
(function(){
"use strict";
let tries=0;
function boot(){
const API=window.ProSheetAPI;
const ready=API&&API.getWorkbook&&API.FN&&API.snapshot&&document.getElementById('xac')&&document.querySelector('.xl-tabs .xt')&&document.querySelectorAll('.xl-rib .xp').length>=5;
if(!ready){ if(++tries<80)return setTimeout(boot,300); console.warn('[addon2] Ribbon not found — load addon2.js after excel-ui.js and excel-pro.js'); return; }
if(window.__ps2Loaded)return; window.__ps2Loaded=true;

/* ================= helpers ================= */
const $=s=>document.querySelector(s),$$=s=>[...document.querySelectorAll(s)];
const L=API.colLetter,COLS=26,WB=()=>API.getWorkbook(),sh=()=>WB().sheets[WB().current],id=(c,r)=>L(c)+(r+1);
const esc=s=>String(s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
const toast=API.toast;
const colIdx=s=>{let n=0;for(const ch of s)n=n*26+ch.charCodeAt(0)-64;return n-1};
const parseRef=t=>{const m=String(t).trim().toUpperCase().match(/^([A-Z]{1,2})(\d+)$/);return m?{c:colIdx(m[1]),r:+m[2]-1}:null};
const parseKey=k=>{const m=k.match(/^([A-Z]+)(\d+)$/);return m?{c:colIdx(m[1]),r:+m[2]-1}:null};
const raw=(c,r)=>{const x=sh().data[id(c,r)];return x&&x.v!==undefined?String(x.v):''};
const cellO=(c,r)=>{const d=sh().data;return d[id(c,r)]||(d[id(c,r)]={v:'',f:{}})};
const sel=()=>{const s=API.getSelection();return{r1:Math.min(s.r1,s.r2),r2:Math.max(s.r1,s.r2),c1:Math.min(s.c1,s.c2),c2:Math.max(s.c1,s.c2)}};
const single=s=>s.r1===s.r2&&s.c1===s.c2;
const region=()=>window.ProSheetCharts&&window.ProSheetCharts.regionSel?window.ProSheetCharts.regionSel():sel();
const target=()=>{const s=sel();return single(s)?region():s};
const data=()=>API.snapshot().sheets[WB().current]||[];
const V=(D,c,r)=>{const v=(D[r]||[])[c];return v===undefined||v===null?'':v};
const isNum=v=>v!==''&&typeof v!=='boolean'&&String(v).trim()!==''&&isFinite(Number(v));
const isF=v=>typeof v==='string'&&v.startsWith('=');
const put=(c,r,v,f)=>{if(c<0||c>=COLS||r<0)return false;if(r>=sh().rows)sh().rows=r+1;const x=cellO(c,r);x.v=String(v);if(f)x.f=Object.assign({},x.f,f);return true};
const eachIn=(s,cb)=>{for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++)cb(c,r)};
const run=(fn,msg)=>{API.pushUndo();fn();API.refresh();if(msg)toast(msg)};
const fmtSel=fn=>{API.pushUndo();eachIn(sel(),(c,r)=>{const x=cellO(c,r);x.f=x.f||{};fn(x.f)});API.refresh()};
const tdAt=(r,c)=>document.querySelector(`#grid td.cell[data-r="${r}"][data-c="${c}"]`);
const dlg=(h,w)=>{const b=$('#dialogBox');b.style.width=w?w+'px':'';b.innerHTML=h;$('#dialogOverlay').classList.add('show')};
const close=()=>$('#dialogOverlay').classList.remove('show');
new MutationObserver(()=>{if(!$('#dialogOverlay').classList.contains('show'))$('#dialogBox').style.width=''}).observe($('#dialogOverlay'),{attributes:true,attributeFilter:['class']});
const go=(c,r)=>{const b=$('#cellref');b.value=id(c,r);b.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'}))};
const iso=d=>d.toISOString().slice(0,10);
const dt=v=>{const d=new Date(v);return isNaN(d)?null:d};
const f4=v=>typeof v==='number'?String(+v.toFixed(4)):String(v);
const inr=n=>'₹'+Math.round(n).toLocaleString('en-IN');
const download=(content,name,type)=>{const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([content],{type}));a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),500);toast('Downloaded '+name)};
const copyText=(t,msg)=>{(navigator.clipboard?navigator.clipboard.writeText(t):Promise.reject()).then(()=>toast(msg||'Copied to clipboard'),()=>{dlg(`<h3>Copy manually</h3><textarea rows="8" style="user-select:text">${esc(t)}</textarea><div class="dialog-actions"><button class="btn primary" id="p2x">Close</button></div>`,460);$('#p2x').onclick=close})};
const tableHtml=(head,rows)=>`<table class="p2t">${head?`<tr>${head.map(h=>`<th>${esc(h)}</th>`).join('')}</tr>`:''}${rows.map(r=>`<tr>${r.map(x=>`<td>${esc(x)}</td>`).join('')}</tr>`).join('')}</table>`;

/* generic small form: fields = [{k,label,type,def,opts}]; opts items are "text" or [value,label]. Dialog closes, then onOk(values) runs. */
function form(title,fields,onOk,ok,w){
  const body=fields.map(f=>`<label>${esc(f.label)}</label>`+(f.opts?`<select data-k="${f.k}">${f.opts.map(o=>{const[v,l]=Array.isArray(o)?o:[o,o];return`<option value="${esc(v)}"${String(v)===String(f.def)?' selected':''}>${esc(l)}</option>`}).join('')}</select>`:`<input data-k="${f.k}" type="${f.type||'text'}" step="any" value="${esc(f.def==null?'':f.def)}">`)).join('');
  dlg(`<h3>${esc(title)}</h3>${body}<div class="dialog-actions"><button class="btn" id="p2c">Cancel</button><button class="btn primary" id="p2o">${ok||'OK'}</button></div>`,w);
  $('#p2c').onclick=close;
  $('#p2o').onclick=()=>{const v={};$$('#dialogBox [data-k]').forEach(e=>v[e.dataset.k]=e.value);close();onOk(v)};
  const first=$('#dialogBox [data-k]');if(first&&first.focus)first.focus();
}
/* live calculator dialog */
function calcDlg(title,fields,compute,sheetFn){
  dlg(`<h3>${esc(title)}</h3>`+fields.map(f=>`<label>${esc(f[1])}</label><input data-k="${f[0]}" type="number" step="any" value="${f[2]}">`).join('')+`<div id="p2co" style="margin-top:12px;font-size:12.5px;line-height:1.9"></div><div class="dialog-actions"><button class="btn" id="p2cc">Close</button>${sheetFn?'<button class="btn primary" id="p2cs">Create sheet</button>':''}</div>`);
  const get=()=>{const v={};$$('#dialogBox [data-k]').forEach(e=>v[e.dataset.k]=+e.value);return v};
  const draw=()=>{$('#p2co').innerHTML=compute(get()).map(([k,v,b])=>`<div style="display:flex;justify-content:space-between;${b?'font-weight:700;color:var(--accent);':''}"><span>${k}</span><span>${v}</span></div>`).join('')};
  $$('#dialogBox [data-k]').forEach(e=>e.oninput=draw);draw();
  $('#p2cc').onclick=close;if(sheetFn)$('#p2cs').onclick=()=>{const v=get();close();sheetFn(v)};
}
function newSheet(name){const w=WB();let n=name,k=2;while(w.sheets[n])n=name+' '+k++;w.sheets[n]={data:{},colWidths:{},rowHeights:{},rows:100,merges:[],frozenRow:0,frozenCol:0};w.order.push(n);w.current=n;return n}
function toSheet(name,head,body,money){money=money||[];API.pushUndo();newSheet(name);
  head.forEach((h,i)=>{put(i,0,h,{bold:true,bg:'#2b3a6b',color:'#ffffff',align:'center'});sh().colWidths[i]=110});
  body.forEach((row,ri)=>row.forEach((v,ci)=>put(ci,ri+1,v,money.includes(ci)?{numFmt:'currency'}:undefined)));API.refresh()}

/* keep typing / pasting inside dialogs away from the grid's global handlers */
document.addEventListener('keydown',e=>{if(e.target&&e.target.closest&&e.target.closest('#dialogBox'))e.stopImmediatePropagation()},true);
document.addEventListener('paste',e=>{if(e.target&&e.target.closest&&e.target.closest('#dialogBox'))e.stopImmediatePropagation()},true);

/* ================= styles ================= */
const css=document.createElement('style');
css.textContent=`
.xl-rib .xp.ps2p{display:none}
body.ps2on .xl-rib .xp:not(.ps2p){display:none!important}
body.ps2on .xl-rib .xp.ps2p.cur{display:flex!important;flex-wrap:wrap;align-items:stretch}
#tabsList .tab[data-ps2lock]::after{content:' 🔒';font-size:10px}
.ps2np{display:none!important}
#p2dvb{position:fixed;display:none;width:16px;align-items:center;justify-content:center;background:var(--surface-3);border:1px solid var(--border);color:var(--text);font-size:8px;cursor:pointer;z-index:60;border-radius:2px}
#p2dvm{position:fixed;display:none;z-index:710;min-width:130px;max-height:220px;overflow:auto;background:var(--surface-2);border:1px solid var(--border);border-radius:6px;padding:4px;box-shadow:0 12px 32px rgba(0,0,0,.45)}
#p2dvm div{padding:5px 9px;font-size:12.5px;cursor:pointer;border-radius:4px;color:var(--text)}
#p2dvm div:hover{background:var(--accent-soft);color:var(--accent)}
.p2t{width:100%;border-collapse:collapse;font-size:12px;color:var(--text)}
.p2t td,.p2t th{padding:4px 7px;border-bottom:1px solid var(--border);text-align:left}
.p2t th{color:var(--text-dim);font-weight:600}
.p2bar{height:14px;background:linear-gradient(90deg,var(--accent),#9b6cff);border-radius:2px}
.p2li{padding:6px 9px;font-size:12.5px;display:flex;justify-content:space-between;gap:8px;align-items:center;cursor:pointer;color:var(--text)}
.p2li:hover{background:var(--accent-soft)}
@page{margin:15mm}
@media print{
 .topbar,.toolbar,.formulabar,.sheettabs,.xl-tabs,.xl-rib,.ctxmenu,#xac,#p2dvb,#p2dvm,.toast,.colhead,.rowhead,.corner,.xaf,.fill-handle{display:none!important}
 html,body{height:auto!important;overflow:visible!important;background:#fff!important}
 body{display:block!important}
 .gridwrap{overflow:visible!important;height:auto!important;flex:none!important}
 td.cell.active-cell,td.cell.selected-range{outline:none!important}
}`;
document.head.appendChild(css);
const printCss=document.createElement('style');printCss.id='p2PrintCss';document.head.appendChild(printCss);

/* ================= new formulas ================= */
const F=API.FN,fl=a=>[a].flat(Infinity),nums=a=>fl(a).map(parseFloat).filter(x=>!isNaN(x));
const mean=a=>a.reduce((s,v)=>s+v,0)/a.length;
const pairs=(x,y)=>{const a=fl(x),b=fl(y),o=[];for(let i=0;i<Math.min(a.length,b.length);i++){const p=parseFloat(a[i]),q=parseFloat(b[i]);if(!isNaN(p)&&!isNaN(q))o.push([p,q])}return o};
const lin=(x,y)=>{const p=pairs(x,y);if(p.length<2)return null;const mx=mean(p.map(t=>t[0])),my=mean(p.map(t=>t[1]));let sxy=0,sxx=0,syy=0;p.forEach(([a,b])=>{sxy+=(a-mx)*(b-my);sxx+=(a-mx)**2;syy+=(b-my)**2});if(!sxx||!syy)return null;return{slope:sxy/sxx,int:my-sxy/sxx*mx,r:sxy/Math.sqrt(sxx*syy),n:p.length}};
const pct=(arr,k)=>{const s=[...arr].sort((a,b)=>a-b);if(!s.length)return NaN;const p=(s.length-1)*k,lo=Math.floor(p),hi=Math.ceil(p);return s[lo]+(s[hi]-s[lo])*(p-lo)};
const gcd=(x,y)=>{x=Math.abs(x);y=Math.abs(y);while(y){[x,y]=[y,x%y]}return x};
const ONES=['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'],TENS=['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
const w99=n=>n<20?ONES[n]:TENS[Math.floor(n/10)]+(n%10?' '+ONES[n%10]:'');
const w999=n=>(n>=100?ONES[Math.floor(n/100)]+' Hundred'+(n%100?' ':''):'')+(n%100?w99(n%100):'');
function wordsInt(n){if(n===0)return'Zero';const p=[];[[1e7,'Crore'],[1e5,'Lakh'],[1e3,'Thousand']].forEach(([d,nm])=>{const q=Math.floor(n/d);if(q){p.push((d===1e7?wordsInt(q):w999(q))+' '+nm);n%=d}});if(n)p.push(w999(n));return p.join(' ')}
Object.assign(F,{
CORREL:(a,b)=>{const l=lin(a,b);return l?l.r:'#DIV/0!'},
SLOPE:(y,x)=>{const l=lin(x,y);return l?l.slope:'#DIV/0!'},
INTERCEPT:(y,x)=>{const l=lin(x,y);return l?l.int:'#DIV/0!'},
RSQ:(y,x)=>{const l=lin(x,y);return l?l.r*l.r:'#DIV/0!'},
FORECAST:(x0,y,x)=>{const l=lin(x,y);return l?l.int+l.slope*parseFloat(x0):'#DIV/0!'},
PERCENTILE:(a,k)=>pct(nums(a),k),QUARTILE:(a,q)=>pct(nums(a),q/4),
GEOMEAN:(...a)=>{const n=nums(a);return n.length&&n.every(v=>v>0)?Math.exp(mean(n.map(v=>Math.log(v)))):'#NUM!'},
FACT:n=>{n=Math.floor(n);if(n<0||n>170)return'#NUM!';let r=1;for(let i=2;i<=n;i++)r*=i;return r},
COMBIN:(n,k)=>{if(k<0||k>n)return'#NUM!';let r=1;for(let i=1;i<=k;i++)r=r*(n-k+i)/i;return Math.round(r)},
PERMUT:(n,k)=>{let r=1;for(let i=0;i<k;i++)r*=n-i;return r},
GCD:(...a)=>nums(a).reduce((x,y)=>gcd(x,y),0),LCM:(...a)=>nums(a).reduce((x,y)=>x&&y?Math.abs(x*y)/gcd(x,y):0,1),
MROUND:(n,m)=>m?Math.round(n/m)*m:0,QUOTIENT:(a,b)=>Math.trunc(a/b),
DEGREES:x=>x*180/Math.PI,RADIANS:x=>x*Math.PI/180,SIN:Math.sin,COS:Math.cos,TAN:Math.tan,ASIN:Math.asin,ACOS:Math.acos,ATAN:Math.atan,
ISEVEN:n=>Math.floor(Number(n))%2===0,ISODD:n=>Math.abs(Math.floor(Number(n)))%2===1,
XOR:(...a)=>fl(a).filter(Boolean).length%2===1,
SWITCH:(x,...a)=>{for(let i=0;i+1<a.length;i+=2)if(String(a[i])===String(x))return a[i+1];return a.length%2?a[a.length-1]:'#N/A'},
ISEMAIL:s=>/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s)),
TEXTBEFORE:(s,d,n=1)=>{s=String(s);d=String(d);let i=-1;for(let k=0;k<n;k++){i=s.indexOf(d,i+1);if(i<0)return'#N/A'}return s.slice(0,i)},
TEXTAFTER:(s,d,n=1)=>{s=String(s);d=String(d);let i=-1;for(let k=0;k<n;k++){i=s.indexOf(d,i+1);if(i<0)return'#N/A'}return s.slice(i+d.length)},
REVERSETEXT:s=>[...String(s)].reverse().join(''),
WORDCOUNT:s=>String(s).trim()?String(s).trim().split(/\s+/).length:0,
DAYS:(e,s)=>{const a=dt(e),b=dt(s);return a&&b?Math.round((a-b)/864e5):'#VALUE!'},
WORKDAY:(s,n)=>{const d=dt(s);if(!d)return'#VALUE!';const st=n<0?-1:1;let k=Math.abs(n);while(k>0){d.setUTCDate(d.getUTCDate()+st);const w=d.getUTCDay();if(w!==0&&w!==6)k--}return iso(d)},
WEEKNUM:v=>{const d=dt(v);if(!d)return'#VALUE!';const j=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.floor(((d-j)/864e5+j.getUTCDay())/7)+1},
YEARFRAC:(a,b)=>{const x=dt(a),y=dt(b);return x&&y?Math.abs(y-x)/864e5/365:'#VALUE!'},
AGE:v=>{const d=dt(v);if(!d)return'#VALUE!';const n=new Date();let a=n.getFullYear()-d.getUTCFullYear();const m=n.getMonth()-d.getUTCMonth();if(m<0||(m===0&&n.getDate()<d.getUTCDate()))a--;return a},
/* timezone-safe replacements (the originals can shift a day for users east of UTC, e.g. India) */
DATE:(y,m,d)=>iso(new Date(Date.UTC(y,m-1,d))),
EOMONTH:(v,m)=>{const d=dt(v);return d?iso(new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+parseInt(m,10)+1,0))):'#VALUE!'},
/* finance */
NPER:(r,p,pv,fv=0,t=0)=>r===0?-(pv+fv)/p:Math.log((p*(1+r*t)-fv*r)/(p*(1+r*t)+pv*r))/Math.log(1+r),
RATE:(n,p,pv,fv=0,t=0,g=0.1)=>{const f=r=>Math.abs(r)<1e-12?pv+p*n+fv:pv*Math.pow(1+r,n)+p*(1+r*t)*(Math.pow(1+r,n)-1)/r+fv;let r=g;for(let i=0;i<100;i++){const y=f(r),d=(f(r+1e-7)-y)/1e-7;if(!d||!isFinite(d))return'#NUM!';const nx=r-y/d;if(Math.abs(nx-r)<1e-10)return nx;r=nx}return'#NUM!'},
IPMT:(r,per,n,pv,fv=0,t=0)=>{if(!r)return 0;const p=F.PMT(r,n,pv,fv,t),B=k=>pv*Math.pow(1+r,k)+p*(Math.pow(1+r,k)-1)/r;return -B(per-1)*r},
PPMT:(r,per,n,pv,fv=0,t=0)=>F.PMT(r,n,pv,fv,t)-F.IPMT(r,per,n,pv,fv,t),
SLN:(c,s,l)=>(c-s)/l,CAGR:(b,e,y)=>Math.pow(e/b,1/y)-1,
XNPV:(r,v,d)=>{const a=nums(v),b=fl(d).map(dt);if(!a.length||b.some(x=>!x))return'#VALUE!';return a.reduce((t,x,i)=>t+x/Math.pow(1+r,(b[i]-b[0])/864e5/365),0)},
EMI:(P,R,N)=>{P=+P;R=+R;N=+N;if(!(N>0)||!(P>=0))return'#NUM!';const r=R/1200;return r?P*r*Math.pow(1+r,N)/(Math.pow(1+r,N)-1):P/N},
SIP:(M,R,Y)=>{const n=Math.round(Y*12),r=R/1200;return r?M*((Math.pow(1+r,n)-1)/r)*(1+r):M*n},
GST:(a,rate)=>a*rate/100,
INRWORDS:v=>{const x=parseFloat(v);if(isNaN(x))return'#VALUE!';const n=Math.round(Math.abs(x)*100),rup=Math.floor(n/100),ps=n%100;return(x<0?'Minus ':'')+'Rupees '+wordsInt(rup)+(ps?' and '+wordsInt(ps)+' Paise':'')+' Only'}
});

/* ================= display decorations (Indian format, show-formulas, validation) ================= */
let showF=false;
const bound=(x,d)=>x===''||x==null?d:Number(x);
function dvOk(R,v){v=String(v).trim();if(v==='')return true;
  if(R.type==='list')return R.list.some(x=>x.toLowerCase()===v.toLowerCase());
  const n=Number(v),lo=bound(R.min,-Infinity),hi=bound(R.max,Infinity);
  if(R.type==='whole')return Number.isInteger(n)&&n>=lo&&n<=hi;
  if(R.type==='decimal')return isFinite(n)&&n>=lo&&n<=hi;
  if(R.type==='length')return v.length>=lo&&v.length<=hi;
  return true}
function fmtInr(n,x){const d=x.d==null?2:x.d,s=x.sym||'';let v=n,suf='';if(x.unit==='L'){v=n/1e5;suf=' L'}else if(x.unit==='Cr'){v=n/1e7;suf=' Cr'}
  return(v<0?'-':'')+s+Math.abs(v).toLocaleString('en-IN',{minimumFractionDigits:d,maximumFractionDigits:d})+suf}
let busy=false,raf=0;
function decorate(){const all=$$('#grid td.cell');if(!all.length||all.every(t=>t.dataset.p2))return;busy=true;
  const s=sh(),dv=s.dv||[];
  all.forEach(td=>{const r=+td.dataset.r,c=+td.dataset.c,cell=s.data[id(c,r)],f=cell&&cell.f,cv=td.querySelector('.cv');
    if(cv&&f&&f.inr){const t=cv.textContent.trim().replace(/,/g,'');if(/^-?\d+(\.\d+)?(e[+-]?\d+)?$/i.test(t))cv.textContent=fmtInr(parseFloat(t),f.inr)}
    if(cv&&showF&&cell&&isF(cell.v))cv.textContent=cell.v;
    if(cell&&!isF(cell.v)&&cell.v!==''){for(const R of dv){if(r>=R.r1&&r<=R.r2&&c>=R.c1&&c<=R.c2&&!dvOk(R,cell.v)){td.style.boxShadow='inset 0 0 0 2px #ff6b6b';td.title='Invalid entry for this cell\'s validation rule';break}}}});
  all.forEach(t=>t.dataset.p2='1');obs.takeRecords();busy=false;setTimeout(updateDv,0)}
const obs=new MutationObserver(()=>{if(busy)return;cancelAnimationFrame(raf);raf=requestAnimationFrame(decorate)});
obs.observe($('#grid'),{childList:true,subtree:true});

/* data-validation dropdown arrow */
const dvb=document.createElement('div');dvb.id='p2dvb';dvb.textContent='▼';document.body.appendChild(dvb);
const dvm=document.createElement('div');dvm.id='p2dvm';document.body.appendChild(dvm);
function updateDv(){const a=API.getActive(),rule=(sh().dv||[]).find(R=>R.type==='list'&&a.r>=R.r1&&a.r<=R.r2&&a.c>=R.c1&&a.c<=R.c2),td=tdAt(a.r,a.c);dvm.style.display='none';
  if(!rule||!td){dvb.style.display='none';return}
  const b=td.getBoundingClientRect(),w=$('#gridwrap').getBoundingClientRect();
  if(b.bottom<w.top||b.top>w.bottom||b.right<w.left||b.left>w.right){dvb.style.display='none';return}
  dvb.style.display='flex';dvb.style.left=(b.right-18)+'px';dvb.style.top=(b.top+2)+'px';dvb.style.height=(b.height-4)+'px';dvb._rule=rule}
dvb.onmousedown=e=>{e.preventDefault();e.stopPropagation()};
dvb.onclick=e=>{e.stopPropagation();const R=dvb._rule;if(!R)return;const b=dvb.getBoundingClientRect();
  dvm.innerHTML=R.list.map(x=>`<div>${esc(x)}</div>`).join('');dvm.style.left=(b.right-130)+'px';dvm.style.top=(b.bottom+2)+'px';dvm.style.display='block';
  [...dvm.children].forEach((d,i)=>d.onclick=()=>{const a=API.getActive();if(sh().protect&&lockedAt(a.c,a.r))return warn();API.pushUndo();cellO(a.c,a.r).v=R.list[i];dvm.style.display='none';API.refresh()})};
document.addEventListener('mouseup',()=>setTimeout(updateDv,0));document.addEventListener('keyup',()=>setTimeout(updateDv,0));
$('#gridwrap').addEventListener('scroll',updateDv);
document.addEventListener('click',e=>{if(!e.target.closest('#p2dvm'))dvm.style.display='none'});

/* ================= ribbon plumbing ================= */
const tabsBar=$('.xl-tabs');
const actCls=(()=>{const t=$$('.xl-tabs .xt').find(x=>[...x.classList].some(c=>c!=='xt'));return t?[...t.classList].find(c=>c!=='xt'):'on'})();
const mine=[];
function panelFor(name){
  const tabs=$$('.xl-tabs .xt'),i=tabs.findIndex(t=>t.textContent.trim()===name),panels=$$('.xl-rib .xp');
  if(i>=0&&panels[i]&&!tabs[i].classList.contains('ps2t'))return panels[i];
  const ex=mine.find(m=>m.name===name);if(ex)return ex.panel;
  const t=document.createElement('div');t.className='xt ps2t';t.textContent=name;
  const p=document.createElement('div');p.className='xp ps2p';
  tabsBar.appendChild(t);$('.xl-rib').appendChild(p);
  t.onclick=e=>{e.stopPropagation();$$('.xl-tabs .xt').forEach(x=>x.classList.remove(actCls));t.classList.add(actCls);$$('.xl-rib .xp').forEach(x=>x.classList.remove('cur'));p.classList.add('cur');document.body.classList.add('ps2on')};
  mine.push({name,tab:t,panel:p});return p}
tabsBar.addEventListener('click',e=>{const t=e.target.closest('.xt');if(t&&!t.classList.contains('ps2t')){document.body.classList.remove('ps2on');mine.forEach(m=>m.tab.classList.remove(actCls))}},true);
const mkb=(i,l,fn)=>{const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${i}</i><span>${l}</span>`;b.onclick=e=>{e.stopPropagation();fn()};return b};
const mks=(i,l,o,fn)=>{const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${i}</i><select><option value="">${l}</option>${o.map(x=>`<option>${esc(x)}</option>`).join('')}</select>`;const s=b.querySelector('select');s.onchange=()=>{if(s.value)fn(s.value);s.value=''};return b};
const grp=(panelName,label,items)=>{const g=document.createElement('div');g.className='xg';const r=document.createElement('div');r.className='xr';items.forEach(x=>r.appendChild(x));g.appendChild(r);g.insertAdjacentHTML('beforeend',`<div class="xl">${label}</div>`);panelFor(panelName).appendChild(g)};

/* ================= PAGE LAYOUT ================= */
const ACC={Indigo:['#6c8cff','#3a4a8a','108,140,255'],Emerald:['#34d399','#1f6b52','52,211,153'],Rose:['#fb7185','#8a3a4a','251,113,133'],Amber:['#ffb454','#8a6020','255,180,84'],Cyan:['#4dd0e1','#1f6b78','77,208,225'],Violet:['#a78bfa','#523a8a','167,139,250']};
function setAccent(n){const a=ACC[n];if(!a)return;const s=document.documentElement.style;s.setProperty('--accent',a[0]);s.setProperty('--accent-dim',a[1]);s.setProperty('--accent-soft',`rgba(${a[2]},.14)`);try{localStorage.setItem('prosheet_accent',n)}catch(e){}}
try{const sa=localStorage.getItem('prosheet_accent');if(sa)setAccent(sa)}catch(e){}
function toggleTheme(){const l=!document.documentElement.classList.contains('ps-light');document.documentElement.classList.toggle('ps-light',l);try{localStorage.setItem('prosheet_theme',l?'light':'dark')}catch(e){}}
const pg=()=>WB().page=WB().page||{o:'Portrait',m:'Normal',p:'A4',s:100};
const MARG={Normal:'15mm',Narrow:'7mm',Wide:'25mm'};
const pageCss=()=>{const p=pg();return`@page{size:${p.p} ${p.o.toLowerCase()};margin:${MARG[p.m]}}@media print{#gridwrap{zoom:${p.s/100}!important}}`};
const setPage=(k,v)=>{pg()[k]=v;printCss.textContent=pageCss();API.refresh();toast('Page setup updated')};
printCss.textContent=pageCss();
function usedRange(){let R=0,C=0;Object.keys(sh().data).forEach(k=>{const x=sh().data[k];if(x&&x.v!==''&&x.v!==undefined){const p=parseKey(k);if(p){R=Math.max(R,p.r);C=Math.max(C,p.c)}}});return{r1:0,c1:0,r2:R,c2:C}}
let printMarks=[],printWasDark=false;
window.addEventListener('beforeprint',()=>{printCss.textContent=pageCss();const a=sh().printArea||usedRange();
  printWasDark=!document.documentElement.classList.contains('ps-light');if(printWasDark)document.documentElement.classList.add('ps-light');
  $$('#grid tr').forEach(tr=>{const rh=tr.querySelector('.rowhead');if(rh){const r=+rh.dataset.r;if(r<a.r1||r>a.r2){tr.classList.add('ps2np');printMarks.push(tr);return}}
    tr.querySelectorAll('td.cell,th.colhead').forEach(e=>{const c=+e.dataset.c;if(c<a.c1||c>a.c2){e.classList.add('ps2np');printMarks.push(e)}})})});
window.addEventListener('afterprint',()=>{printMarks.forEach(e=>e.classList.remove('ps2np'));printMarks=[];if(printWasDark)document.documentElement.classList.remove('ps-light')});
let zoom=1;const setZoom=z=>{zoom=Math.max(.5,Math.min(2,z));$('#gridwrap').style.zoom=zoom;toast('Zoom '+Math.round(zoom*100)+'%')};
const TABC={None:'',Red:'#ef4444',Orange:'#f97316',Yellow:'#eab308',Green:'#22c55e',Teal:'#14b8a6',Blue:'#3b82f6',Purple:'#a855f7',Pink:'#ec4899'};
const visible=()=>WB().order.filter(n=>!WB().sheets[n].hidden);
function moveSheet(d){const w=WB(),i=w.order.indexOf(w.current),j=i+d;if(j<0||j>=w.order.length)return;[w.order[i],w.order[j]]=[w.order[j],w.order[i]];API.refresh()}
function hideSheet(){const w=WB();if(visible().length<=1)return toast("Can't hide the only visible sheet");const cur=w.current;w.sheets[cur].hidden=true;w.current=visible()[0];API.refresh();toast(`Sheet "${cur}" hidden`)}
function unhideSheet(){const h=WB().order.filter(n=>WB().sheets[n].hidden);if(!h.length)return toast('No hidden sheets');
  dlg(`<h3>Unhide sheet</h3><div class="xfl">${h.map(n=>`<div class="p2li" data-n="${esc(n)}">${esc(n)}</div>`).join('')}</div><div class="dialog-actions"><button class="btn" id="p2x">Close</button></div>`);
  $('#p2x').onclick=close;$$('#dialogBox .p2li').forEach(d=>d.onclick=()=>{WB().sheets[d.dataset.n].hidden=false;close();API.refresh()})}
function decTabs(){const w=WB();$$('#tabsList .tab').forEach(t=>{const s=w.sheets[t.textContent];if(!s)return;t.style.boxShadow=s.tabColor?`inset 0 -3px 0 ${s.tabColor}`:'';t.style.display=s.hidden?'none':'';t.toggleAttribute('data-ps2lock',!!s.protect)})}
new MutationObserver(decTabs).observe($('#tabsList'),{childList:true});

/* ================= REVIEW ================= */
let spell=false;
const hash=s=>{let h=5381;for(const ch of String(s))h=((h<<5)+h+ch.charCodeAt(0))|0;return String(h)};
const lockedAt=(c,r)=>{const x=sh().data[id(c,r)];return !!sh().protect&&!(x&&x.f&&x.f.unlocked)};
const anyLocked=s=>{if(!sh().protect)return false;for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++)if(lockedAt(c,r))return true;return false};
const warn=()=>toast('🔒 Locked cell — unprotect the sheet to edit it');
function protectDlg(){form('Protect sheet',[{k:'p',label:'Password (optional — stops casual edits, not real security)',type:'password'}],v=>{const s=sh();s.protect=true;s.protectHash=v.p?hash(v.p):'';API.refresh();toast('🔒 Sheet protected. Cells marked "unlocked" stay editable.')},'Protect')}
function unprotect(){const s=sh();if(!s.protect)return toast('This sheet is not protected');const done=()=>{s.protect=false;s.protectHash='';API.refresh();toast('Sheet unprotected')};if(!s.protectHash)return done();
  form('Unprotect sheet',[{k:'p',label:'Password',type:'password'}],v=>{hash(v.p)===s.protectHash?done():toast('Wrong password')},'Unprotect')}
function lockCells(unlock){if(sh().protect)return toast('Unprotect the sheet first');fmtSel(f=>{if(unlock)f.unlocked=true;else delete f.unlocked});toast(unlock?'Cells unlocked (editable when protected)':'Cells locked')}
/* block typing / delete / paste / fill / formula-bar editing on locked cells */
document.addEventListener('keydown',e=>{if(!sh().protect)return;const a=document.activeElement;if(a&&(a.tagName==='INPUT'||a.tagName==='SELECT'||a.tagName==='TEXTAREA'||a.isContentEditable))return;
  const ctrl=e.ctrlKey||e.metaKey,k=e.key.toLowerCase();let bad=false;
  if(ctrl){if(['v','x','d'].includes(k)||e.key===';')bad=anyLocked(sel())}
  else if(e.key==='Delete'||e.key==='Backspace')bad=anyLocked(sel());
  else if(e.key==='Enter'||e.key==='F2'||(e.key.length===1&&!e.altKey)){const ac=API.getActive();bad=lockedAt(ac.c,ac.r)}
  if(bad){e.preventDefault();e.stopImmediatePropagation();warn()}},true);
document.addEventListener('dblclick',e=>{const td=e.target.closest&&e.target.closest('#grid td.cell');if(td&&lockedAt(+td.dataset.c,+td.dataset.r)){e.stopImmediatePropagation();e.preventDefault();warn()}},true);
document.addEventListener('focus',e=>{if(e.target.id==='formulaInput'){const a=API.getActive();if(lockedAt(a.c,a.r)){e.stopImmediatePropagation();e.target.blur();warn()}}},true);
document.addEventListener('paste',e=>{if(sh().protect&&!(e.target.closest&&e.target.closest('#dialogBox'))&&anyLocked(sel())){e.stopImmediatePropagation();e.preventDefault();warn()}},true);
document.addEventListener('mousedown',e=>{if(sh().protect&&e.target.closest&&e.target.closest('.fill-handle')){e.stopImmediatePropagation();e.preventDefault();warn()}},true);

const VK=()=>'prosheet_ver_'+API.getDocId();
const vers=()=>{try{return JSON.parse(localStorage.getItem(VK()))||[]}catch(e){return[]}};
const saveVers=v=>{try{localStorage.setItem(VK(),JSON.stringify(v));return true}catch(e){toast('Browser storage is full — delete old versions');return false}};
function saveVersion(name){const l=vers();l.unshift({t:Date.now(),n:name||'Version',j:JSON.stringify(WB())});return saveVers(l.slice(0,15))}
function histDlg(){const l=vers();
  dlg(`<h3>Version history</h3>${l.length?`<div class="xfl" style="max-height:260px">${l.map((v,i)=>`<div style="cursor:default"><span><b>${esc(v.n)}</b><br><small style="color:var(--text-faint)">${new Date(v.t).toLocaleString()}</small></span><span><button class="btn" data-r="${i}">Restore</button> <button class="btn" data-d="${i}" style="color:var(--danger)">✕</button></span></div>`).join('')}</div>`:'<div style="color:var(--text-dim);font-size:12.5px">No saved versions yet. Use “Save version” first.</div>'}<div class="dialog-actions"><button class="btn" id="p2x">Close</button></div>`,440);
  $('#p2x').onclick=close;
  $$('#dialogBox [data-d]').forEach(b=>b.onclick=()=>{const l2=vers();l2.splice(+b.dataset.d,1);saveVers(l2);histDlg()});
  $$('#dialogBox [data-r]').forEach(b=>b.onclick=()=>{const v=vers()[+b.dataset.r];if(!v)return;let p;try{p=JSON.parse(v.j)}catch(e){return toast('Version is corrupted')}
    saveVersion('Before restore');const w=WB();API.pushUndo();Object.keys(w).forEach(k=>delete w[k]);Object.assign(w,p);if(!w.sheets[w.current])w.current=w.order[0];close();API.refresh();toast('Restored “'+v.n+'”')})}
function statsDlg(){const w=WB();const rows=w.order.map(n=>{const s=w.sheets[n];let cells=0,fm=0,numc=0,words=0;Object.values(s.data).forEach(x=>{if(!x||x.v===''||x.v===undefined)return;cells++;if(isF(x.v))fm++;else if(isNum(x.v))numc++;else words+=String(x.v).trim().split(/\s+/).length});return[n,cells,fm,numc,words,Object.keys(s.notes||{}).length]});
  dlg(`<h3>Workbook statistics</h3>${tableHtml(['Sheet','Cells','Formulas','Numbers','Words','Notes'],rows)}<div class="dialog-actions"><button class="btn primary" id="p2x">Close</button></div>`,460);$('#p2x').onclick=close}
function notesList(){const n=sh().notes||{},k=Object.keys(n);
  dlg(`<h3>Notes on this sheet</h3>${k.length?`<div class="xfl">${k.map(x=>`<div class="p2li" data-k="${x}"><b>${x}</b><span style="opacity:.75;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(n[x])}</span></div>`).join('')}</div>`:'<div style="color:var(--text-dim);font-size:12.5px">No notes yet. Use Home → New Note.</div>'}<div class="dialog-actions"><button class="btn" id="p2x">Close</button></div>`,420);
  $('#p2x').onclick=close;$$('#dialogBox .p2li').forEach(d=>d.onclick=()=>{close();const p=parseRef(d.dataset.k);if(p)go(p.c,p.r)})}
function errDlg(){const D=data(),errs=[];D.forEach((row,r)=>row.forEach((v,c)=>{if(typeof v==='string'&&/^#[A-Z/0!?]+!?$/.test(v))errs.push([c,r,v])}));
  dlg(`<h3>Error check</h3>${errs.length?`<div class="xfl">${errs.map(([c,r,v])=>`<div class="p2li" data-c="${c}" data-r="${r}"><b>${id(c,r)}</b><span style="color:var(--danger)">${esc(v)}</span></div>`).join('')}</div>`:'<div style="color:var(--success);font-size:13px">✔ No formula errors found on this sheet.</div>'}<div class="dialog-actions"><button class="btn" id="p2x">Close</button></div>`,380);
  $('#p2x').onclick=close;$$('#dialogBox .p2li').forEach(d=>d.onclick=()=>{close();go(+d.dataset.c,+d.dataset.r)})}

/* ================= ANALYZE ================= */
function numsIn(s){const D=data(),a=[];eachIn(s,(c,r)=>{const v=V(D,c,r);if(isNum(v))a.push(+v)});return a}
function describe(a){const n=a.length,m=mean(a),s=[...a].sort((x,y)=>x-y),sd=n>1?Math.sqrt(a.reduce((t,v)=>t+(v-m)**2,0)/(n-1)):0;
  const cnt={};a.forEach(v=>cnt[v]=(cnt[v]||0)+1);let mode='—',best=1;Object.keys(cnt).forEach(k=>{if(cnt[k]>best){best=cnt[k];mode=k}});
  const sk=sd&&n>2?a.reduce((t,v)=>t+((v-m)/sd)**3,0)*n/((n-1)*(n-2)):0;
  return[['Count',n],['Sum',a.reduce((x,y)=>x+y,0)],['Mean',m],['Median',pct(a,.5)],['Mode',mode],['Std deviation (sample)',sd],['Variance',sd*sd],['Minimum',s[0]],['Quartile 1',pct(a,.25)],['Quartile 3',pct(a,.75)],['Maximum',s[n-1]],['Range',s[n-1]-s[0]],['IQR',pct(a,.75)-pct(a,.25)],['Skewness',sk],['Coefficient of variation',m?sd/m:'—']].map(([k,v])=>[k,typeof v==='number'?+v.toFixed(4):v])}
function describeDlg(){const a=numsIn(target());if(a.length<2)return toast('Select a range with numbers');const rows=describe(a);
  dlg(`<h3>Descriptive statistics</h3>${tableHtml(null,rows)}<div class="dialog-actions"><button class="btn" id="p2x">Close</button><button class="btn primary" id="p2s">Send to new sheet</button></div>`,380);
  $('#p2x').onclick=close;$('#p2s').onclick=()=>{close();toSheet('Statistics',['Statistic','Value'],rows)}}
function histogram(){const s=target(),a=numsIn(s);if(a.length<2)return toast('Select a range with numbers');
  form('Histogram',[{k:'b',label:'Number of bins',type:'number',def:8}],v=>{const k=Math.max(2,Math.min(30,Math.round(+v.b)||8)),mn=Math.min(...a),mx=Math.max(...a),w=(mx-mn)/k||1,bins=[...Array(k)].map((_,i)=>({lo:mn+i*w,hi:mn+(i+1)*w,n:0}));
    a.forEach(x=>{bins[Math.min(k-1,Math.floor((x-mn)/w))].n++});const top=Math.max(...bins.map(b=>b.n))||1;
    dlg(`<h3>Histogram</h3><div style="max-height:300px;overflow:auto">${bins.map(b=>`<div style="display:flex;gap:8px;align-items:center;font-size:11.5px;margin:3px 0;color:var(--text)"><span style="width:120px;font-family:'DM Mono',monospace;color:var(--text-dim)">${f4(b.lo)} – ${f4(b.hi)}</span><div class="p2bar" style="width:${b.n/top*100}%;min-width:2px"></div><span>${b.n}</span></div>`).join('')}</div><div class="dialog-actions"><button class="btn" id="p2x">Close</button><button class="btn primary" id="p2w">Write beside selection</button></div>`,460);
    $('#p2x').onclick=close;$('#p2w').onclick=()=>{close();run(()=>{const c0=s.c2+2;put(c0,s.r1,'Bin',{bold:true});put(c0+1,s.r1,'Count',{bold:true});bins.forEach((b,i)=>{put(c0,s.r1+1+i,`${f4(b.lo)} – ${f4(b.hi)}`);put(c0+1,s.r1+1+i,b.n)})},'Histogram written')}},'Show')}
function corrDlg(){const s=target();if(s.c2<=s.c1)return toast('Select two or more numeric columns');if(s.c2-s.c1>9)return toast('Select at most 10 columns');
  const D=data(),cols=[];for(let c=s.c1;c<=s.c2;c++){const h=V(D,c,s.r1),hdr=h!==''&&!isNum(h),a=[];for(let r=s.r1+(hdr?1:0);r<=s.r2;r++)a.push(V(D,c,r));cols.push({name:hdr?String(h):L(c),a})}
  const cell=(x,y)=>{const l=lin(x.a,y.a);if(!l)return'<td>—</td>';const r=l.r,col=r>=0?'74,222,128':'255,107,107';return`<td style="background:rgba(${col},${Math.abs(r)*.45});text-align:center">${r.toFixed(3)}</td>`};
  dlg(`<h3>Correlation matrix</h3><div style="overflow:auto"><table class="p2t"><tr><th></th>${cols.map(c=>`<th>${esc(c.name)}</th>`).join('')}</tr>${cols.map(x=>`<tr><th>${esc(x.name)}</th>${cols.map(y=>cell(x,y)).join('')}</tr>`).join('')}</table></div><div class="dialog-actions"><button class="btn primary" id="p2x">Close</button></div>`,520);$('#p2x').onclick=close}
function regress(){const s=target();if(s.c2<=s.c1)return toast('Select two columns: X first, then Y');
  const D=data(),X=s.c1,Y=s.c1+1,h=V(D,X,s.r1),hdr=h!==''&&!isNum(h),r0=s.r1+(hdr?1:0),xs=[],ys=[];
  for(let r=r0;r<=s.r2;r++){const a=V(D,X,r),b=V(D,Y,r);if(isNum(a)&&isNum(b)){xs.push(+a);ys.push(+b)}}
  const l=lin(xs,ys);if(!l)return toast('Need at least two numeric pairs that vary');
  dlg(`<h3>Linear regression</h3><div style="font-size:12px;color:var(--text-dim);margin-bottom:8px">X = column ${L(X)}, Y = column ${L(Y)} · ${l.n} points</div>${tableHtml(null,[['Slope',f4(l.slope)],['Intercept',f4(l.int)],['R²',f4(l.r*l.r)],['Correlation (r)',f4(l.r)],['Equation',`y = ${f4(l.slope)}x ${l.int<0?'−':'+'} ${Math.abs(+l.int.toFixed(4))}`]])}<label>Predict Y for X =</label><input id="rg1" type="number" step="any"><div id="rg2" style="margin-top:6px;color:var(--accent);font-weight:700"></div><div class="dialog-actions"><button class="btn" id="rg3">Close</button><button class="btn primary" id="rg4">Insert trend column</button></div>`,400);
  $('#rg1').oninput=e=>{const x=parseFloat(e.target.value);$('#rg2').textContent=isNaN(x)?'':'Y ≈ '+f4(l.int+l.slope*x)};$('#rg3').onclick=close;
  $('#rg4').onclick=()=>{close();const out=Math.max(s.c2,Y)+1;if(out>=COLS)return toast('No free column to the right');
    run(()=>{if(hdr)put(out,s.r1,'Trend',{bold:true});const a=r0+1,b=s.r2+1;for(let r=r0;r<=s.r2;r++)put(out,r,`=FORECAST(${L(X)}${r+1},$${L(Y)}$${a}:$${L(Y)}$${b},$${L(X)}$${a}:$${L(X)}$${b})`)},'Trend column added')}}
const OUT={'IQR method (1.5×)':a=>{const q1=pct(a,.25),q3=pct(a,.75),i=q3-q1;return v=>v<q1-1.5*i||v>q3+1.5*i},'Z-score > 3':a=>{const m=mean(a),sd=Math.sqrt(a.reduce((t,v)=>t+(v-m)**2,0)/(a.length-1))||1;return v=>Math.abs((v-m)/sd)>3},'Z-score > 2':a=>{const m=mean(a),sd=Math.sqrt(a.reduce((t,v)=>t+(v-m)**2,0)/(a.length-1))||1;return v=>Math.abs((v-m)/sd)>2}};
function outliers(name){const s=target(),a=numsIn(s);if(a.length<4)return toast('Need at least 4 numbers');const test=OUT[name](a),D=data();let n=0;
  run(()=>eachIn(s,(c,r)=>{const v=V(D,c,r);if(isNum(v)&&test(+v)){const x=cellO(c,r);x.f=x.f||{};x.f.bg='rgba(255,107,107,0.35)';n++}}),'');toast(n?`${n} outlier(s) highlighted`:'No outliers found')}
const COLK=['Running total','% of total','Cumulative %','Rank (high→low)','Rank (low→high)','Moving avg (3)','Min–max scaled','Z-score'];
function addCol(kind){const s=target(),X=s.c1,out=s.c2+1;if(out>=COLS)return toast('No free column to the right');
  const D=data(),h=V(D,X,s.r1),hdr=h!==''&&!isNum(h),r0=s.r1+(hdr?1:0);if(r0>s.r2)return toast('No data rows selected');
  const cl=L(X),a=r0+1,b=s.r2+1,Dr=`$${cl}$${a}:$${cl}$${b}`;
  run(()=>{if(hdr)put(out,s.r1,kind,{bold:true});
    for(let r=r0;r<=s.r2;r++){const i=r+1,x=`${cl}${i}`;let f='',fm;
      switch(kind){case'Running total':f=`=SUM($${cl}$${a}:${x})`;break;
        case'% of total':f=`=${x}/SUM(${Dr})`;fm={numFmt:'percent'};break;
        case'Cumulative %':f=`=SUM($${cl}$${a}:${x})/SUM(${Dr})`;fm={numFmt:'percent'};break;
        case'Rank (high→low)':f=`=RANK(${x},${Dr})`;break;
        case'Rank (low→high)':f=`=RANK(${x},${Dr},1)`;break;
        case'Moving avg (3)':if(r-r0>=2)f=`=AVERAGE(${cl}${i-2}:${x})`;break;
        case'Min–max scaled':f=`=(${x}-MIN(${Dr}))/(MAX(${Dr})-MIN(${Dr}))`;break;
        case'Z-score':f=`=(${x}-AVERAGE(${Dr}))/STDEV(${Dr})`;break}
      if(f)put(out,r,f,fm)}},kind+' column added in '+L(out))}
function goalSeek(){const a=API.getActive();
  form('Goal seek',[{k:'s',label:'Set cell (contains a formula)',def:id(a.c,a.r)},{k:'t',label:'To value',type:'number',def:'0'},{k:'c',label:'By changing cell (a constant)',def:''}],v=>{
    const S=parseRef(v.s),C=parseRef(v.c),T=parseFloat(v.t);if(!S||!C||isNaN(T))return toast('Enter valid cell references and a number');
    if(!isF(raw(S.c,S.r)))return toast('The "Set cell" must contain a formula');
    API.pushUndo();const x=cellO(C.c,C.r),orig=x.v;
    const f=n=>{x.v=String(n);const y=Number(V(data(),S.c,S.r));return y-T};
    let p=Number(orig)||1,q=p+Math.max(1,Math.abs(p)*.1),fp=f(p),fq=f(q),ok=false;
    for(let i=0;i<60;i++){if(!isFinite(fq)){break}if(Math.abs(fq)<1e-7){ok=true;break}if(fq===fp)break;const nx=q-fq*(q-p)/(fq-fp);p=q;fp=fq;q=nx;fq=f(q)}
    if(ok){x.v=String(+q.toFixed(8));API.refresh();toast(`Solution found: ${id(C.c,C.r)} = ${+q.toFixed(6)}`)}else{x.v=orig;API.refresh();toast('Goal seek could not find a solution')}},'Solve')}
function pivot(){const s=region(),D=data();if(s.c2<=s.c1||s.r2<=s.r1)return toast('Select a table with a header row first');
  const names=[];for(let c=s.c1;c<=s.c2;c++)names.push([String(c),L(c)+' — '+(String(V(D,c,s.r1))||'(blank)')]);
  form('Pivot summary → new sheet',[{k:'g',label:'Group rows by',opts:names},{k:'v',label:'Values',opts:names,def:names[names.length-1][0]},{k:'a',label:'Summarize with',opts:['Sum','Count','Average','Min','Max']}],v=>{
    const gc=+v.g,vc=+v.v,m=new Map(),all=[];
    for(let r=s.r1+1;r<=s.r2;r++){const k=String(V(D,gc,r));if(k==='')continue;const x=V(D,vc,r);if(v.a!=='Count'&&!isNum(x))continue;const n=v.a==='Count'?1:+x;if(!m.has(k))m.set(k,[]);m.get(k).push(n);all.push(n)}
    if(!m.size)return toast('No usable data found');
    const agg=a=>v.a==='Sum'||v.a==='Count'?a.reduce((x,y)=>x+y,0):v.a==='Average'?mean(a):v.a==='Min'?Math.min(...a):Math.max(...a);
    const keys=[...m.keys()].sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}));
    const body=keys.map(k=>[k,+agg(m.get(k)).toFixed(4)]);body.push(['Grand total',+agg(all).toFixed(4)]);
    toSheet('Pivot',[String(V(D,gc,s.r1)||'Group'),`${v.a} of ${V(D,vc,s.r1)||L(vc)}`],body);sh().colWidths[0]=160;sh().colWidths[1]=170;
    put(0,body.length,'Grand total',{bold:true});put(1,body.length,body[body.length-1][1],{bold:true});API.refresh()},'Create')}

/* ================= DATA TOOLS ================= */
function textToCols(){const s=target();form('Text to columns (splits the first selected column)',[{k:'d',label:'Delimiter',opts:['Comma','Semicolon','Space','Tab','Pipe |','Dash -','Custom']},{k:'x',label:'Custom delimiter (when "Custom")',def:''}],v=>{
  const d={'Comma':',','Semicolon':';','Space':' ','Tab':'\t','Pipe |':'|','Dash -':'-'}[v.d]||v.x;if(!d)return toast('Enter a delimiter');
  run(()=>{let mx=0;for(let r=s.r1;r<=s.r2;r++){const t=raw(s.c1,r);if(!t||isF(t))continue;const p=t.split(d).map(x=>x.trim());p.forEach((x,i)=>put(s.c1+i,r,x));mx=Math.max(mx,p.length)}toast(`Split into up to ${mx} columns`)})},'Split')}
function joinCols(){const s=target();if(s.c2-s.c1<1||s.c2+1>=COLS)return toast('Select two or more columns with a free column to their right');
  form('Join columns',[{k:'d',label:'Separator',def:' '}],v=>{const D=data();run(()=>{for(let r=s.r1;r<=s.r2;r++){const p=[];for(let c=s.c1;c<=s.c2;c++){const x=V(D,c,r);if(x!=='')p.push(x)}put(s.c2+1,r,p.join(v.d))}},`Joined into column ${L(s.c2+1)}`)},'Join')}
const TX={'Sentence case':v=>v.toLowerCase().replace(/(^\s*|[.!?]\s+)([a-z])/g,(m,a,b)=>a+b.toUpperCase()),'Swap case':v=>[...v].map(ch=>ch===ch.toUpperCase()?ch.toLowerCase():ch.toUpperCase()).join(''),'Reverse text':v=>[...v].reverse().join(''),'Remove line breaks':v=>v.replace(/[\r\n]+/g,' '),'Remove all spaces':v=>v.replace(/\s+/g,''),'Keep digits only':v=>v.replace(/\D+/g,''),'Keep letters only':v=>v.replace(/[^A-Za-z\s]+/g,''),'Extract email':v=>(v.match(/[\w.+-]+@[\w-]+(\.[\w-]+)+/)||[''])[0],'Extract first number':v=>(v.match(/-?\d+(\.\d+)?/)||[''])[0],'Remove first character':v=>v.slice(1),'Remove last character':v=>v.slice(0,-1),'Add prefix…':null,'Add suffix…':null};
function textXform(name){const apply=fn=>run(()=>{let n=0;eachIn(target(),(c,r)=>{const x=sh().data[id(c,r)];if(x&&typeof x.v==='string'&&x.v!==''&&!isF(x.v)){const nv=fn(x.v);if(nv!==x.v){x.v=nv;n++}}});toast(`${name}: ${n} cell(s) changed`)});
  if(name==='Add prefix…'||name==='Add suffix…'){const pre=name==='Add prefix…';form(name.replace('…',''),[{k:'t',label:'Text'}],v=>apply(s=>pre?v.t+s:s+v.t));return}apply(TX[name])}
const OP={'Add…':(x,k)=>x+k,'Subtract…':(x,k)=>x-k,'Multiply by…':(x,k)=>x*k,'Divide by…':(x,k)=>k?x/k:x,'Increase by %…':(x,k)=>x*(1+k/100),'Decrease by %…':(x,k)=>x*(1-k/100),'Round to decimals…':(x,k)=>+x.toFixed(Math.max(0,Math.min(10,k))),'Negate':x=>-x,'Absolute value':x=>Math.abs(x)};
function operate(name){const fn=OP[name],need=fn.length>1;
  const doit=k=>run(()=>{let n=0;eachIn(target(),(c,r)=>{const x=sh().data[id(c,r)];if(x&&typeof x.v==='string'&&!isF(x.v)&&isNum(x.v)){x.v=String(+(+fn(+x.v,k)).toFixed(10));n++}});toast(`${n} cell(s) updated`)});
  need?form(name.replace('…',''),[{k:'k',label:'Value',type:'number',def:'0'}],v=>doit(+v.k)):doit()}
function fillBlanks(){const s=target();run(()=>{let n=0;for(let c=s.c1;c<=s.c2;c++){let last=null;for(let r=s.r1;r<=s.r2;r++){const x=sh().data[id(c,r)];if(x&&x.v!==''&&x.v!==undefined)last=x;else if(last){put(c,r,last.v,JSON.parse(JSON.stringify(last.f||{})));n++}}}toast(`Filled ${n} blank cell(s)`)})}
function deleteBlankRows(){const s=target(),blank=new Set();
  for(let r=s.r1;r<=s.r2;r++){let e=true;for(let c=0;c<COLS;c++)if(raw(c,r)!==''){e=false;break}if(e)blank.add(r)}
  if(!blank.size)return toast('No completely blank rows in the selection');
  run(()=>{const nd={};Object.keys(sh().data).forEach(k=>{const p=parseKey(k);if(!p||blank.has(p.r))return;let shift=0;blank.forEach(b=>{if(b<p.r)shift++});nd[id(p.c,p.r-shift)]=sh().data[k]});sh().data=nd;toast(`Removed ${blank.size} blank row(s)`)})}
function toValues(){const s=target(),D=data();run(()=>{let n=0;eachIn(s,(c,r)=>{const x=sh().data[id(c,r)];if(x&&isF(x.v)){const v=V(D,c,r);if(!(typeof v==='string'&&v.startsWith('#'))){x.v=String(v);n++}}});toast(`${n} formula(s) converted to values`)})}
function reorder(name){const s=target();if(s.r1===s.r2)return toast('Select several rows');
  run(()=>{const rows=[];for(let r=s.r1;r<=s.r2;r++){const row=[];for(let c=s.c1;c<=s.c2;c++){const x=sh().data[id(c,r)];row.push(x?JSON.parse(JSON.stringify(x)):null)}rows.push(row)}
    if(name==='Shuffle rows'){for(let i=rows.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[rows[i],rows[j]]=[rows[j],rows[i]]}}else rows.reverse();
    rows.forEach((row,ri)=>row.forEach((x,ci)=>{const k=id(s.c1+ci,s.r1+ri);if(x)sh().data[k]=x;else delete sh().data[k]}))},name+' done')}
function seriesDlg(){form('Fill series',[{k:'t',label:'Type',opts:['Numbers','Dates (days)','Dates (months)']},{k:'s',label:'Start (number or yyyy-mm-dd)',def:'1'},{k:'p',label:'Step',type:'number',def:'1'},{k:'n',label:'Count',type:'number',def:'10'},{k:'d',label:'Direction',opts:['Down','Right']}],v=>{
  const s=sel(),n=Math.max(1,Math.min(1000,Math.round(+v.n)||1)),step=+v.p||1,num=v.t==='Numbers',base=num?Number(v.s):dt(v.s);
  if(num?isNaN(base):!base)return toast('Invalid start value');
  run(()=>{for(let i=0;i<n;i++){let val;if(num)val=base+step*i;else{const d=new Date(base);if(v.t==='Dates (days)')d.setUTCDate(d.getUTCDate()+step*i);else d.setUTCMonth(d.getUTCMonth()+step*i);val=iso(d)}v.d==='Down'?put(s.c1,s.r1+i,val):put(s.c1+i,s.r1,val)}},`Filled ${n} values`)},'Fill')}
function randDlg(){const NAMES=['Aarav','Diya','Rohan','Ananya','Kabir','Meera','Vihaan','Isha','Arjun','Priya','Aditya','Saanvi'];
  form('Random data (fills the selection)',[{k:'t',label:'Type',opts:['Integers','Decimals','Dates','Yes / No','First names']},{k:'a',label:'Min number / start date',def:'1'},{k:'b',label:'Max number / end date',def:'100'}],v=>{
    const s=target(),a=parseFloat(v.a),b=parseFloat(v.b),d1=dt(v.a)||new Date(),d2=dt(v.b)||new Date(+d1+365*864e5);
    run(()=>eachIn(s,(c,r)=>{let x;switch(v.t){case'Integers':x=Math.floor(Math.random()*((isNaN(b)?100:b)-(isNaN(a)?1:a)+1))+(isNaN(a)?1:a);break;case'Decimals':x=+(((isNaN(a)?0:a))+Math.random()*((isNaN(b)?1:b)-(isNaN(a)?0:a))).toFixed(2);break;case'Dates':x=iso(new Date(+d1+Math.random()*(d2-d1)));break;case'Yes / No':x=Math.random()<.5?'Yes':'No';break;default:x=NAMES[Math.floor(Math.random()*NAMES.length)]}put(c,r,x)}),'Random data filled')},'Fill')}
const dvClear=s=>{sh().dv=(sh().dv||[]).filter(R=>!(R.r1<=s.r2&&R.r2>=s.r1&&R.c1<=s.c2&&R.c2>=s.c1))};
function dvDlg(){const s=sel();form('Data validation',[{k:'t',label:'Allow',opts:['List','Whole number','Decimal','Text length']},{k:'l',label:'List items (comma-separated)',def:'Yes,No'},{k:'a',label:'Minimum (for numbers / length)',def:''},{k:'b',label:'Maximum (for numbers / length)',def:''}],v=>{
  const rule={r1:s.r1,c1:s.c1,r2:s.r2,c2:s.c2,type:{'List':'list','Whole number':'whole','Decimal':'decimal','Text length':'length'}[v.t],list:v.l.split(',').map(x=>x.trim()).filter(Boolean),min:v.a,max:v.b};
  if(rule.type==='list'&&!rule.list.length)return toast('Add at least one list item');
  dvClear(s);(sh().dv=sh().dv||[]).push(rule);API.refresh();toast('Validation applied — invalid entries get a red outline')},'Apply')}
function compareCols(){const s=sel();if(s.c2<=s.c1)return toast('Select two columns to compare');const D=data(),A=s.c1,B=s.c2,norm=v=>String(v).trim().toLowerCase(),setA=new Set(),setB=new Set();
  for(let r=s.r1;r<=s.r2;r++){const a=V(D,A,r),b=V(D,B,r);if(a!=='')setA.add(norm(a));if(b!=='')setB.add(norm(b))}let n=0;
  run(()=>{for(let r=s.r1;r<=s.r2;r++){[[A,setB],[B,setA]].forEach(([c,other])=>{const v=V(D,c,r);if(v!==''&&!other.has(norm(v))){const x=cellO(c,r);x.f=x.f||{};x.f.bg='rgba(255,180,84,0.35)';n++}})}},'');toast(n?`${n} unmatched cell(s) highlighted`:'Both columns contain the same values')}
const SP={'Blanks':(v,rw)=>rw==='','Formulas':(v,rw)=>isF(rw),'Errors':v=>typeof v==='string'&&/^#[A-Z/0!?]+!?$/.test(v),'Numbers':v=>isNum(v),'Text':v=>v!==''&&!isNum(v),'Clear highlights':null};
function special(name){const s=target(),D=data();
  if(name==='Clear highlights'){fmtSel(f=>{if(f.bg==='rgba(255,180,84,0.35)'||f.bg==='rgba(255,107,107,0.35)')f.bg='transparent'});return}
  let n=0;run(()=>eachIn(s,(c,r)=>{if(SP[name](V(D,c,r),raw(c,r))){const x=cellO(c,r);x.f=x.f||{};x.f.bg='rgba(255,180,84,0.35)';n++}}),'');toast(`${n} ${name.toLowerCase()} cell(s) highlighted`)}

/* ================= FINANCE ================= */
function emiSheet(v){const N=Math.round(v.N);if(!(N>0&&N<=600))return toast('Tenure must be 1–600 months');const e=F.EMI(v.P,v.R,N),r=v.R/1200;let bal=v.P;const body=[];
  for(let i=1;i<=N;i++){const it=bal*r,pr=e-it;bal=Math.max(0,bal-pr);body.push([i,+e.toFixed(2),+it.toFixed(2),+pr.toFixed(2),+bal.toFixed(2)])}
  toSheet('Loan EMI',['Month','EMI','Interest','Principal','Balance'],body,[1,2,3,4]);
  put(6,0,'Loan amount',{bold:true});put(7,0,v.P,{numFmt:'currency'});put(6,1,'Rate % p.a.',{bold:true});put(7,1,v.R);put(6,2,'Total interest',{bold:true});put(7,2,+(e*N-v.P).toFixed(2),{numFmt:'currency'});sh().colWidths[6]=120;sh().colWidths[7]=120;API.refresh()}
function growthSheet(name,years,fn){const body=[];for(let y=1;y<=years;y++){const[i,val]=fn(y);body.push([y,Math.round(i),Math.round(val),Math.round(val-i)])}toSheet(name,['Year','Invested','Value','Gain'],body,[1,2,3])}
const CALC={
 'Loan EMI':()=>calcDlg('Loan / EMI calculator',[['P','Loan amount (₹)',1000000],['R','Interest rate (% per year)',8.5],['N','Tenure (months)',120]],v=>{const e=F.EMI(v.P,v.R,Math.round(v.N));if(typeof e!=='number')return[['Result','Check the inputs']];const t=e*Math.round(v.N);return[['Monthly EMI',inr(e),1],['Total interest',inr(t-v.P)],['Total payment',inr(t)],['Interest share',((t-v.P)/t*100).toFixed(1)+'%']]},emiSheet),
 'SIP':()=>calcDlg('SIP calculator',[['M','Monthly investment (₹)',10000],['R','Expected return (% per year)',12],['Y','Years',10]],v=>{const fv=F.SIP(v.M,v.R,v.Y),inv=v.M*Math.round(v.Y*12);return[['Total invested',inr(inv)],['Estimated returns',inr(fv-inv)],['Maturity value',inr(fv),1]]},v=>growthSheet('SIP growth',Math.min(60,Math.round(v.Y)),y=>[v.M*12*y,F.SIP(v.M,v.R,y)])),
 'Lumpsum / FD':()=>calcDlg('Lumpsum / FD calculator',[['P','Amount (₹)',100000],['R','Interest rate (% per year)',7],['Y','Years',5],['N','Compounding per year',4]],v=>{const fv=v.P*Math.pow(1+v.R/100/v.N,v.N*v.Y);return[['Invested',inr(v.P)],['Interest earned',inr(fv-v.P)],['Maturity value',inr(fv),1]]},v=>growthSheet('FD growth',Math.min(60,Math.round(v.Y)),y=>[v.P,v.P*Math.pow(1+v.R/100/v.N,v.N*y)])),
 'CAGR':()=>calcDlg('CAGR calculator',[['B','Beginning value',100000],['E','Ending value',250000],['Y','Years',5]],v=>{const c=F.CAGR(v.B,v.E,v.Y);return[['CAGR',isFinite(c)?(c*100).toFixed(2)+'%':'—',1],['Absolute return',((v.E/v.B-1)*100).toFixed(1)+'%']]})};
function gstDlg(){const s=target();if(s.c2+2>=COLS)return toast('Need two free columns to the right of the selection');
  form('GST calculator',[{k:'r',label:'GST rate %',opts:['5','12','18','28'],def:'18'},{k:'m',label:'Amounts are…',opts:['Excluding GST (add GST)','Including GST (split out GST)']}],v=>{
    const D=data(),h=V(D,s.c1,s.r1),hdr=h!==''&&!isNum(h),r0=s.r1+(hdr?1:0),cl=L(s.c1),a=s.c2+1,b=s.c2+2,add=v.m.startsWith('Excl'),money={numFmt:'currency'};
    run(()=>{if(hdr){put(a,s.r1,add?'GST':'Base amount',{bold:true});put(b,s.r1,add?'Total':'GST',{bold:true})}
      for(let r=r0;r<=s.r2;r++){const x=`${cl}${r+1}`,A=`${L(a)}${r+1}`;if(add){put(a,r,`=${x}*${v.r}/100`,money);put(b,r,`=${x}+${A}`,money)}else{put(a,r,`=${x}/(1+${v.r}/100)`,money);put(b,r,`=${x}-${A}`,money)}}},`GST @ ${v.r}% applied`)},'Apply')}
const INRF={'₹ 12,34,567.00':{sym:'₹',d:2},'12,34,567':{sym:'',d:0},'₹ in Lakhs (L)':{sym:'₹',d:2,unit:'L'},'₹ in Crores (Cr)':{sym:'₹',d:2,unit:'Cr'},'Clear Indian format':null};
function inrFormat(k){const x=INRF[k];fmtSel(f=>{delete f.xf;f.numFmt='general';if(x)f.inr=Object.assign({},x);else delete f.inr})}
function amountWords(){const s=target();if(s.c2+1>=COLS)return toast('Need a free column to the right');run(()=>{for(let r=s.r1;r<=s.r2;r++)if(isNum(V(data(),s.c1,r)))put(s.c2+1,r,`=INRWORDS(${L(s.c1)}${r+1})`)},'Amount-in-words column added')}
const TPL={
 'GST Invoice':()=>{API.pushUndo();newSheet('Invoice');const H={bold:true,bg:'#2b3a6b',color:'#ffffff',align:'center'},w=(ref,v,f)=>{const p=parseRef(ref);put(p.c,p.r,v,f)};
  w('A1','TAX INVOICE',{bold:true,fontSize:20});w('A3','Invoice no:',{bold:true});w('B3','INV-001');w('A4','Date:',{bold:true});w('B4',iso(new Date()));w('A5','Bill to:',{bold:true});w('B5','Customer name');
  ['#','Item','Qty','Rate','GST %','Amount','GST','Total'].forEach((h,i)=>put(i,6,h,H));
  for(let i=0;i<6;i++){const r=7+i,n=r+1;put(0,r,i+1);if(!i){put(1,r,'Sample item');put(2,r,1);put(3,r,1000);put(4,r,18)}put(5,r,`=C${n}*D${n}`,{numFmt:'currency'});put(6,r,`=F${n}*E${n}/100`,{numFmt:'currency'});put(7,r,`=F${n}+G${n}`,{numFmt:'currency'})}
  w('E15','Total',{bold:true});w('F15','=SUM(F8:F13)',{bold:true,numFmt:'currency'});w('G15','=SUM(G8:G13)',{bold:true,numFmt:'currency'});w('H15','=SUM(H8:H13)',{bold:true,numFmt:'currency'});w('A17','In words:',{bold:true});w('B17','=INRWORDS(H15)');
  sh().colWidths[1]=180;API.refresh()},
 'Monthly Budget':()=>{API.pushUndo();newSheet('Budget');const H={bold:true,bg:'#2b3a6b',color:'#ffffff',align:'center'};['Category','Planned','Actual','Difference','% used'].forEach((h,i)=>put(i,0,h,H));
  ['Rent','Food','Transport','Utilities','Savings','Other'].forEach((c,i)=>{const r=i+1,n=r+1;put(0,r,c);put(1,r,0,{numFmt:'currency'});put(2,r,0,{numFmt:'currency'});put(3,r,`=B${n}-C${n}`,{numFmt:'currency'});put(4,r,`=IF(B${n}=0,0,C${n}/B${n})`,{numFmt:'percent'})});
  put(0,7,'Total',{bold:true});['B','C','D'].forEach(l=>put(colIdx(l),7,`=SUM(${l}2:${l}7)`,{bold:true,numFmt:'currency'}));put(4,7,'=IF(B8=0,0,C8/B8)',{bold:true,numFmt:'percent'});sh().colWidths[0]=140;API.refresh()},
 'Expense Tracker':()=>{API.pushUndo();newSheet('Expenses');const H={bold:true,bg:'#2b3a6b',color:'#ffffff',align:'center'};['Date','Category','Description','Amount','Payment mode'].forEach((h,i)=>{put(i,0,h,H)});
  put(6,0,'Total spent',{bold:true});put(7,0,'=SUM(D2:D200)',{bold:true,numFmt:'currency'});sh().colWidths[2]=200;
  sh().dv=[{r1:1,c1:1,r2:199,c2:1,type:'list',list:['Food','Travel','Bills','Shopping','Health','Other'],min:'',max:''},{r1:1,c1:4,r2:199,c2:4,type:'list',list:['Cash','UPI','Card','Bank transfer'],min:'',max:''},{r1:1,c1:3,r2:199,c2:3,type:'decimal',list:[],min:'0',max:''}];API.refresh()}};

/* ================= AUTOMATE ================= */
const SCK='prosheet_scripts';
const scripts=()=>{try{return JSON.parse(localStorage.getItem(SCK))||{}}catch(e){return{}}};
const SAMPLES={'Number the rows':"const n = range(selection).length;\nfor (let i = 0; i < n; i++) set('A' + (i + 1), i + 1);\nlog('Numbered', n, 'rows in column A');",'Column totals':"const rows = range(selection);\nconst totals = rows[0].map((_, c) => rows.reduce((s, r) => s + (Number(r[c]) || 0), 0));\nlog('Totals:', totals);","Trim + UPPERCASE first column":"const start = selection.split(':')[0];\nconst col = start.replace(/\\d/g, ''), row = +start.replace(/\\D/g, '');\nrange(selection).forEach((r, i) => set(col + (row + i), String(r[0]).trim().toUpperCase()));\nlog('Done');"};
const DEFAULT_SCRIPT="// Helpers: get('A1'), set('B1', 5), range('A1:C3'), setRange('E1', [[1,2],[3,4]]), log(), toast(), selection\nconst rows = range(selection);\nlog('Selected rows:', rows.length);\n";
function scriptDlg(preset){const saved=scripts();
  dlg(`<h3>Script editor</h3><div style="display:flex;gap:6px;margin-bottom:6px"><select id="sc1" style="flex:1"><option value="">— saved scripts —</option>${Object.keys(saved).map(k=>`<option>${esc(k)}</option>`).join('')}</select><button class="btn" id="sc2">Save</button><button class="btn" id="sc3">Delete</button></div><textarea id="sc4" rows="10" spellcheck="false" style="font-family:'DM Mono',monospace;font-size:12px;user-select:text;white-space:pre">${esc(preset||DEFAULT_SCRIPT)}</textarea><div id="sc5" style="font:11.5px 'DM Mono',monospace;color:var(--text-dim);margin-top:6px;max-height:100px;overflow:auto;white-space:pre-wrap;user-select:text"></div><div class="dialog-actions"><button class="btn" id="sc6">Close</button><button class="btn primary" id="sc7">▶ Run</button></div>`,600);
  $('#sc6').onclick=close;
  $('#sc1').onchange=e=>{if(e.target.value)$('#sc4').value=scripts()[e.target.value]||''};
  $('#sc2').onclick=()=>{const n=prompt('Script name?');if(!n)return;const s=scripts();s[n]=$('#sc4').value;try{localStorage.setItem(SCK,JSON.stringify(s))}catch(e){}toast('Script saved');scriptDlg($('#sc4').value)};
  $('#sc3').onclick=()=>{const n=$('#sc1').value;if(!n||!confirm('Delete script "'+n+'"?'))return;const s=scripts();delete s[n];localStorage.setItem(SCK,JSON.stringify(s));scriptDlg($('#sc4').value)};
  $('#sc7').onclick=async()=>{const out=$('#sc5');out.textContent='';let snap=null;const S=()=>snap||(snap=data());
    const log=(...a)=>{out.textContent+=a.map(x=>typeof x==='object'?JSON.stringify(x):String(x)).join(' ')+'\n'};
    const ref=r=>{const p=parseRef(r);if(!p)throw new Error('Bad cell reference: '+r);return p};
    const get=a=>{const p=ref(a);return V(S(),p.c,p.r)};
    const set=(a,v)=>{const p=ref(a);snap=null;put(p.c,p.r,v)};
    const range=a=>{const[x,y]=String(a).split(':'),p=ref(x),q=ref(y||x),o=[];for(let r=p.r;r<=q.r;r++){const row=[];for(let c=p.c;c<=q.c;c++)row.push(V(S(),c,r));o.push(row)}return o};
    const setRange=(a,m)=>{const p=ref(a);snap=null;m.forEach((row,i)=>row.forEach((v,j)=>put(p.c+j,p.r+i,v)))};
    const s=sel(),selStr=single(s)?id(s.c1,s.r1):`${id(s.c1,s.r1)}:${id(s.c2,s.r2)}`;
    API.pushUndo();
    try{const fn=new Function('get','set','range','setRange','log','selection','toast','return (async()=>{\n'+$('#sc4').value+'\n})()');await fn(get,set,range,setRange,log,selStr,toast);log('✔ Finished')}catch(e){log('✖ '+e.message)}
    API.refresh()}}
function matrix(){const s=region(),D=data(),m=[];for(let r=s.r1;r<=s.r2;r++){const row=[];for(let c=s.c1;c<=s.c2;c++)row.push(V(D,c,r));m.push(row)}return m}
const heads=m=>{const used=new Set();return m[0].map((h,i)=>{let k=String(h||'').trim().replace(/\s+/g,'_').replace(/[^\w]/g,'')||'col'+(i+1);while(used.has(k))k+='_';used.add(k);return k})};
const EXP={
 'Copy as Markdown':()=>{const m=matrix();if(m.length<2)return toast('Select a table with a header row');const e=v=>String(v).replace(/\|/g,'\\|');copyText(`| ${m[0].map(e).join(' | ')} |\n| ${m[0].map(()=>'---').join(' | ')} |\n`+m.slice(1).map(r=>`| ${r.map(e).join(' | ')} |`).join('\n'),'Markdown table copied')},
 'Copy as HTML table':()=>{const m=matrix();copyText('<table>\n'+m.map((r,i)=>`  <tr>${r.map(v=>`<${i?'td':'th'}>${esc(v)}</${i?'td':'th'}>`).join('')}</tr>`).join('\n')+'\n</table>','HTML table copied')},
 'Download JSON (records)':()=>{const m=matrix();if(m.length<2)return toast('Select a table with a header row');const h=heads(m);download(JSON.stringify(m.slice(1).map(r=>Object.fromEntries(h.map((k,i)=>[k,isNum(r[i])?+r[i]:r[i]]))),null,2),(API.getName()||'sheet')+'.json','application/json')},
 'Download SQL INSERTs':()=>{const m=matrix();if(m.length<2)return toast('Select a table with a header row');form('SQL export',[{k:'t',label:'Table name',def:'my_table'}],v=>{const h=heads(m),q=x=>isNum(x)?String(+x):"'"+String(x).replace(/'/g,"''")+"'";const t=v.t.replace(/\W/g,'_')||'my_table';
   const types=h.map((_,i)=>m.slice(1).every(r=>r[i]===''||isNum(r[i]))?'REAL':'TEXT');
   download(`CREATE TABLE ${t} (${h.map((k,i)=>`${k} ${types[i]}`).join(', ')});\n`+m.slice(1).map(r=>`INSERT INTO ${t} (${h.join(', ')}) VALUES (${r.map(q).join(', ')});`).join('\n'),t+'.sql','text/plain')},'Download')},
 'Download HTML page':()=>{const m=matrix();download(`<!DOCTYPE html><meta charset="utf-8"><title>${esc(API.getName())}</title><style>body{font:14px system-ui;margin:24px}table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:6px 10px}th{background:#2b3a6b;color:#fff}</style><table>${m.map((r,i)=>`<tr>${r.map(v=>`<${i?'td':'th'}>${esc(v)}</${i?'td':'th'}>`).join('')}</tr>`).join('')}</table>`,(API.getName()||'sheet')+'.html','text/html')}};
function nameMgr(){const n=WB().namedRanges||{},k=Object.keys(n);
  dlg(`<h3>Name manager</h3>${k.length?`<div class="xfl">${k.map(x=>`<div class="p2li" data-k="${esc(x)}"><span><b>${esc(x)}</b> <small style="opacity:.7">${esc(n[x].sheet)}!${id(n[x].c1,n[x].r1)}:${id(n[x].c2,n[x].r2)}</small></span><button class="btn" data-d="${esc(x)}" style="color:var(--danger)">✕</button></div>`).join('')}</div>`:'<div style="color:var(--text-dim);font-size:12.5px">No named ranges. Use Data → Name this range…</div>'}<div class="dialog-actions"><button class="btn" id="p2x">Close</button></div>`,440);
  $('#p2x').onclick=close;
  $$('#dialogBox [data-d]').forEach(b=>b.onclick=e=>{e.stopPropagation();delete WB().namedRanges[b.dataset.d];API.refresh();nameMgr()});
  $$('#dialogBox .p2li').forEach(d=>d.onclick=()=>{const r=n[d.dataset.k];close();if(!WB().sheets[r.sheet])return toast('Sheet no longer exists');WB().current=r.sheet;API.refresh();const b=$('#cellref');b.value=`${id(r.c1,r.r1)}:${id(r.c2,r.r2)}`;b.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'}))})}

/* ================= build the ribbon ================= */
grp('Page Layout','Themes',[mkb('🌓','Dark / Light',toggleTheme),mks('🎨','Accent',Object.keys(ACC),setAccent)]);
grp('Page Layout','Page Setup',[mks('⟳','Orientation',['Portrait','Landscape'],v=>setPage('o',v)),mks('▭','Margins',Object.keys(MARG),v=>setPage('m',v)),mks('📄','Paper size',['A4','A3','A5','Letter','Legal'],v=>setPage('p',v)),mkb('▣','Print area',()=>{sh().printArea=sel();API.refresh();toast('Print area set to selection')}),mkb('✖','Clear area',()=>{sh().printArea=null;API.refresh();toast('Print area cleared')})]);
grp('Page Layout','Scale & Print',[mks('📐','Print scale',['100%','90%','80%','70%','60%','50%'],v=>setPage('s',parseInt(v,10))),mkb('🖨','Print',()=>window.print())]);
grp('Page Layout','Zoom',[mkb('＋','Zoom in',()=>setZoom(zoom+.1)),mkb('－','Zoom out',()=>setZoom(zoom-.1)),mkb('◎','100%',()=>setZoom(1))]);
grp('Page Layout','Sheet',[mks('🏷','Tab color',Object.keys(TABC),k=>{sh().tabColor=TABC[k];API.refresh()}),mkb('◀','Move left',()=>moveSheet(-1)),mkb('▶','Move right',()=>moveSheet(1)),mkb('🙈','Hide sheet',hideSheet),mkb('👁','Unhide',unhideSheet)]);

grp('Review','Proofing',[mkb('ABC','Spell check',()=>{spell=!spell;$('#cellEditor').spellcheck=spell;toast('Spell check '+(spell?'on (while editing a cell)':'off'))}),mkb('#','Workbook stats',statsDlg)]);
grp('Review','Notes',[mkb('🗒','Notes list',notesList),mkb('🗑','Delete notes',()=>{if(!Object.keys(sh().notes||{}).length)return toast('No notes on this sheet');if(confirm('Delete all notes on this sheet?')){sh().notes={};API.refresh()}})]);
grp('Review','Protect',[mkb('🔒','Protect sheet',protectDlg),mkb('🔓','Unprotect',unprotect),mkb('🔐','Lock cells',()=>lockCells(false)),mkb('✎','Unlock cells',()=>lockCells(true))]);
grp('Review','Auditing',[mkb('ƒ','Show formulas',()=>{showF=!showF;API.refresh();toast('Show formulas '+(showF?'on':'off'))}),mkb('⚠','Error check',errDlg)]);
grp('Review','Versions',[mkb('💾','Save version',()=>form('Save version',[{k:'n',label:'Version name',def:'Version '+(vers().length+1)}],v=>{if(saveVersion(v.n))toast('Version saved')},'Save')),mkb('🕘','History',histDlg)]);

grp('Analyze','Statistics',[mkb('Σ','Describe',describeDlg),mkb('▥','Histogram',histogram),mkb('⊞','Correlation',corrDlg),mkb('📈','Regression',regress)]);
grp('Analyze','Data quality',[mks('⚑','Highlight outliers',Object.keys(OUT),outliers)]);
grp('Analyze','Calculated column',[mks('＋','Add column',COLK,addCol)]);
grp('Analyze','What-if & Pivot',[mkb('🎯','Goal seek',goalSeek),mkb('⊟','Pivot summary',pivot)]);

grp('Data Tools','Text',[mkb('⫼','Text to columns',textToCols),mkb('⧉','Join columns',joinCols),mks('Aa','Transform',Object.keys(TX),textXform)]);
grp('Data Tools','Cells',[mkb('⬇','Fill blanks',fillBlanks),mkb('␡','Delete blank rows',deleteBlankRows),mkb('=→1','Formulas → values',toValues),mks('±','Operate on numbers',Object.keys(OP),operate),mks('🔀','Reorder rows',['Shuffle rows','Reverse rows'],reorder)]);
grp('Data Tools','Generate',[mkb('1,2,3','Fill series',seriesDlg),mkb('🎲','Random data',randDlg)]);
grp('Data Tools','Validation',[mkb('✔','Data validation',dvDlg),mkb('✖','Clear validation',()=>{dvClear(sel());API.refresh();toast('Validation cleared')})]);
grp('Data Tools','Compare',[mkb('⇄','Compare columns',compareCols),mks('🔦','Highlight cells',Object.keys(SP),special)]);

grp('Finance','Calculators',[mks('🧮','Calculators',Object.keys(CALC),k=>CALC[k]()),mkb('%','GST',gstDlg)]);
grp('Finance','Indian ₹',[mks('₹','Lakh / Crore format',Object.keys(INRF),inrFormat),mkb('Aa','Amount in words',amountWords)]);
grp('Finance','Templates',[mks('🗂','Templates',Object.keys(TPL),k=>TPL[k]())]);

grp('Automate','Script',[mkb('⌘','Script editor',()=>scriptDlg()),mks('📜','Samples',Object.keys(SAMPLES),k=>scriptDlg(SAMPLES[k]))]);
grp('Automate','Export',Object.keys(EXP).map(k=>mkb('⇪',k.replace('Download ','').replace('Copy as ','→ '),EXP[k])));
grp('Automate','Names',[mkb('🏷','Name manager',nameMgr)]);

decTabs();
toast('Add-on 2 loaded — new tabs: Page Layout, Review, Analyze, Data Tools, Finance, Automate');
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
