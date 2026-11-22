/* Pro Sheet — Excel look + charts, sparklines, pivot tables, data validation,
   text-to-columns, relative AutoFill, real freeze panes, named-range jump.
   Load AFTER addon.js:  <script src="excel-ui.js"></script>
   Needs one tiny patch in prosheet.html: add  FN: FN,  inside window.ProSheetAPI = { ... }  */
(function(){
"use strict";
try{ if(localStorage.getItem('prosheet_theme')===null) localStorage.setItem('prosheet_theme','light'); }catch(e){}

function boot(){
const API=window.ProSheetAPI;
if(!API||!API.getWorkbook||!document.getElementById('chartBtn')) return setTimeout(boot,200);
const $=s=>document.querySelector(s), L=API.colLetter, COLS=26;
const WB=()=>API.getWorkbook(), sh=()=>WB().sheets[WB().current];
const id=(c,r)=>L(c)+(r+1);
const raw=(c,r)=>{const x=sh().data[id(c,r)];return x&&x.v!==undefined?String(x.v):""};
const cellO=(c,r)=>{const d=sh().data;return d[id(c,r)]||(d[id(c,r)]={v:"",f:{}})};
const sel=()=>{const s=API.getSelection();return{r1:Math.min(s.r1,s.r2),r2:Math.max(s.r1,s.r2),c1:Math.min(s.c1,s.c2),c2:Math.max(s.c1,s.c2)}};
const act=a=>{const e=document.querySelector(`.dropdown-item[data-act="${a}"]`);if(e)e.click()};
const addon=t=>{const e=[...document.querySelectorAll('#addonMenu .ai')].find(x=>x.textContent.includes(t));if(e)e.click()};
const num=v=>{const n=parseFloat(v);return isNaN(n)?NaN:n};
const esc=s=>String(s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const dlg=h=>{$('#dialogBox').innerHTML=h;$('#dialogOverlay').classList.add('show')};
const close=()=>$('#dialogOverlay').classList.remove('show');
const colN=s=>[...s].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)-1;

/* ================= STYLE ================= */
const css=document.createElement('style');
css.textContent=`
html.ps-light{--bg:#fff;--surface:#fff;--surface-2:#f3f3f3;--surface-3:#e1e1e1;--border:#d4d4d4;--border-soft:#e1e1e1;--text:#222;--text-dim:#444;--text-faint:#777;--header-bg:#f3f3f3;--grid-line:#dadada;--accent:#217346;--accent-dim:#b7dcc6;--accent-soft:rgba(33,115,70,.13)}
html.ps-light .colhead.sel,html.ps-light .rowhead.sel{background:#d2d2d2;color:#217346;font-weight:700}
html.ps-light .colhead.sel{border-bottom:2px solid #217346}
html.ps-light .rowhead.sel{border-right:2px solid #217346}
.topbar{background:#217346!important;border:0!important;padding:5px 12px!important;color:#fff}
.topbar .logo,.topbar #filename,.topbar .tbtn{color:#fff}
.topbar .logo .mark{background:#fff;color:#217346}
.topbar .menubar{display:none}
.topbar .filename-wrap{flex:1;align-items:center}
.topbar #filename{text-align:center;font-weight:600}
.topbar #filename:hover,.topbar #filename:focus{background:rgba(255,255,255,.18);border-color:transparent}
.topbar .save-status{color:rgba(255,255,255,.8)!important}
.topbar .btn{background:rgba(255,255,255,.16);border-color:transparent;color:#fff}
.topbar .btn:hover{background:rgba(255,255,255,.28)}
.topbar .btn.primary{background:#fff;color:#217346}
.topbar .account-chip{background:rgba(255,255,255,.16);border-color:transparent;color:#fff}
.xl-tabs{display:flex;background:var(--header-bg);flex-shrink:0;padding:0 6px}
.xt{padding:7px 16px;font-size:13px;color:var(--text);cursor:pointer;border-bottom:3px solid transparent}
.xt:hover{background:var(--surface-3)}
.xt.on{color:#217346;font-weight:700;border-bottom-color:#217346}
.xt:first-child{background:#217346;color:#fff;border-bottom-color:#217346}
.xl-rib{flex-shrink:0;background:var(--header-bg);border-bottom:1px solid var(--border);border-top:1px solid var(--border)}
.xp{display:none;align-items:stretch;padding:4px 8px;min-height:86px;overflow-x:auto}
.xp.on{display:flex}
.xp .toolbar{background:transparent;border:0;flex-wrap:wrap;align-content:flex-start;padding:6px 4px;gap:4px}
.xg{display:flex;flex-direction:column;padding:0 8px;border-right:1px solid var(--border);justify-content:space-between;flex-shrink:0}
.xr{display:flex;gap:2px;padding-top:4px;align-items:flex-start;flex:1}
.xl{font-size:11px;color:var(--text-faint);text-align:center;padding:3px 0 1px}
.xb{display:flex;flex-direction:column;align-items:center;gap:2px;padding:4px 7px;min-width:54px;border-radius:3px;cursor:pointer;font-size:11.5px;color:var(--text);user-select:none;text-align:center}
.xb:hover{background:var(--surface-3)}
.xb i{font-style:normal;font-size:22px;line-height:26px}
.xb select{font-size:11.5px;background:var(--surface);color:var(--text);border:1px solid var(--border);padding:2px;max-width:96px}
.formulabar{background:var(--surface)!important}
#cellref{background:var(--surface);color:var(--text)}
.xfb{width:30px;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#777;border-right:1px solid var(--border);font-size:13px}
.xfb:hover{background:var(--surface-3)}
.xfb:nth-of-type(2){color:#217346}
td.cell{font-family:Calibri,Carlito,'DM Sans',sans-serif!important;font-size:14px}
td.cell.active-cell{outline:2px solid #217346}
.fill-handle{background:#217346!important}
.sheettabs{background:var(--header-bg);padding:0 8px;border-top:1px solid var(--border)}
.tab{border-radius:0;padding:6px 16px}
.tab.active{background:var(--surface);color:#217346;font-weight:700;border-bottom:3px solid #217346;border-color:var(--border) var(--border) #217346}
.tab-add{border:1px solid #888;border-radius:50%;margin:0 8px}
.xstat{display:flex;align-items:center;gap:14px;padding:2px 12px;background:var(--header-bg);border-top:1px solid var(--border);font-size:11.5px;color:var(--text-dim);flex-shrink:0}
.xstat .status-right{margin-left:auto}
.xstat input[type=range]{width:110px;accent-color:#217346}
.xchart{position:absolute;z-index:40;background:var(--surface);border:1px solid var(--border);box-shadow:0 4px 14px rgba(0,0,0,.18);resize:both;overflow:hidden;min-width:240px;min-height:160px;display:flex;flex-direction:column}
.xch{display:flex;align-items:center;gap:8px;padding:3px 8px;background:var(--header-bg);cursor:move;font-size:12px;color:var(--text);flex-shrink:0}
.xch b{flex:1;font-weight:600}
.xch select{font-size:11px}
.xch .x{cursor:pointer;color:var(--text-faint)}
.xch .x:hover{color:#c0392b}
.xcb{flex:1;min-height:0;overflow:hidden}
.xcb text{fill:var(--text-dim)}
.xcb .gl{stroke:var(--border)}
.xcb .ax{stroke:var(--text-faint)}
#xdv{position:absolute;z-index:40;width:22px;appearance:none;-webkit-appearance:none;border:1px solid #999;color:transparent;cursor:pointer;background:#f3f3f3 url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='10' height='6'><path d='M0 0h10L5 6z' fill='%23555'/></svg>") no-repeat center}
#xdv option{color:#000}`;
document.head.appendChild(css);

/* ================= RIBBON ================= */
const tabRow=document.createElement('div');tabRow.className='xl-tabs';
const rib=document.createElement('div');rib.className='xl-rib';
$('.topbar').after(tabRow);tabRow.after(rib);
const panels={};
function btn(icon,label,fn){const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${icon}</i><span>${label}</span>`;b.onclick=e=>{e.stopPropagation();fn()};return b}
function fnSel(label,names){const b=document.createElement('div');b.className='xb';
  b.innerHTML=`<i>ƒx</i><select><option value="">${label}</option>${names.map(n=>`<option>${n}</option>`).join('')}</select>`;
  b.querySelector('select').onchange=e=>{if(e.target.value)insertFn(e.target.value);e.target.value=''};return b}
function group(label,items){const g=document.createElement('div');g.className='xg';const r=document.createElement('div');r.className='xr';items.forEach(i=>r.appendChild(i));g.appendChild(r);const l=document.createElement('div');l.className='xl';l.textContent=label;g.appendChild(l);return g}
function panel(name,groups){const p=document.createElement('div');p.className='xp';groups.forEach(g=>p.appendChild(g));rib.appendChild(p);panels[name]=p;const t=document.createElement('div');t.className='xt';t.textContent=name;t.onclick=()=>show(name);tabRow.appendChild(t)}
function show(n){Object.entries(panels).forEach(([k,p])=>p.classList.toggle('on',k===n));[...tabRow.children].forEach(t=>t.classList.toggle('on',t.textContent===n))}

const CT=[['bar','📊','Bar'],['line','📈','Line'],['pie','🥧','Pie'],['doughnut','🍩','Doughnut'],['scatter','⁙','Scatter'],['area','◭','Area'],['histogram','▥','Histogram'],['waterfall','⇵','Waterfall'],['radar','🕸','Radar'],['combo','📉','Combo']];
const A=(i,l,a)=>btn(i,l,()=>act(a));

panel('File',[
 group('Workbook',[A('📄','New sheet','new'),A('➕','Add sheets','addMultipleSheets'),A('💾','Save','save'),A('🖨','Print','print')]),
 group('Import',[A('📥','CSV','importCsv'),A('📗','Excel','importXlsx')]),
 group('Export',[A('📤','CSV','exportCsv'),A('📗','Excel .xlsx','exportXlsx'),A('{ }','JSON','exportJson')])]);
panel('Home',[]);panels.Home.appendChild($('#toolbar'));
panel('Insert',[
 group('Tables',[btn('🧮','PivotTable',pivotDialog)]),
 group('Charts',CT.map(t=>btn(t[1],t[2],()=>addChart(t[0])))),
 group('Sparklines',[btn('〰','Line',()=>spark_('line')),btn('▮','Column',()=>spark_('column')),btn('±','Win/Loss',()=>spark_('winloss'))]),
 group('Cells',[A('⬆','Row above','rowAbove'),A('⬇','Row below','rowBelow'),A('⬅','Col left','colLeft'),A('➡','Col right','colRight')])]);
panel('Formulas',[
 group('Function Library',[btn('Σ','AutoSum',()=>$('#sumBtn').click()),
  fnSel('Math & Stats',['SUM','AVERAGE','COUNT','COUNTA','MIN','MAX','MEDIAN','ROUND','ABS','STDEV']),
  fnSel('Logical',['IF','IFS','AND','OR','NOT','IFERROR']),
  fnSel('Lookup',['VLOOKUP','HLOOKUP','XLOOKUP','INDEX','MATCH']),
  fnSel('Conditional',['COUNTIF','SUMIF','AVERAGEIF']),
  fnSel('Text',['CONCAT','TEXTJOIN','LEFT','RIGHT','MID','LEN','UPPER','LOWER','PROPER','TRIM']),
  fnSel('Date',['TODAY','NOW','DATE','YEAR','MONTH','DAY','DATEDIF','EDATE','EOMONTH','NETWORKDAYS']),
  fnSel('Sparkline',['SPARKLINE'])]),
 group('Defined Names',[A('🏷','Define name','namedRange')])]);
panel('Data',[
 group('Sort',[A('↓A','Sort A→Z','sortAsc'),A('↓Z','Sort Z→A','sortDesc')]),
 group('Data Tools',[btn('🧹','Remove duplicates',()=>addon('Remove duplicate')),btn('⫼','Text to columns',textToCols),btn('☑','Data validation',dvDialog),btn('🧮','PivotTable',pivotDialog)]),
 group('Analyze',[A('🎨','Conditional formatting','condFormat'),btn('🔴','Highlight dupes',()=>addon('Highlight dup')),A('🔍','Find & replace','find')])]);
panel('View',[
 group('Freeze Panes',[A('❄','Freeze row 1','freezeRow'),A('❄','Freeze col A','freezeCol'),A('👁','Unhide all','unhideAll')]),
 group('Theme',[btn('🌓','Light / Dark',()=>addon('theme'))])]);
show('Home');

/* formula bar ✕ ✓, Ready bar, zoom, tab order */
const fx=$('.fx-icon');
[['✕','Escape'],['✓','Enter']].forEach(([t,k])=>{const b=document.createElement('div');b.className='xfb';b.textContent=t;b.onmousedown=e=>e.preventDefault();b.onclick=()=>$('#formulaInput').dispatchEvent(new KeyboardEvent('keydown',{key:k}));fx.before(b)});
$('#tabsList').after($('#addSheetBtn'));
const sb=document.createElement('div');sb.className='xstat';
sb.innerHTML='<span>Ready</span>';sb.appendChild($('.status-right'));
const zm=document.createElement('span');zm.innerHTML='− <input type="range" min="50" max="200" value="100" id="xzoom"> + <b id="xzv">100%</b>';
sb.appendChild(zm);$('#sheettabs').after(sb);
$('#xzoom').oninput=e=>{$('#grid').style.zoom=e.target.value/100;$('#xzv').textContent=e.target.value+'%'};
$('#chartBtn').onclick=()=>addChart('bar');
document.addEventListener('click',e=>{if(e.target.closest('[data-act="chart"]')){e.stopImmediatePropagation();addChart('bar')}},true);

/* name box: typing a defined name jumps to its range */
$('#cellref').addEventListener('keydown',e=>{if(e.key!=='Enter')return;const nr=(WB().namedRanges||{});const k=Object.keys(nr).find(n=>n.toLowerCase()===e.target.value.trim().toLowerCase());
  if(k){const r=nr[k];e.target.value=id(r.c1,r.r1)+(r.r1===r.r2&&r.c1===r.c2?'':':'+id(r.c2,r.r2))}},true);

/* ================= FORMULA HELPERS ================= */
function insertFn(n){const a=API.getActive(),s=sel(),multi=s.r1!==s.r2||s.c1!==s.c2;
  API.pushUndo();cellO(a.c,a.r).v=`=${n}(${multi?id(s.c1,s.r1)+':'+id(s.c2,s.r2):''})`;API.refresh();$('#formulaInput').focus()}
function shiftF(f,dr,dc){return f.replace(/(?<![A-Za-z0-9_])(\$?)([A-Z]{1,2})(\$?)(\d+)(?![\d(A-Za-z_])/g,(m,a,col,b,row)=>{let ci=colN(col),ri=+row;if(!a)ci+=dc;if(!b)ri+=dr;return ci<0||ri<1?'#REF!':a+L(ci)+b+ri})}
function cp(src,dr,dc){const o=JSON.parse(JSON.stringify(src));if(typeof o.v==='string'&&o.v[0]==='=')o.v=shiftF(o.v,dr,dc);return o}
function fillBy(dir){const s=sel();API.pushUndo();
  if(dir==='d'){if(s.r1===s.r2)return;for(let c=s.c1;c<=s.c2;c++){const src=sh().data[id(c,s.r1)];if(src)for(let r=s.r1+1;r<=s.r2;r++)sh().data[id(c,r)]=cp(src,r-s.r1,0)}}
  else{if(s.c1===s.c2)return;for(let r=s.r1;r<=s.r2;r++){const src=sh().data[id(s.c1,r)];if(src)for(let c=s.c1+1;c<=s.c2;c++)sh().data[id(c,r)]=cp(src,0,c-s.c1)}}
  API.refresh();API.toast(dir==='d'?'Filled down':'Filled right')}
document.addEventListener('keydown',e=>{const a=document.activeElement;if(a&&(a.tagName==='INPUT'||a.tagName==='SELECT'||a.isContentEditable))return;
  if((e.ctrlKey||e.metaKey)&&['d','r'].includes(e.key.toLowerCase())){e.preventDefault();e.stopImmediatePropagation();fillBy(e.key.toLowerCase())}},true);
/* drag-handle AutoFill: shift relative references in copied formulas */
let fillSrc=null;
$('#grid').addEventListener('mousedown',e=>{if(e.target.closest('.fill-handle'))fillSrc={...API.getActive()}},true);
document.addEventListener('mouseup',()=>{if(!fillSrc)return;const f=fillSrc;fillSrc=null;
  setTimeout(()=>{const src=sh().data[id(f.c,f.r)];if(!src||typeof src.v!=='string'||src.v[0]!=='=')return;const s=sel();
    for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++)if(r!==f.r||c!==f.c)cellO(c,r).v=shiftF(src.v,r-f.r,c-f.c);API.refresh()},0)});

/* ================= SPARKLINES ================= */
if(API.FN)API.FN.SPARKLINE=(d,t)=>'SPARK|'+(t||'line')+'|'+[d].flat(Infinity).map(x=>isNaN(+x)?0:+x).join(',');
else console.warn('Pro Sheet: add "FN: FN," to window.ProSheetAPI to enable SPARKLINE()');
function sparkSvg(ty,v){const n=v.length,mn=Math.min(...v),mx=Math.max(...v),R=(mx-mn)||1,H=20,Y=x=>H-2-(x-mn)/R*(H-4),bw=100/n;
  const open='<svg viewBox="0 0 100 20" preserveAspectRatio="none" style="width:100%;height:18px;display:block">';
  if(ty==='line')return open+`<polyline fill="none" stroke="#4472c4" stroke-width="1.6" vector-effect="non-scaling-stroke" points="${v.map((x,i)=>(n>1?i/(n-1)*98+1:50)+','+Y(x)).join(' ')}"/></svg>`;
  if(ty==='winloss')return open+v.map((x,i)=>x?`<rect x="${i*bw+bw*.1}" y="${x>0?2:10}" width="${bw*.8}" height="8" fill="${x>0?'#70ad47':'#ed7d31'}"/>`:'').join('')+'</svg>';
  return open+v.map((x,i)=>`<rect x="${i*bw+bw*.1}" y="${Y(x)}" width="${bw*.8}" height="${H-1-Y(x)}" fill="#4472c4"/>`).join('')+'</svg>'}
function spark_(ty){const s=sel();API.pushUndo();
  if(s.c2>s.c1)for(let r=s.r1;r<=s.r2;r++)cellO(s.c2+1,r).v=`=SPARKLINE(${id(s.c1,r)}:${id(s.c2,r)},"${ty}")`;
  else cellO(s.c1+1,s.r1).v=`=SPARKLINE(${id(s.c1,s.r1)}:${id(s.c1,s.r2)},"${ty}")`;
  API.refresh();API.toast('Sparkline inserted')}

/* ================= CHARTS ================= */
const PAL=['#4472c4','#ed7d31','#a5a5a5','#ffc000','#5b9bd5','#70ad47','#264478','#9e480e'];
function nice(mx,mn){const sp=(mx-mn)||1,s0=sp/5,mag=Math.pow(10,Math.floor(Math.log10(s0))),m=s0/mag,st=(m<1.5?1:m<3?2:m<7?5:10)*mag,t=[];
  for(let v=Math.floor(mn/st)*st;v<=Math.ceil(mx/st)*st+st/2;v+=st)t.push(+v.toFixed(10));return t}
function regionSel(){const s=sel();if(s.r1!==s.r2||s.c1!==s.c2)return s;let{r1,r2,c1,c2}=s;
  const rowHas=(r,a,b)=>{for(let c=Math.max(a,0);c<=Math.min(b,COLS-1);c++)if(raw(c,r)!=="")return 1};
  const colHas=(c,a,b)=>{for(let r=Math.max(a,0);r<=b;r++)if(raw(c,r)!=="")return 1};
  for(let ch=1;ch;){ch=0;
    if(r1>0&&rowHas(r1-1,c1-1,c2+1)){r1--;ch=1}
    if(r2<sh().rows-1&&rowHas(r2+1,c1-1,c2+1)){r2++;ch=1}
    if(c1>0&&colHas(c1-1,r1-1,r2+1)){c1--;ch=1}
    if(c2<COLS-1&&colHas(c2+1,r1-1,r2+1)){c2++;ch=1}}
  return{r1,r2,c1,c2}}
function getData(c,snap){const g=(r,k)=>(snap[r]||[])[k],multi=c.c2>c.c1,vc=multi?c.c1+1:c.c1;let hdr=false;
  for(let k=vc;k<=c.c2;k++)if(typeof g(c.r1,k)==='string'&&g(c.r1,k)!=="")hdr=true;
  const rs=c.r1+(hdr?1:0),cats=[],series=[];
  for(let r=rs;r<=c.r2;r++)cats.push(multi?String(g(r,c.c1)??''):String(r-rs+1));
  for(let k=vc;k<=c.c2;k++){const v=[];for(let r=rs;r<=c.r2;r++){const n=Number(g(r,k));v.push(g(r,k)===""||isNaN(n)?0:n)}series.push({name:hdr?String(g(c.r1,k)):'Series '+(k-vc+1),v})}
  return{cats,series}}
function draw(type,D,W,H){
  const S=D.series,pie=type==='pie'||type==='doughnut',showLeg=pie||(S.length>1&&type!=='histogram'&&type!=='waterfall');
  let cats=D.cats,o='';
  const T=(x,y,s,a)=>`<text x="${x}" y="${y}" text-anchor="${a||'middle'}" font-size="10">${esc(s)}</text>`;
  const lg=nm=>{let x=10,r='';nm.forEach((s,i)=>{r+=`<rect x="${x}" y="${H-13}" width="9" height="9" fill="${PAL[i%8]}"/>`+T(x+13,H-5,s,'start');x+=String(s).length*6+28});return r};
  if(pie){const v=S[0].v.map(x=>Math.max(0,x)),tot=v.reduce((a,b)=>a+b,0)||1,cx=W/2,cy=(H-20)/2,R=Math.max(10,Math.min(W,H-26)/2-8),r0=type==='doughnut'?R*.55:0;let a=-Math.PI/2;
    v.forEach((x,i)=>{const sw=Math.min(x/tot*2*Math.PI,6.2831),b=a+sw,big=sw>Math.PI?1:0,P=(g,q)=>`${cx+q*Math.cos(g)},${cy+q*Math.sin(g)}`;
      if(x>0)o+=`<path d="M${P(a,R)} A${R},${R} 0 ${big} 1 ${P(b,R)} ${r0?`L${P(b,r0)} A${r0},${r0} 0 ${big} 0 ${P(a,r0)}`:`L${cx},${cy}`} Z" fill="${PAL[i%8]}" stroke="#fff"/>`;
      if(x/tot>.05){const m=(a+b)/2,q=r0?(R+r0)/2:R*.65;o+=`<text x="${cx+q*Math.cos(m)}" y="${cy+q*Math.sin(m)+3}" text-anchor="middle" font-size="10" style="fill:#fff;font-weight:700">${Math.round(x/tot*100)}%</text>`}a=b});
    return o+lg(cats)}
  const n0=cats.length;
  if(type==='radar'){const n=n0,cx=W/2,cy=(H-20)/2+4,R=Math.max(10,Math.min(W,H-40)/2-24),mx=Math.max(1,...S.flatMap(s=>s.v)),P=(i,k)=>{const a=-Math.PI/2+i/n*2*Math.PI;return[cx+R*k*Math.cos(a),cy+R*k*Math.sin(a)]};
    for(let k=1;k<=5;k++)o+=`<polygon class="gl" points="${cats.map((_,i)=>P(i,k/5)).join(' ')}" fill="none"/>`;
    cats.forEach((c,i)=>{const[x,y]=P(i,1),[lx,ly]=P(i,1.13);o+=`<line class="gl" x1="${cx}" y1="${cy}" x2="${x}" y2="${y}"/>`+T(lx,ly+3,String(c).slice(0,10))});
    S.forEach((s,i)=>{o+=`<polygon points="${s.v.map((v,j)=>P(j,v/mx)).join(' ')}" fill="${PAL[i%8]}" fill-opacity=".25" stroke="${PAL[i%8]}" stroke-width="2"/>`});
    return o+(showLeg?lg(S.map(s=>s.name)):'')}
  let ser=S.map(s=>s.v),wf=null;
  if(type==='histogram'){const a=S[0].v,k=Math.max(1,Math.ceil(Math.log2(a.length||1)+1)),mn=Math.min(...a),mx=Math.max(...a),w=(mx-mn)/k||1,cnt=Array(k).fill(0);
    a.forEach(x=>cnt[Math.min(k-1,Math.floor((x-mn)/w))]++);ser=[cnt];cats=cnt.map((_,i)=>(mn+i*w).toFixed(1)+'–'+(mn+(i+1)*w).toFixed(1))}
  if(type==='waterfall'){let run=0;wf=S[0].v.map(v=>{const a=run;run+=v;return{lo:Math.min(a,run),hi:Math.max(a,run),v}});wf.push({lo:Math.min(0,run),hi:Math.max(0,run),t:1});cats=cats.concat('Total');ser=[wf.flatMap(b=>[b.lo,b.hi])]}
  const all=ser.flat(),lo=Math.min(0,...all),hi=Math.max(0,...all),tk=nice(hi===lo?lo+1:hi,lo),a0=tk[0],a1=tk[tk.length-1];
  const m={l:48,r:14,t:12,b:showLeg?46:30},pw=W-m.l-m.r,ph=H-m.t-m.b,Y=v=>m.t+ph-(v-a0)/(a1-a0)*ph,n=cats.length,bw=pw/n,X=i=>m.l+bw*(i+.5);
  tk.forEach(v=>{o+=`<line class="gl" x1="${m.l}" x2="${W-m.r}" y1="${Y(v)}" y2="${Y(v)}"/>`+T(m.l-5,Y(v)+3,+v.toFixed(2),'end')});
  o+=`<line class="ax" x1="${m.l}" x2="${W-m.r}" y1="${Y(0)}" y2="${Y(0)}"/>`;
  if(type==='scatter'){const xs=cats.map((c,i)=>isNaN(+c)?i+1:+c),xn=Math.min(...xs),xx=Math.max(...xs),xt=nice(xx===xn?xx+1:xx,xn),b0=xt[0],b1=xt[xt.length-1],XX=v=>m.l+(v-b0)/(b1-b0)*pw;
    xt.forEach(v=>{o+=T(XX(v),H-m.b+14,+v.toFixed(2))});
    S.forEach((s,i)=>s.v.forEach((v,j)=>{o+=`<circle cx="${XX(xs[j])}" cy="${Y(v)}" r="4" fill="${PAL[i%8]}" fill-opacity=".85"/>`}));
    return o+(showLeg?lg(S.map(s=>s.name)):'')}
  const st=Math.ceil(n/Math.max(1,Math.floor(pw/48)));cats.forEach((c,i)=>{if(i%st===0)o+=T(X(i),H-m.b+14,String(c).slice(0,10))});
  const bars=(arrs,gap)=>{const k=arrs.length,gw=bw*(1-gap)/k;arrs.forEach((a,s)=>a.forEach((v,i)=>{o+=`<rect x="${m.l+bw*i+bw*gap/2+gw*s}" y="${Math.min(Y(v),Y(0))}" width="${gw}" height="${Math.abs(Y(v)-Y(0))}" fill="${PAL[s%8]}"/>`}))};
  const line=(a,s,fill)=>{const pts=a.map((v,i)=>`${X(i)},${Y(v)}`).join(' ');
    if(fill)o+=`<polygon points="${X(0)},${Y(0)} ${pts} ${X(a.length-1)},${Y(0)}" fill="${PAL[s%8]}" fill-opacity=".35"/>`;
    o+=`<polyline points="${pts}" fill="none" stroke="${PAL[s%8]}" stroke-width="2.2"/>`+a.map((v,i)=>`<circle cx="${X(i)}" cy="${Y(v)}" r="3" fill="${PAL[s%8]}"/>`).join('')};
  if(type==='bar')bars(ser,.3);
  else if(type==='histogram')bars(ser,.04);
  else if(type==='line')ser.forEach((a,s)=>line(a,s));
  else if(type==='area')ser.forEach((a,s)=>line(a,s,1));
  else if(type==='combo'){bars([ser[0]],.3);ser.slice(1).forEach((a,s)=>line(a,s+1))}
  else if(wf)wf.forEach((b,i)=>{o+=`<rect x="${X(i)-bw*.35}" y="${Y(b.hi)}" width="${bw*.7}" height="${Math.max(1,Y(b.lo)-Y(b.hi))}" fill="${b.t?PAL[0]:b.v>=0?PAL[5]:PAL[1]}"/>`});
  return o+(showLeg?lg(S.map(s=>s.name)):'')}

const charts=()=>sh().charts||(sh().charts=[]);
function paint(c,p,snap){const b=p.querySelector('.xcb'),W=b.clientWidth||c.w,H=b.clientHeight||c.h-26;
  try{b.innerHTML=`<svg viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" style="display:block">${draw(c.type,getData(c,snap),W,H)}</svg>`}catch(e){b.textContent='Cannot chart this range'}}
function mkPanel(c){const p=document.createElement('div');p.className='xchart';p.dataset.id=c.id;
  p.style.cssText=`left:${c.x}px;top:${c.y}px;width:${c.w}px;height:${c.h}px`;
  p.innerHTML=`<div class="xch"><b>${esc(c.title)}</b><select>${CT.map(t=>`<option value="${t[0]}">${t[2]}</option>`).join('')}</select><span class="x" title="Delete chart">✕</span></div><div class="xcb"></div>`;
  const sl=p.querySelector('select');sl.value=c.type;sl.onchange=()=>{c.type=sl.value;API.refresh()};
  p.querySelector('.x').onclick=()=>{const a=charts();a.splice(a.indexOf(c),1);API.refresh()};
  p.querySelector('.xch').onmousedown=e=>{if(e.target.closest('select,.x'))return;const sx=e.clientX,sy=e.clientY,ox=c.x,oy=c.y;
    const mv=ev=>{c.x=Math.max(0,ox+ev.clientX-sx);c.y=Math.max(0,oy+ev.clientY-sy);p.style.left=c.x+'px';p.style.top=c.y+'px'};
    const up=()=>{document.removeEventListener('mousemove',mv);document.removeEventListener('mouseup',up);API.refresh()};
    document.addEventListener('mousemove',mv);document.addEventListener('mouseup',up);e.preventDefault()};
  let first=true,t;new ResizeObserver(()=>{if(first){first=false;return}clearTimeout(t);t=setTimeout(()=>{c.w=p.offsetWidth;c.h=p.offsetHeight;paint(c,p,API.snapshot().sheets[WB().current]||[]);API.refresh()},300)}).observe(p);
  $('#gridwrap').appendChild(p);return p}
function renderCharts(){const host=$('#gridwrap'),list=charts(),snap=API.snapshot().sheets[WB().current]||[];
  host.querySelectorAll('.xchart').forEach(p=>{if(!list.some(c=>c.id===p.dataset.id))p.remove()});
  list.forEach(c=>{let p=host.querySelector(`.xchart[data-id="${c.id}"]`);if(!p)p=mkPanel(c);paint(c,p,snap)})}
function addChart(type){const s=regionSel();if(s.r1===s.r2&&s.c1===s.c2&&raw(s.c1,s.r1)===""){API.toast('Select some data first');return}
  const n=charts().length,g=$('#gridwrap');
  charts().push({id:'ch'+Date.now(),type,...s,x:g.scrollLeft+120+n*24,y:g.scrollTop+50+n*24,w:460,h:300,title:CT.find(t=>t[0]===type)[2]+' chart'});
  API.refresh();API.toast('Chart inserted — drag its title to move, drag the corner to resize')}

/* ================= DATA VALIDATION ================= */
const vals=()=>sh().validations||(sh().validations=[]);
const vFor=(c,r)=>vals().find(v=>r>=v.r1&&r<=v.r2&&c>=v.c1&&c<=v.c2);
function vOK(v,x){if(x===""||x[0]==='=')return true;const n=num(x),ok=(q)=>(v.min==null||q>=v.min)&&(v.max==null||q<=v.max);
  if(v.type==='list')return v.items.includes(x);
  if(v.type==='whole')return !isNaN(n)&&Number.isInteger(+x)&&ok(n);
  if(v.type==='decimal')return !isNaN(n)&&ok(n);
  if(v.type==='length')return ok(x.length);return true}
const shadow={};let chk=false;
function checkVal(){if(chk||!(sh().validations||[]).length)return;const w=WB().current;let bad=0;
  sh().validations.forEach(v=>{for(let r=v.r1;r<=v.r2;r++)for(let c=v.c1;c<=v.c2;c++){const k=w+':'+id(c,r),x=raw(c,r);
    if(vOK(v,x))shadow[k]=x;else{bad++;cellO(c,r).v=shadow[k]!==undefined?shadow[k]:""}}});
  if(bad){chk=true;API.toast(bad+' entry rejected by data validation');API.refresh();chk=false}}
function dvDialog(){const s=sel();
  dlg(`<h3>Data validation</h3><div style="font-size:11px;color:var(--text-dim)">Applies to ${id(s.c1,s.r1)}:${id(s.c2,s.r2)}</div>
  <label>Allow</label><select id="dvT"><option value="list">List (dropdown)</option><option value="whole">Whole number</option><option value="decimal">Decimal</option><option value="length">Text length</option></select>
  <div id="dvA"><label>Items (comma-separated)</label><input id="dvI" placeholder="Yes, No, Maybe"></div>
  <div id="dvB" style="display:none"><label>Minimum (optional)</label><input id="dvMin" type="number"><label>Maximum (optional)</label><input id="dvMax" type="number"></div>
  <div class="dialog-actions"><button class="btn" id="dvX" style="color:var(--danger)">Clear</button><button class="btn" id="dvC">Cancel</button><button class="btn primary" id="dvO">OK</button></div>`);
  $('#dvT').onchange=e=>{const l=e.target.value==='list';$('#dvA').style.display=l?'block':'none';$('#dvB').style.display=l?'none':'block'};
  $('#dvC').onclick=close;
  const drop=()=>{sh().validations=vals().filter(v=>!(v.r1<=s.r2&&v.r2>=s.r1&&v.c1<=s.c2&&v.c2>=s.c1))};
  $('#dvX').onclick=()=>{drop();API.refresh();close()};
  $('#dvO').onclick=()=>{const type=$('#dvT').value,v={...s,type};
    if(type==='list'){v.items=$('#dvI').value.split(',').map(x=>x.trim()).filter(Boolean);if(!v.items.length){API.toast('Enter at least one item');return}}
    else{const a=$('#dvMin').value,b=$('#dvMax').value;v.min=a===''?null:+a;v.max=b===''?null:+b}
    drop();vals().push(v);for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++)shadow[WB().current+':'+id(c,r)]=raw(c,r);
    API.refresh();close();API.toast('Data validation applied')}}
const dv=document.createElement('select');dv.id='xdv';dv.style.display='none';$('#gridwrap').appendChild(dv);
dv.onchange=()=>{const a=API.getActive();API.pushUndo();cellO(a.c,a.r).v=dv.value;API.refresh()};
function dvDropdown(){const a=API.getActive(),v=vFor(a.c,a.r),td=document.querySelector(`#grid td.cell[data-r="${a.r}"][data-c="${a.c}"]`);
  if(!v||v.type!=='list'||!td){dv.style.display='none';return}
  dv.innerHTML='<option value=""></option>'+v.items.map(i=>`<option>${esc(i)}</option>`).join('');dv.value=raw(a.c,a.r);
  Object.assign(dv.style,{display:'block',left:td.offsetLeft+td.offsetWidth-23+'px',top:td.offsetTop+'px',height:td.offsetHeight+'px'})}

/* ================= TEXT TO COLUMNS ================= */
function textToCols(){const s=sel();
  dlg(`<h3>Text to columns</h3><div style="font-size:11px;color:var(--text-dim)">Splits ${id(s.c1,s.r1)}:${id(s.c1,s.r2)} into the columns to its right</div>
  <label>Delimiter</label><select id="ttD"><option value=",">Comma</option><option value=";">Semicolon</option><option value=" ">Space</option><option value="\t">Tab</option><option value="|">Pipe</option><option value="*">Custom…</option></select>
  <input id="ttC" placeholder="Custom delimiter" style="display:none;margin-top:6px">
  <div class="dialog-actions"><button class="btn" id="ttX">Cancel</button><button class="btn primary" id="ttO">Split</button></div>`);
  $('#ttD').onchange=e=>$('#ttC').style.display=e.target.value==='*'?'block':'none';$('#ttX').onclick=close;
  $('#ttO').onclick=()=>{let d=$('#ttD').value;if(d==='*')d=$('#ttC').value;if(!d){API.toast('Enter a delimiter');return}
    API.pushUndo();for(let r=s.r1;r<=s.r2;r++){const v=raw(s.c1,r);if(!v||v[0]==='=')continue;v.split(d).forEach((p,i)=>{if(s.c1+i<COLS)cellO(s.c1+i,r).v=p.trim()})}
    API.refresh();close();API.toast('Text split into columns')}}

/* ================= PIVOT TABLE ================= */
function pivotDialog(){const s=regionSel(),snap=API.snapshot().sheets[WB().current]||[],hdr=[];
  for(let c=s.c1;c<=s.c2;c++)hdr.push(String((snap[s.r1]||[])[c]??L(c)));
  if(s.r2-s.r1<1||hdr.length<2){API.toast('Select a table with a header row and 2+ columns');return}
  const op=(x,i,d)=>`<option value="${i}"${i===d?' selected':''}>${esc(x)}</option>`;
  dlg(`<h3>PivotTable</h3><div style="font-size:11px;color:var(--text-dim)">Source: ${id(s.c1,s.r1)}:${id(s.c2,s.r2)} (first row = headers)</div>
  <label>Rows</label><select id="pvR">${hdr.map((x,i)=>op(x,i,0)).join('')}</select>
  <label>Columns (optional)</label><select id="pvC"><option value="-1">(none)</option>${hdr.map((x,i)=>op(x,i,-9)).join('')}</select>
  <label>Values</label><select id="pvV">${hdr.map((x,i)=>op(x,i,hdr.length-1)).join('')}</select>
  <label>Summarize by</label><select id="pvA"><option>Sum</option><option>Count</option><option>Average</option><option>Min</option><option>Max</option></select>
  <div class="dialog-actions"><button class="btn" id="pvX">Cancel</button><button class="btn primary" id="pvO">Create on new sheet</button></div>`);
  $('#pvX').onclick=close;
  $('#pvO').onclick=()=>{const ri=+$('#pvR').value,ci=+$('#pvC').value,vi=+$('#pvV').value,ag=$('#pvA').value,rows=snap.slice(s.r1+1,s.r2+1);
    const key=(r,i)=>String((r||[])[s.c1+i]??''),val=r=>(r||[])[s.c1+vi];
    const agg=rs=>{const a=rs.map(val),n=a.filter(x=>typeof x==='number');if(ag==='Count')return a.length;if(!n.length)return '';const sum=n.reduce((x,y)=>x+y,0);return ag==='Sum'?sum:ag==='Average'?sum/n.length:ag==='Min'?Math.min(...n):Math.max(...n)};
    const rk=[...new Set(rows.map(r=>key(r,ri)))].sort(),ck=ci<0?[]:[...new Set(rows.map(r=>key(r,ci)))].sort();
    const out=[[hdr[ri],...(ck.length?ck:[ag+' of '+hdr[vi]]),...(ck.length?['Grand Total']:[])]];
    rk.forEach(k=>{const rr=rows.filter(r=>key(r,ri)===k);out.push([k,...(ck.length?ck.map(c=>agg(rr.filter(r=>key(r,ci)===c))):[agg(rr)]),...(ck.length?[agg(rr)]:[])])});
    out.push(['Grand Total',...(ck.length?ck.map(c=>agg(rows.filter(r=>key(r,ci)===c))):[agg(rows)]),...(ck.length?[agg(rows)]:[])]);
    const w=WB();let nm='Pivot',k=1;while(w.sheets[nm]){k++;nm='Pivot'+k}
    const ns={data:{},colWidths:{},rowHeights:{},rows:Math.max(100,out.length+5),merges:[],frozenRow:0,frozenCol:0};
    out.forEach((row,r)=>row.forEach((v,c)=>{if(c<COLS&&v!=="")ns.data[L(c)+(r+1)]={v:String(typeof v==='number'?+v.toFixed(2):v),f:(r===0||r===out.length-1||c===row.length-1&&ck.length)?{bold:true}:{}}}));
    if(ns.data.A1)ns.data.A1.f={bold:true};
    w.sheets[nm]=ns;w.order.push(nm);w.current=nm;close();API.refresh();API.toast('PivotTable created on '+nm)}}

/* ================= FREEZE PANES + OBSERVER ================= */
function freeze(){const s=sh(),g=$('#grid');
  if(s.frozenRow)g.querySelectorAll('tbody tr:first-child > *').forEach(e=>{e.style.position='sticky';e.style.top='26px';e.style.zIndex=e.classList.contains('rowhead')?26:12;if(e.tagName==='TD'&&!e.style.background)e.style.background='var(--surface)'});
  if(s.frozenCol)g.querySelectorAll('tbody tr').forEach(tr=>{const e=tr.children[1];if(e){e.style.position='sticky';e.style.left='var(--rowhead-w)';e.style.zIndex=Math.max(+e.style.zIndex||0,11);if(!e.style.background)e.style.background='var(--surface)'}})}
let busy=false,raf=0,lastSheet=null;
function sync(){busy=true;
  if(WB().current!==lastSheet){lastSheet=WB().current;renderCharts()}
  document.querySelectorAll('#grid td.cell .cv').forEach(e=>{const t=e.textContent;if(t.startsWith('SPARK|')){const p=t.split('|');e.innerHTML=sparkSvg(p[1],p[2].split(',').map(Number));e.parentNode.style.padding='1px 3px'}});
  freeze();dvDropdown();obs.takeRecords();busy=false}
const obs=new MutationObserver(()=>{if(busy)return;cancelAnimationFrame(raf);raf=requestAnimationFrame(sync)});
obs.observe($('#grid'),{childList:true,subtree:true});
const prev=window.__psOnChange;let ct;
window.__psOnChange=()=>{if(prev)prev();checkVal();clearTimeout(ct);ct=setTimeout(renderCharts,120)};
sync();
window.ProSheetCharts={draw,regionSel};
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
