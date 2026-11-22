/* Pro Sheet — Excel Pro pack
   Adds: 45+ new formulas (SUMIFS, COUNTIFS, PMT, NPV, IRR…), formula autocomplete + argument hints,
   Insert Function dialog, AutoFilter, multi-level Custom Sort, Format as Table, Cell Styles,
   Data Bars / Color Scales / Icon Sets / Top-Bottom / Above-Below average / Duplicates,
   Notes (red-corner comments), Format Painter, ₹/%/comma/accounting/scientific + inc/dec decimals,
   Borders menu, Find & Replace (find next), Freeze at selection, row height / column width / AutoFit,
   Show/Hide gridlines, headings, formula bar, and shortcuts (Ctrl+A, Ctrl+F/H, F5, Shift+F2, Ctrl+Shift+L).
   Load AFTER excel-ui.js:  <script src="excel-pro.js"></script>  */
(function(){
"use strict";
function boot(){
const API=window.ProSheetAPI;
if(!API||!API.getWorkbook||!API.FN||!window.ProSheetCharts||$$('.xl-rib .xp').length<6)return setTimeout(boot,250);

/* ---------- helpers ---------- */
const L=API.colLetter,COLS=26,WB=()=>API.getWorkbook(),sh=()=>WB().sheets[WB().current],id=(c,r)=>L(c)+(r+1);
const raw=(c,r)=>{const x=sh().data[id(c,r)];return x&&x.v!==undefined?String(x.v):""};
const cellO=(c,r)=>{const d=sh().data;return d[id(c,r)]||(d[id(c,r)]={v:"",f:{}})};
const sel=()=>{const s=API.getSelection();return{r1:Math.min(s.r1,s.r2),r2:Math.max(s.r1,s.r2),c1:Math.min(s.c1,s.c2),c2:Math.max(s.c1,s.c2)}};
const esc=s=>String(s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const dlg=h=>{$('#dialogBox').innerHTML=h;$('#dialogOverlay').classList.add('show')},close=()=>$('#dialogOverlay').classList.remove('show');
const toast=API.toast,snapRows=()=>API.snapshot().sheets[WB().current]||[];
const tdAt=(r,c)=>document.querySelector(`#grid td.cell[data-r="${r}"][data-c="${c}"]`);
const eachSel=cb=>{const s=sel();for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++)cb(c,r)};
const applyF=(fn,msg)=>{API.pushUndo();eachSel((c,r)=>{const x=cellO(c,r);x.f=x.f||{};fn(x.f,c,r)});API.refresh();if(msg)toast(msg)};
const region=()=>window.ProSheetCharts.regionSel();
function $(s){return document.querySelector(s)}function $$(s){return[...document.querySelectorAll(s)]}

/* ---------- new formulas ---------- */
const F=API.FN,fl=a=>[a].flat(Infinity),nums=a=>fl(a).map(parseFloat).filter(x=>!isNaN(x));
function mt(v,crit){const m=String(crit).match(/^(>=|<=|<>|>|<|=)?([\s\S]*)$/),op=m[1]||'=',rhs=m[2],a=parseFloat(v),b=parseFloat(rhs);
  if(!isNaN(a)&&!isNaN(b)&&String(v).trim()!=='')return op==='>'?a>b:op==='<'?a<b:op==='>='?a>=b:op==='<='?a<=b:op==='<>'?a!==b:a===b;
  const s=String(v).toLowerCase(),t=rhs.toLowerCase();
  if(/[*?]/.test(t)){const re=new RegExp('^'+t.replace(/[.+^${}()|[\]\\]/g,'\\$&').replace(/\*/g,'.*').replace(/\?/g,'.')+'$');return op==='<>'?!re.test(s):re.test(s)}
  return op==='<>'?s!==t:s===t}
const pairs=a=>{const r=[],c=[];for(let i=0;i<a.length;i+=2){r.push(fl(a[i]));c.push(a[i+1])}return[r,c]};
const hits=(r,c)=>{const o=[];for(let i=0;i<r[0].length;i++)if(r.every((x,k)=>mt(x[i],c[k])))o.push(i);return o};
Object.assign(F,{
COUNTIFS:(...a)=>{const[r,c]=pairs(a);return hits(r,c).length},
SUMIFS:(s,...a)=>{const[r,c]=pairs(a),v=fl(s);return hits(r,c).reduce((t,i)=>t+(parseFloat(v[i])||0),0)},
AVERAGEIFS:(s,...a)=>{const[r,c]=pairs(a),v=fl(s),h=hits(r,c);return h.length?h.reduce((t,i)=>t+(parseFloat(v[i])||0),0)/h.length:'#DIV/0!'},
MAXIFS:(s,...a)=>{const[r,c]=pairs(a),v=fl(s);const x=hits(r,c).map(i=>parseFloat(v[i])).filter(n=>!isNaN(n));return x.length?Math.max(...x):0},
MINIFS:(s,...a)=>{const[r,c]=pairs(a),v=fl(s);const x=hits(r,c).map(i=>parseFloat(v[i])).filter(n=>!isNaN(n));return x.length?Math.min(...x):0},
SUMPRODUCT:(...a)=>{const A=a.map(x=>fl(x).map(v=>parseFloat(v)||0));return A[0].reduce((t,_,i)=>t+A.reduce((p,r)=>p*(r[i]||0),1),0)},
PRODUCT:(...a)=>nums(a).reduce((x,y)=>x*y,1),SUMSQ:(...a)=>nums(a).reduce((x,y)=>x+y*y,0),
LARGE:(r,k)=>{const n=nums(r).sort((a,b)=>b-a);return n[k-1]??'#NUM!'},SMALL:(r,k)=>{const n=nums(r).sort((a,b)=>a-b);return n[k-1]??'#NUM!'},
RANK:(n,r,o)=>{const a=nums(r);return a.filter(x=>o?x<n:x>n).length+1},
CHOOSE:(i,...a)=>a[i-1]??'#VALUE!',
ISNUMBER:v=>typeof v==='number'||(v!==''&&!isNaN(Number(v))),ISTEXT:v=>typeof v==='string'&&isNaN(Number(v)),ISERROR:v=>String(v).startsWith('#'),
IFNA:(v,f)=>v==='#N/A'?f:v,NA:()=>'#N/A',
VALUE:v=>{const n=parseFloat(String(v).replace(/[₹$,%\s]/g,''));return isNaN(n)?'#VALUE!':String(v).includes('%')?n/100:n},
EXACT:(a,b)=>String(a)===String(b),CLEAN:s=>String(s).replace(/[\x00-\x1f]/g,''),CHAR:n=>String.fromCharCode(n),
CEILING:(n,s=1)=>Math.ceil(n/s)*s,FLOOR:(n,s=1)=>Math.floor(n/s)*s,EXP:Math.exp,LN:Math.log,LOG:(n,b=10)=>Math.log(n)/Math.log(b),SIGN:Math.sign,
RAND:()=>Math.random(),RANDBETWEEN:(a,b)=>Math.floor(Math.random()*(b-a+1))+a,
PMT:(r,n,pv,fv=0,t=0)=>r===0?-(pv+fv)/n:-(r*(pv*Math.pow(1+r,n)+fv))/((1+r*t)*(Math.pow(1+r,n)-1)),
FV:(r,n,p,pv=0,t=0)=>r===0?-(pv+p*n):-(pv*Math.pow(1+r,n)+p*(1+r*t)*(Math.pow(1+r,n)-1)/r),
PV:(r,n,p,fv=0,t=0)=>r===0?-(fv+p*n):-(fv+p*(1+r*t)*(Math.pow(1+r,n)-1)/r)/Math.pow(1+r,n),
NPV:(r,...a)=>nums(a).reduce((t,v,i)=>t+v/Math.pow(1+r,i+1),0),
IRR:v=>{const c=nums(v);let g=.1;for(let i=0;i<100;i++){let f=0,d=0;c.forEach((x,k)=>{f+=x/Math.pow(1+g,k);d-=k*x/Math.pow(1+g,k+1)});const n=g-f/d;if(!isFinite(n))return'#NUM!';if(Math.abs(n-g)<1e-9)return n;g=n}return g},
TEXT:(v,f)=>{f=String(f);const n=parseFloat(v);
  if(/[ymdhs]/i.test(f)&&!/^[#0.,%\s]+$/.test(f)){const d=new Date(v);if(isNaN(d))return String(v);const p=x=>String(x).padStart(2,'0');
    return f.replace(/yyyy|yy|mmmm|mmm|mm|dd|hh|ss|m|d/gi,t=>{switch(t.toLowerCase()){case'yyyy':return d.getFullYear();case'yy':return String(d.getFullYear()).slice(2);case'mmmm':return d.toLocaleString('en',{month:'long'});case'mmm':return d.toLocaleString('en',{month:'short'});case'mm':return p(d.getMonth()+1);case'dd':return p(d.getDate());case'hh':return p(d.getHours());case'ss':return p(d.getSeconds());case'm':return d.getMonth()+1;default:return d.getDate()}})}
  if(isNaN(n))return String(v);const pct=f.includes('%'),dec=(f.split('.')[1]||'').replace(/[^0#]/g,'').length,x=pct?n*100:n;let s=Math.abs(x).toFixed(dec);
  if(f.includes(','))s=Number(s).toLocaleString('en-US',{minimumFractionDigits:dec,maximumFractionDigits:dec});
  return(x<0?'-':'')+f.match(/^[^#0,.]*/)[0].replace(/"/g,'')+s+(pct?'%':'')}
});

/* ---------- function metadata (Insert Function + autocomplete) ---------- */
const META=`SUM|number1, [number2], …|Adds numbers
AVERAGE|number1, …|Average of the numbers
COUNT|value1, …|Counts cells with numbers
COUNTA|value1, …|Counts non-empty cells
COUNTBLANK|range|Counts empty cells
MIN|number1, …|Smallest value
MAX|number1, …|Largest value
MEDIAN|number1, …|Middle value
MODE|number1, …|Most frequent value
STDEV|number1, …|Sample standard deviation
STDEVP|number1, …|Population standard deviation
VARIANCE|number1, …|Sample variance
VARP|number1, …|Population variance
LARGE|array, k|k-th largest value
SMALL|array, k|k-th smallest value
RANK|number, ref, [order]|Rank of a number in a list
PRODUCT|number1, …|Multiplies numbers
SUMSQ|number1, …|Sum of squares
SUMPRODUCT|array1, [array2], …|Sum of products of ranges
ROUND|number, digits|Rounds to a number of digits
ROUNDUP|number, digits|Rounds away from zero
ROUNDDOWN|number, digits|Rounds toward zero
TRUNC|number, [digits]|Cuts off decimals
INT|number|Rounds down to an integer
CEILING|number, [significance]|Rounds up to a multiple
FLOOR|number, [significance]|Rounds down to a multiple
MOD|number, divisor|Remainder after division
ABS|number|Absolute value
SIGN|number|-1, 0 or 1
SQRT|number|Square root
POWER|number, power|Raises to a power
EXP|number|e raised to a power
LN|number|Natural logarithm
LOG|number, [base]|Logarithm
PI||Value of pi
RAND||Random number between 0 and 1
RANDBETWEEN|bottom, top|Random whole number
IF|condition, then, else|One value if true, another if false
IFS|cond1, value1, cond2, value2, …|First matching condition
IFERROR|value, fallback|Fallback when value is an error
IFNA|value, fallback|Fallback for #N/A
AND|logical1, …|TRUE if all are true
OR|logical1, …|TRUE if any is true
NOT|logical|Reverses TRUE / FALSE
ISNUMBER|value|TRUE if a number
ISTEXT|value|TRUE if text
ISERROR|value|TRUE if an error
COUNTIF|range, criteria|Counts cells meeting a condition
COUNTIFS|range1, crit1, …|Counts cells meeting all conditions
SUMIF|range, criteria, [sum_range]|Sums cells meeting a condition
SUMIFS|sum_range, range1, crit1, …|Sums cells meeting all conditions
AVERAGEIF|range, criteria, [avg_range]|Average where a condition is met
AVERAGEIFS|avg_range, range1, crit1, …|Average where all conditions are met
MAXIFS|max_range, range1, crit1, …|Largest value meeting conditions
MINIFS|min_range, range1, crit1, …|Smallest value meeting conditions
VLOOKUP|lookup, table, col, [exact]|Finds a value in the first column
HLOOKUP|lookup, table, row, [exact]|Finds a value in the first row
XLOOKUP|lookup, lookup_array, return_array|Modern lookup
INDEX|array, row, [col]|Value at a position
MATCH|lookup, array|Position of a value
CHOOSE|index, value1, value2, …|Picks a value by index
CONCAT|text1, …|Joins text
CONCATENATE|text1, …|Joins text
TEXTJOIN|delimiter, ignore_empty, text1, …|Joins text with a delimiter
TEXT|value, format|Formats as text (0.00, #,##0, 0%, dd-mm-yyyy)
VALUE|text|Converts text to a number
LEN|text|Number of characters
UPPER|text|Uppercase
LOWER|text|Lowercase
PROPER|text|Capitalizes each word
TRIM|text|Removes extra spaces
CLEAN|text|Removes non-printing characters
EXACT|text1, text2|Case-sensitive comparison
LEFT|text, [n]|First characters
RIGHT|text, [n]|Last characters
MID|text, start, length|Characters from the middle
REPT|text, times|Repeats text
SUBSTITUTE|text, old, new|Replaces text
FIND|find, within, [start]|Position (case-sensitive)
SEARCH|find, within, [start]|Position (ignores case)
CHAR|number|Character from a code
TODAY||Today's date
NOW||Current date and time
DATE|year, month, day|Builds a date
YEAR|date|Year of a date
MONTH|date|Month of a date
DAY|date|Day of a date
WEEKDAY|date, [type]|Day of the week
EDATE|date, months|Date n months away
EOMONTH|date, months|End of a month
DATEDIF|start, end, unit (D/M/Y)|Difference between dates
NETWORKDAYS|start, end|Working days between dates
PMT|rate, nper, pv, [fv], [type]|Loan payment per period
FV|rate, nper, pmt, [pv], [type]|Future value
PV|rate, nper, pmt, [fv], [type]|Present value
NPV|rate, value1, …|Net present value
IRR|values|Internal rate of return
NA||Returns #N/A
SPARKLINE|range, type|Mini chart in a cell (line, column, winloss)`.split('\n').map(l=>l.split('|'));
const sigOf=n=>(META.find(m=>m[0]===n)||[])[1];

/* ---------- styles ---------- */
const css=document.createElement('style');
css.textContent=`
#xac{position:fixed;z-index:700;display:none;min-width:250px;max-width:360px;background:var(--surface);color:var(--text);border:1px solid var(--border);box-shadow:0 4px 14px rgba(0,0,0,.25);font:12px 'Segoe UI','DM Sans',sans-serif}
#xac .h{padding:5px 8px;background:var(--header-bg);border-bottom:1px solid var(--border);color:var(--text-dim)}#xac .h b{color:var(--text)}
#xac .i{padding:4px 8px;cursor:pointer;display:flex;gap:10px;align-items:baseline}#xac .i small{color:var(--text-faint);overflow:hidden;text-overflow:ellipsis;white-space:nowrap}#xac .i.on{background:var(--accent-soft)}
.xaf{position:absolute;right:2px;top:50%;transform:translateY(-50%);font-size:8px;line-height:12px;border:1px solid #888;background:#f3f3f3;color:#333;padding:0 3px;cursor:pointer;z-index:3}.xaf.on{background:#217346;color:#fff;border-color:#217346}
td.xnote::after{content:'';position:absolute;top:0;right:0;border-top:7px solid #d00;border-left:7px solid transparent}
body.xnogrid td.cell{border-right-color:transparent!important;border-bottom-color:transparent!important}
body.xnohead .colhead,body.xnohead .rowhead,body.xnohead .corner{display:none}
body.xnofb .formulabar{display:none}
.xfl{max-height:230px;overflow:auto;margin-top:8px;border:1px solid var(--border);border-radius:6px}
.xfl div{padding:5px 9px;font-size:12.5px;cursor:pointer;display:flex;justify-content:space-between;gap:8px}.xfl div:hover{background:var(--accent-soft)}.xfl div.on{background:var(--accent);color:#fff}
.dialog textarea{width:100%;background:var(--surface);color:var(--text);border:1px solid var(--border);border-radius:7px;padding:8px;font:12.5px 'DM Sans',sans-serif;user-select:text}
.dialog .row{display:flex;gap:12px;margin-top:10px;font-size:12px}.dialog .row label{display:flex;gap:5px;margin:0;align-items:center}.dialog .row input{width:auto}`;
document.head.appendChild(css);

/* ---------- ribbon helpers ---------- */
const tabs=$$('.xl-tabs .xt'),P=n=>$$('.xl-rib .xp')[tabs.findIndex(t=>t.textContent===n)];
const mkb=(i,l,fn)=>{const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${i}</i><span>${l}</span>`;b.onclick=e=>{e.stopPropagation();fn()};return b};
const mks=(i,l,o,fn)=>{const b=document.createElement('div');b.className='xb';b.innerHTML=`<i>${i}</i><select><option value="">${l}</option>${o.map(x=>`<option>${x}</option>`).join('')}</select>`;const s=b.querySelector('select');s.onchange=()=>{if(s.value)fn(s.value);s.value=''};return b};
const grp=(p,l,items)=>{const g=document.createElement('div');g.className='xg';const r=document.createElement('div');r.className='xr';items.forEach(x=>r.appendChild(x));g.appendChild(r);g.insertAdjacentHTML('beforeend',`<div class="xl">${l}</div>`);P(p).appendChild(g)};
P('Home').style.flexWrap='wrap';

/* ---------- number formats ---------- */
function fmtX(n,x){const d=x.d==null?2:x.d,ab=Math.abs(n),s=x.s||'₹',num=v=>x.g!==false?v.toLocaleString(undefined,{minimumFractionDigits:d,maximumFractionDigits:d}):v.toFixed(d);
  switch(x.k){case'percent':return(n*100).toFixed(d)+'%';case'currency':return(n<0?'-':'')+s+num(ab);case'accounting':return n<0?'('+s+num(ab)+')':s+' '+num(n)+' ';case'sci':return n.toExponential(d).toUpperCase();default:return(n<0?'-':'')+num(ab)}}
const setX=(k,e)=>applyF(f=>{f.xf=Object.assign({k,d:2,g:true},e);f.numFmt='general'});
function incDec(dl){applyF((f,c,r)=>{let x=f.xf;if(!x){const n=f.numFmt;x=n==='currency'?{k:'currency',d:2,g:true}:n==='percent'?{k:'percent',d:1}:n==='number'?{k:'number',d:2,g:true}:{k:'number',d:null,g:false};
  if(x.d==null){const t=(tdAt(r,c)||{}).textContent||'';x.d=(t.split('.')[1]||'').replace(/\D.*$/,'').length}}
  x.d=Math.max(0,Math.min(10,x.d+dl));f.xf=x;f.numFmt='general'})}
const MF={General:()=>applyF(f=>{delete f.xf;f.numFmt='general'}),Number:()=>setX('number'),Comma:()=>setX('number'),'Currency ₹':()=>setX('currency',{s:'₹'}),'Currency $':()=>setX('currency',{s:'$'}),Accounting:()=>setX('accounting'),Percent:()=>setX('percent',{d:0}),Scientific:()=>setX('sci'),'Short date':()=>applyF(f=>{delete f.xf;f.numFmt='date'})};
grp('Home','Number+',[mkb('₹','Currency',()=>setX('currency',{s:'₹'})),mkb('%','Percent',()=>setX('percent',{d:0})),mkb(',','Comma',()=>setX('number')),mkb('.00→','+ Decimal',()=>incDec(1)),mkb('←.0','− Decimal',()=>incDec(-1)),mks('123','More formats',Object.keys(MF),k=>MF[k]())]);

/* ---------- borders, cell styles, table styles ---------- */
const BD={thin:'1px solid #444',thick:'2px solid #000',dbl:'3px double #000'},BN={t:'top',r:'right',b:'bottom',l:'left'};
function border(k){const s=sel();applyF((f,c,r)=>{if(k==='none'){f.bd={};f.border='';return}
  const bd=f.bd=f.bd||{},T=r===s.r1,B=r===s.r2,Lf=c===s.c1,R=c===s.c2;
  if(k==='all')['t','r','b','l'].forEach(x=>bd[x]=BD.thin);
  if(k==='bottom'&&B)bd.b=BD.thin;if(k==='top'&&T)bd.t=BD.thin;if(k==='left'&&Lf)bd.l=BD.thin;if(k==='right'&&R)bd.r=BD.thin;
  if(k==='outside'||k==='thick'){const v=k==='thick'?BD.thick:BD.thin;if(T)bd.t=v;if(B)bd.b=v;if(Lf)bd.l=v;if(R)bd.r=v}
  if(k==='dbl'&&B)bd.b=BD.dbl;
  f.border=Object.entries(bd).map(([a,v])=>`border-${BN[a]}:${v};`).join('')})}
const BO={'Bottom border':'bottom','Top border':'top','Left border':'left','Right border':'right','Outside borders':'outside','Thick box':'thick','Double bottom':'dbl','All borders':'all','No border':'none'};
const CS={Normal:{},Good:{bg:'#c6efce',color:'#006100'},Bad:{bg:'#ffc7ce',color:'#9c0006'},Neutral:{bg:'#ffeb9c',color:'#9c5700'},Input:{bg:'#ffcc99',color:'#3f3f76'},Calculation:{bg:'#f2f2f2',color:'#fa7d00',bold:true},Note:{bg:'#ffffcc',color:'#222'},Title:{bold:true,fontSize:20,color:'#1f3864'},'Heading 1':{bold:true,fontSize:14,color:'#1f3864'},Total:{bold:true,border:'border-top:1px solid #000;border-bottom:3px double #000;'}};
const TS={Blue:['#4472c4','#d9e2f3'],Green:['#70ad47','#e2efd9'],Orange:['#ed7d31','#fbe5d6'],Gold:['#ffc000','#fff2cc'],Gray:['#7f7f7f','#ededed'],Navy:['#203864','#d6dce5']};
function fmtTable(n){const[h,b]=TS[n],s=region();API.pushUndo();
  for(let r=s.r1;r<=s.r2;r++)for(let c=s.c1;c<=s.c2;c++){const x=cellO(c,r);x.f=x.f||{};const hd=r===s.r1;
    Object.assign(x.f,hd?{bold:true,bg:h,color:'#fff'}:{bold:false,bg:(r-s.r1)%2?b:'#ffffff',color:'#222'});x.f.border=`border-bottom:1px solid ${h};`}
  sh().af={r:s.r1,c1:s.c1,c2:s.c2,r2:s.r2,f:{}};API.refresh();toast('Formatted as table '+n+' — filter buttons added')}
grp('Home','Styles',[mkb('🎨','Cond. Format',cfDlg),mks('▦','Format as Table',Object.keys(TS),fmtTable),mks('Aa','Cell Styles',Object.keys(CS),k=>applyF(f=>{['bold','color','bg','fontSize','border','bd'].forEach(x=>delete f[x]);Object.assign(f,JSON.parse(JSON.stringify(CS[k])))})),mks('▭','Borders',Object.keys(BO),k=>border(BO[k]))]);

/* ---------- conditional formatting (visual rules) ---------- */
const CFN={bar:'Data bars',scale:'Color scales',icons:'Icon sets',top:'Top / Bottom',avg:'Above / Below average',dup:'Duplicate / Unique'};
const CFV={bar:['Blue','Green','Orange','Red'],scale:['Green–Yellow–Red','Red–Yellow–Green','White–Blue','White–Red','Blue–White–Red'],icons:['3 Arrows','3 Traffic lights','3 Symbols'],top:['Top 10','Top 5','Bottom 10','Bottom 5','Top 10%'],avg:['Above average','Below average'],dup:['Duplicate values','Unique values']};
const BARC={Blue:'#9dc3e6',Green:'#a9d18e',Orange:'#f4b183',Red:'#ff9b9b'};
const SC={'Green–Yellow–Red':['#f8696b','#ffeb84','#63be7b'],'Red–Yellow–Green':['#63be7b','#ffeb84','#f8696b'],'White–Blue':['#ffffff','#5a8ac6'],'White–Red':['#ffffff','#f8696b'],'Blue–White–Red':['#5a8ac6','#ffffff','#f8696b']};
const IC={'3 Arrows':['▼','▶','▲'],'3 Traffic lights':['●','●','●'],'3 Symbols':['✖','!','✔']},ICC=['#c00000','#e6a800','#2e9d3b'];
const rgb=h=>[1,3,5].map(i=>parseInt(h.slice(i,i+2),16));
function colorAt(st,t){const p=Math.max(0,Math.min(1,t))*(st.length-1),i=Math.min(Math.floor(p),st.length-2),a=rgb(st[i]),b=rgb(st[i+1]),k=p-i;return`rgb(${a.map((v,j)=>Math.round(v+(b[j]-v)*k)).join(',')})`}
const pnum=t=>{t=String(t).trim();if(!/\d/.test(t)||/^\d{4}-\d{2}-\d{2}/.test(t)||/[A-DF-Za-df-z]{2,}/.test(t))return NaN;const v=parseFloat(t.replace(/[^0-9.eE]/g,''));return/^\(.*\)$/.test(t)||/^-/.test(t)?-Math.abs(v):v};
function cfDlg(){const s=sel();
  dlg(`<h3>Conditional formatting</h3><div style="font-size:11px;color:var(--text-dim)">Applies to ${id(s.c1,s.r1)}:${id(s.c2,s.r2)}</div><label>Rule type</label><select id="cfK">${Object.keys(CFN).map(k=>`<option value="${k}">${CFN[k]}</option>`).join('')}</select><label>Style</label><select id="cfV"></select>
  <div class="dialog-actions"><button class="btn" id="cfX" style="color:var(--danger)">Clear here</button><button class="btn" id="cfC">Cancel</button><button class="btn primary" id="cfO">Apply</button></div>`);
  const fill=()=>$('#cfV').innerHTML=CFV[$('#cfK').value].map(v=>`<option>${v}</option>`).join('');fill();$('#cfK').onchange=fill;$('#cfC').onclick=close;
  $('#cfO').onclick=()=>{(sh().cfx=sh().cfx||[]).push({...s,k:$('#cfK').value,v:$('#cfV').value});close();API.refresh();toast('Rule applied')};
  $('#cfX').onclick=()=>{sh().cfx=(sh().cfx||[]).filter(R=>!(R.r1<=s.r2&&R.r2>=s.r1&&R.c1<=s.c2&&R.c2>=s.c1));close();API.refresh()}}
function cfRun(R,map){const cs=[];for(let r=R.r1;r<=R.r2;r++)for(let c=R.c1;c<=R.c2;c++){const td=map[r+','+c];if(td)cs.push({td,t:td.textContent,n:pnum(td.textContent)})}
  const ns=cs.map(x=>x.n).filter(isFinite);if(R.k!=='dup'&&!ns.length)return;
  const mn=Math.min(...ns),mx=Math.max(...ns),rg=(mx-mn)||1,mark=(td,g)=>{td.style.background=g?'#c6efce':'#ffc7ce';td.style.color=g?'#006100':'#9c0006'};
  let thr,top,mean=ns.reduce((a,b)=>a+b,0)/(ns.length||1),cnt={};
  if(R.k==='top'){const m=R.v.match(/(Top|Bottom) (\d+)(%?)/);top=m[1]==='Top';let k=+m[2];if(m[3])k=Math.max(1,Math.ceil(ns.length*k/100));thr=[...ns].sort((a,b)=>top?b-a:a-b)[Math.min(k,ns.length)-1]}
  if(R.k==='dup')cs.forEach(x=>{if(x.t!=='')cnt[x.t]=(cnt[x.t]||0)+1});
  cs.forEach(({td,t,n})=>{const ok=isFinite(n);
    if(R.k==='bar'&&ok){const lo=Math.min(0,mn),w=(n-lo)/((mx-lo)||1)*100;td.style.background=`linear-gradient(90deg,${BARC[R.v]} ${w}%,transparent ${w}%)`}
    else if(R.k==='scale'&&ok)td.style.background=colorAt(SC[R.v],(n-mn)/rg);
    else if(R.k==='icons'&&ok){const q=(n-mn)/rg,i=q>=.67?2:q>=.33?1:0,cv=td.querySelector('.cv');if(cv)cv.innerHTML=`<span style="color:${ICC[i]}">${IC[R.v][i]}</span> `+esc(t)}
    else if(R.k==='top'&&ok&&(top?n>=thr:n<=thr))mark(td,top);
    else if(R.k==='avg'&&ok&&(R.v[0]==='A'?n>mean:n<mean))mark(td,R.v[0]==='A');
    else if(R.k==='dup'&&t!==''&&(R.v[0]==='D'?cnt[t]>1:cnt[t]===1))mark(td,R.v[0]==='U')})}

/* ---------- notes ---------- */
function noteDlg(){const a=API.getActive(),k=id(a.c,a.r),n=sh().notes=sh().notes||{};
  dlg(`<h3>Note — ${k}</h3><textarea id="ntT" rows="5">${esc(n[k]||'')}</textarea><div class="dialog-actions"><button class="btn" id="ntD" style="color:var(--danger)">Delete</button><button class="btn" id="ntC">Cancel</button><button class="btn primary" id="ntO">Save</button></div>`);
  $('#ntT').focus();$('#ntC').onclick=close;$('#ntD').onclick=()=>{delete n[k];close();API.refresh()};
  $('#ntO').onclick=()=>{const v=$('#ntT').value.trim();if(v)n[k]=v;else delete n[k];close();API.refresh();toast('Note saved')}}

/* ---------- filter & sort ---------- */
const AF=()=>sh().af;
function sortBlock(r1,r2,c1,c2,keys){const rows=snapRows(),list=[];API.pushUndo();
  for(let r=r1;r<=r2;r++){const cells=[];for(let c=c1;c<=c2;c++){const x=sh().data[id(c,r)];cells.push(x?JSON.parse(JSON.stringify(x)):{v:"",f:{}})}list.push({cells,vals:rows[r]||[]})}
  list.sort((p,q)=>{for(const k of keys){const A=p.vals[k.c],B=q.vals[k.c],ba=A===''||A==null,bb=B===''||B==null;if(ba||bb){if(ba!==bb)return ba?1:-1;continue}
    const n=Number(A),m=Number(B),d=!isNaN(n)&&!isNaN(m)?n-m:String(A).localeCompare(String(B));if(d)return k.asc?d:-d}return 0});
  list.forEach((it,i)=>it.cells.forEach((x,j)=>{sh().data[id(c1+j,r1+i)]=x}));API.refresh()}
function reapply(){const a=AF(),rows=snapRows(),h=sh().hiddenRows=sh().hiddenRows||{};a.r2=Math.max(a.r2,rows.length-1);
  for(let r=a.r+1;r<=a.r2;r++){let ok=true;for(const c in a.f)if(!a.f[c].includes(String((rows[r]||[])[c]??''))){ok=false;break}if(ok)delete h[r];else h[r]=true}API.refresh()}
function toggleAF(){const a=AF();if(a){const h=sh().hiddenRows||{};for(let r=a.r+1;r<=a.r2;r++)delete h[r];sh().af=null;API.refresh();toast('Filter removed');return}
  const s=region();if(s.r2<=s.r1&&s.c1===s.c2){toast('Select a table first');return}sh().af={r:s.r1,c1:s.c1,c2:s.c2,r2:s.r2,f:{}};API.refresh();toast('Filter on — use the ▼ in the header row')}
function afMenu(c){const a=AF(),rows=snapRows(),cur=a.f[c];
  const vals=[...new Set(rows.slice(a.r+1,a.r2+1).map(r=>String((r||[])[c]??'')))].sort((x,y)=>isNaN(x-y)?x.localeCompare(y):x-y);
  dlg(`<h3>Filter: ${esc(String((rows[a.r]||[])[c]??L(c)))}</h3><div style="display:flex;gap:6px"><button class="btn" id="afA">Sort A → Z</button><button class="btn" id="afZ">Z → A</button></div>
  <label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="afAll" style="width:auto"> (Select all)</label>
  <div class="xfl" style="padding:6px">${vals.map(v=>`<label style="display:flex;gap:6px;margin:2px 0;align-items:center"><input type="checkbox" style="width:auto" value="${esc(v)}"${!cur||cur.includes(v)?' checked':''}>${esc(v)||'(Blanks)'}</label>`).join('')}</div>
  <div class="dialog-actions"><button class="btn" id="afX">Clear filter</button><button class="btn primary" id="afO">OK</button></div>`);
  const boxes=()=>$$('.xfl input'),sync=()=>$('#afAll').checked=boxes().every(b=>b.checked);sync();boxes().forEach(b=>b.onchange=sync);
  $('#afAll').onchange=e=>boxes().forEach(b=>b.checked=e.target.checked);
  $('#afO').onclick=()=>{const ch=boxes().filter(b=>b.checked).map(b=>b.value);if(ch.length===boxes().length)delete a.f[c];else a.f[c]=ch;close();reapply()};
  $('#afX').onclick=()=>{delete a.f[c];close();reapply()};
  const srt=asc=>{close();sortBlock(a.r+1,a.r2,a.c1,a.c2,[{c,asc}]);reapply()};$('#afA').onclick=()=>srt(true);$('#afZ').onclick=()=>srt(false)}
$('#grid').addEventListener('mousedown',e=>{const b=e.target.closest('.xaf');if(b){e.stopPropagation();e.preventDefault();afMenu(+b.dataset.c)}},true);
function sortDlg(){const s=region(),rows=snapRows(),hd=s.r1<s.r2&&(rows[s.r1]||[]).slice(s.c1,s.c2+1).some(v=>typeof v==='string'&&v!=='');
  const opt=(i)=>`<select id="so${i}">${i>1?'<option value="-1">(none)</option>':''}${[...Array(s.c2-s.c1+1)].map((_,k)=>`<option value="${s.c1+k}">${L(s.c1+k)}${rows[s.r1]&&rows[s.r1][s.c1+k]!==''&&rows[s.r1][s.c1+k]!=null?' — '+esc(String(rows[s.r1][s.c1+k])):''}</option>`).join('')}</select><select id="sd${i}" style="margin-top:4px"><option value="1">A → Z / Smallest first</option><option value="0">Z → A / Largest first</option></select>`;
  dlg(`<h3>Custom sort</h3><div style="font-size:11px;color:var(--text-dim)">${id(s.c1,s.r1)}:${id(s.c2,s.r2)}</div><label style="display:flex;gap:6px;align-items:center"><input type="checkbox" id="soH" style="width:auto"${hd?' checked':''}> My data has headers</label>
  <label>Sort by</label>${opt(1)}<label>Then by</label>${opt(2)}<label>Then by</label>${opt(3)}<div class="dialog-actions"><button class="btn" id="soC">Cancel</button><button class="btn primary" id="soO">Sort</button></div>`);
  $('#soC').onclick=close;$('#soO').onclick=()=>{const keys=[1,2,3].map(i=>({c:+$('#so'+i).value,asc:$('#sd'+i).value==='1'})).filter(k=>k.c>=0);
    sortBlock(s.r1+($('#soH').checked?1:0),s.r2,s.c1,s.c2,keys);close();if(AF())reapply()}}

/* ---------- Insert Function + autocomplete ---------- */
function insFn(n){const a=API.getActive(),s=sel(),m=s.r1!==s.r2||s.c1!==s.c2;API.pushUndo();cellO(a.c,a.r).v=`=${n}(${m?id(s.c1,s.r1)+':'+id(s.c2,s.r2):''})`;API.refresh();close();$('#formulaInput').focus()}
function fnDlg(){let pick=null;
  dlg(`<h3>Insert function</h3><input id="xfS" placeholder="Search for a function…" autocomplete="off"><div class="xfl" id="xfL"></div><div id="xfD" style="font-size:11.5px;color:var(--text-dim);margin-top:8px;min-height:34px"></div><div class="dialog-actions"><button class="btn" id="xfC">Cancel</button><button class="btn primary" id="xfO">OK</button></div>`);
  const draw=()=>{const q=$('#xfS').value.trim().toLowerCase();$('#xfL').innerHTML=META.filter(m=>!q||m[0].toLowerCase().includes(q)||m[2].toLowerCase().includes(q)).map(m=>`<div data-n="${m[0]}" class="${m[0]===pick?'on':''}"><b>${m[0]}</b><span style="opacity:.7;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${esc(m[2])}</span></div>`).join('');
    $$('#xfL div').forEach(d=>{d.onclick=()=>{pick=d.dataset.n;const m=META.find(x=>x[0]===pick);$('#xfD').innerHTML=`<b>${pick}(${esc(m[1])})</b><br>${esc(m[2])}`;draw()};d.ondblclick=()=>insFn(d.dataset.n)})};
  draw();$('#xfS').oninput=draw;$('#xfS').focus();$('#xfC').onclick=close;$('#xfO').onclick=()=>pick?insFn(pick):toast('Pick a function')}
$('.fx-icon').style.cursor='pointer';$('.fx-icon').title='Insert function';$('.fx-icon').onclick=fnDlg;

const ac=document.createElement('div');ac.id='xac';document.body.appendChild(ac);let acList=[],acIdx=0;
const acTarget=()=>{const e=document.activeElement;return e&&(e.id==='cellEditor'||e.id==='formulaInput')?e:null};
const getT=e=>e.id==='cellEditor'?e.textContent:e.value;
function setT(e,t){if(e.id==='cellEditor'){e.textContent=t;const r=document.createRange();r.selectNodeContents(e);r.collapse(false);const s=getSelection();s.removeAllRanges();s.addRange(r)}else e.value=t;e.dispatchEvent(new Event('input',{bubbles:true}))}
const acHide=()=>{ac.style.display='none';acList=[]};
function acDraw(){const e=acTarget();if(!e)return acHide();const t=getT(e);if(t[0]!=='=')return acHide();
  const m=t.match(/([A-Za-z]+)$/);acList=m?Object.keys(F).filter(n=>n.startsWith(m[1].toUpperCase())).sort().slice(0,8):[];if(acIdx>=acList.length)acIdx=0;
  let depth=0,hint='';for(let i=t.length-1;i>=0;i--){const ch=t[i];if(ch===')')depth++;else if(ch==='('){if(!depth){const mm=t.slice(0,i).match(/([A-Za-z]+)$/);if(mm)hint=mm[1].toUpperCase();break}depth--}}
  const sg=hint&&sigOf(hint);if(!acList.length&&!sg)return acHide();
  ac.innerHTML=(sg?`<div class="h">${hint}(<b>${esc(sg)}</b>)</div>`:'')+acList.map((n,i)=>`<div class="i${i===acIdx?' on':''}" data-n="${n}"><b>${n}</b><small>${esc((META.find(x=>x[0]===n)||[])[2]||'')}</small></div>`).join('');
  const r=e.getBoundingClientRect();ac.style.left=r.left+'px';ac.style.top=r.bottom+2+'px';ac.style.display='block';
  $$('#xac .i').forEach(d=>d.onmousedown=ev=>{ev.preventDefault();acPick(d.dataset.n)})}
function acPick(n){const e=acTarget();if(!e)return;setT(e,getT(e).replace(/[A-Za-z]+$/,n+'('));acHide();acDraw()}
document.addEventListener('input',acDraw);document.addEventListener('focusout',()=>setTimeout(()=>{if(!acTarget())acHide()},120));
document.addEventListener('keydown',e=>{if(ac.style.display==='none')return;
  if(e.key==='Enter'||e.key==='Escape'){if(e.key==='Escape'&&acList.length){/* let editor cancel too */}acHide();return}
  if(!acList.length||!acTarget())return;
  if(e.key==='ArrowDown'||e.key==='ArrowUp'){e.preventDefault();e.stopImmediatePropagation();acIdx=(acIdx+(e.key==='ArrowDown'?1:-1)+acList.length)%acList.length;acDraw()}
  else if(e.key==='Tab'){e.preventDefault();e.stopImmediatePropagation();acPick(acList[acIdx])}},true);

/* ---------- find & replace ---------- */
function findDlg(){
  dlg(`<h3>Find &amp; Replace</h3><label>Find what</label><input id="fdF"><label>Replace with</label><input id="fdR"><div class="row"><label><input type="checkbox" id="fdC">Match case</label><label><input type="checkbox" id="fdE">Entire cell</label><label><input type="checkbox" id="fdV">Formulas</label></div><div id="fdS" style="font-size:11px;color:var(--text-faint);margin-top:8px;min-height:14px"></div>
  <div class="dialog-actions" style="flex-wrap:wrap"><button class="btn" id="fdN">Find next</button><button class="btn" id="fdP">Replace</button><button class="btn" id="fdA">Replace all</button><button class="btn" id="fdX">Close</button></div>`);
  $('#fdF').focus();$('#fdX').onclick=close;
  const rx=g=>{const t=$('#fdF').value;if(!t)return null;let p=t.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');if($('#fdE').checked)p='^'+p+'$';return new RegExp(p,($('#fdC').checked?'':'i')+(g?'g':''))};
  const st=m=>$('#fdS').textContent=m;
  const txt=(rows,c,r)=>$('#fdV').checked?raw(c,r):String((rows[r]||[])[c]??'');
  const go=(c,r)=>{const b=$('#cellref');b.value=id(c,r);b.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'}))};
  const next=()=>{const re=rx();if(!re)return false;const rows=snapRows(),a=API.getActive(),R=sh().rows,tot=R*COLS;
    for(let k=1;k<=tot;k++){const i=(a.r*COLS+a.c+k)%tot,r=Math.floor(i/COLS),c=i%COLS,t=txt(rows,c,r);if(t!==''&&re.test(t)){go(c,r);st('Found in '+id(c,r));return true}}st('No matches found');return false};
  $('#fdN').onclick=next;
  $('#fdP').onclick=()=>{const re=rx(true);if(!re)return;const a=API.getActive(),v=raw(a.c,a.r);
    if(v!==''&&new RegExp(re.source,re.flags.replace('g','')).test(v)){API.pushUndo();cellO(a.c,a.r).v=v.replace(re,$('#fdR').value);API.refresh()}next()};
  $('#fdA').onclick=()=>{const re=rx(true);if(!re)return;API.pushUndo();let n=0;
    Object.keys(sh().data).forEach(k=>{const x=sh().data[k];if(typeof x.v==='string'&&x.v&&new RegExp(re.source,re.flags.replace('g','')).test(x.v)){x.v=x.v.replace(re,$('#fdR').value);n++}});
    API.refresh();st(`Replaced in ${n} cell(s)`)}}

/* ---------- sizes, freeze, view ---------- */
function sizeDlg(kind){const s=sel(),col=kind==='col',cur=col?(sh().colWidths[s.c1]||100):(sh().rowHeights[s.r1]||26);
  dlg(`<h3>${col?'Column width':'Row height'}</h3><label>Pixels</label><input id="szV" type="number" min="10" value="${cur}"><div class="dialog-actions"><button class="btn" id="szC">Cancel</button><button class="btn primary" id="szO">OK</button></div>`);
  $('#szC').onclick=close;$('#szO').onclick=()=>{const v=Math.max(10,+$('#szV').value||cur);API.pushUndo();
    if(col)for(let c=s.c1;c<=s.c2;c++)sh().colWidths[c]=v;else for(let r=s.r1;r<=s.r2;r++)sh().rowHeights[r]=v;close();API.refresh()}}
function autoFit(){const s=sel();API.pushUndo();for(let c=s.c1;c<=s.c2;c++){let m=0;$$(`#grid td.cell[data-c="${c}"]`).forEach(td=>{m=Math.max(m,td.textContent.length)});sh().colWidths[c]=Math.min(400,Math.max(60,m*7.5+22))}API.refresh();toast('Column width auto-fitted')}
function freezeAt(){const a=API.getActive();if(!a.r&&!a.c){toast('Select a cell below/right of what to freeze');return}sh().frozenRow=a.r;sh().frozenCol=a.c;API.refresh();toast(`Frozen ${a.r} row(s) and ${a.c} column(s)`)}
function freezeApply(){const s=sh(),fr=s.frozenRow|0,fc=s.frozenCol|0;if(fr<=1&&fc<=1)return;const rows=[...$('#grid tbody').children];
  if(fr>1){let top=26;rows.filter(tr=>+tr.firstChild.dataset.r<fr).forEach(tr=>{[...tr.children].forEach(e=>{e.style.position='sticky';e.style.top=top+'px';e.style.zIndex=e.classList.contains('rowhead')?26:12;if(e.tagName==='TD'&&!e.style.background)e.style.background='var(--surface)'});top+=tr.offsetHeight})}
  if(fc>1)rows.forEach(tr=>{let left=44;[...tr.children].forEach(e=>{if(e.tagName!=='TD'||+e.dataset.c>=fc)return;e.style.position='sticky';e.style.left=left+'px';e.style.zIndex=Math.max(+e.style.zIndex||0,11);if(!e.style.background)e.style.background='var(--surface)';left+=e.offsetWidth})})}
const tog=cls=>document.body.classList.toggle(cls);
let painter=null;
document.addEventListener('mouseup',e=>{if(!painter||!e.target.closest('#grid td.cell'))return;const f=painter;painter=null;setTimeout(()=>applyF(t=>{Object.keys(t).forEach(k=>delete t[k]);Object.assign(t,JSON.parse(JSON.stringify(f)))},'Format painted'),0)});
const paintBtn=()=>{const a=API.getActive(),x=sh().data[id(a.c,a.r)];painter=JSON.parse(JSON.stringify((x&&x.f)||{}));toast('Format Painter on — select the target cells')};

/* ---------- ribbon groups ---------- */
grp('Home','Clipboard+',[mkb('🖌','Format Painter',paintBtn),mkb('🗒','New Note',noteDlg)]);
grp('Home','Cells+',[mkb('↔','Col width',()=>sizeDlg('col')),mkb('↕','Row height',()=>sizeDlg('row')),mkb('⇔','AutoFit',autoFit)]);
grp('Home','Editing',[mkb('🔎','Find & Replace',findDlg),mkb('⇅','Custom Sort',sortDlg),mkb('⏷','Filter',toggleAF)]);
grp('Insert','Comments',[mkb('🗒','Note',noteDlg)]);
grp('Formulas','Insert',[mkb('ƒx','Insert Function',fnDlg)]);
grp('Data','Sort & Filter',[mkb('⇅','Custom Sort',sortDlg),mkb('⏷','Filter',toggleAF)]);
grp('Data','Visual rules',[mkb('🎨','Cond. Formatting',cfDlg),mks('▦','Format as Table',Object.keys(TS),fmtTable)]);
grp('View','Freeze+',[mkb('❄','Freeze at selection',freezeAt),mkb('☀','Unfreeze',()=>{sh().frozenRow=0;sh().frozenCol=0;API.refresh()})]);
grp('View','Show',[mkb('▦','Gridlines',()=>tog('xnogrid')),mkb('①','Headings',()=>tog('xnohead')),mkb('fx','Formula bar',()=>tog('xnofb'))]);

/* ---------- shortcuts ---------- */
document.addEventListener('keydown',e=>{const a=document.activeElement;if(a&&(a.tagName==='INPUT'||a.tagName==='SELECT'||a.tagName==='TEXTAREA'||a.isContentEditable))return;
  if($('#dialogOverlay').classList.contains('show'))return;const c=e.ctrlKey||e.metaKey,k=e.key.toLowerCase(),stop=()=>{e.preventDefault();e.stopImmediatePropagation()};
  if(c&&k==='a'){stop();const b=$('#cellref');b.value='A1:'+L(COLS-1)+sh().rows;b.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter'}))}
  else if(c&&(k==='f'||k==='h')){stop();findDlg()}
  else if((c&&k==='g')||e.key==='F5'){stop();$('#cellref').focus();$('#cellref').select()}
  else if(e.shiftKey&&e.key==='F2'){stop();noteDlg()}
  else if(c&&e.shiftKey&&k==='l'){stop();toggleAF()}},true);

/* ---------- post-render decoration ---------- */
let busy=false,raf=0;
function decorate(){const all=$$('#grid td.cell');if(!all.length||all.every(t=>t.dataset.xd))return;busy=true;
  const s=sh(),af=s.af,notes=s.notes||{},map={};
  all.forEach(td=>{const r=+td.dataset.r,c=+td.dataset.c,k=id(c,r),cell=s.data[k],f=cell&&cell.f;map[r+','+c]=td;
    if(f&&f.xf){const cv=td.querySelector('.cv'),n=parseFloat(cv&&cv.textContent);if(cv&&!isNaN(n)&&/^-?[\d.]+(e[+-]?\d+)?$/i.test(cv.textContent.trim()))cv.textContent=fmtX(n,f.xf)}
    if(notes[k]){td.classList.add('xnote');td.title=notes[k]}
    if(af&&r===af.r&&c>=af.c1&&c<=af.c2)td.insertAdjacentHTML('beforeend',`<span class="xaf${af.f[c]?' on':''}" data-c="${c}">▼</span>`)});
  (s.cfx||[]).forEach(R=>cfRun(R,map));
  freezeApply();all.forEach(t=>t.dataset.xd='1');obs.takeRecords();busy=false}
const obs=new MutationObserver(()=>{if(busy)return;cancelAnimationFrame(raf);raf=requestAnimationFrame(decorate)});
obs.observe($('#grid'),{childList:true,subtree:true});
decorate();
toast('Excel Pro pack loaded');
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',boot);else boot();
})();
