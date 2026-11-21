/* Pro Sheet add-on: extra data tools, themes and shortcuts.
   Needs the small ProSheetAPI patch in prosheet.html (see instructions). */
(function(){
"use strict";
function boot(){
  const API = window.ProSheetAPI;
  if(!API || !API.getWorkbook || !document.getElementById('chartBtn')){ return setTimeout(boot, 200); }
  const L = API.colLetter, COLS = 26;
  const sh = ()=>{ const w = API.getWorkbook(); return w.sheets[w.current]; };
  const sel = ()=>{ const s = API.getSelection(); return {r1:Math.min(s.r1,s.r2), r2:Math.max(s.r1,s.r2), c1:Math.min(s.c1,s.c2), c2:Math.max(s.c1,s.c2)}; };
  const id = (c,r)=>L(c)+(r+1);
  const raw = (c,r)=>{ const x = sh().data[id(c,r)]; return x && x.v !== undefined ? String(x.v) : ""; };
  const cell = (c,r)=>{ const d = sh().data; return d[id(c,r)] || (d[id(c,r)] = {v:"", f:{}}); };
  const isF = v=>typeof v==='string' && v.startsWith('=');
  const run = (fn,msg)=>{ API.pushUndo(); fn(); API.refresh(); if(msg) API.toast(msg); };

  /* ---------- tools ---------- */
  const tools = {
    removeDupes(){
      const s = sel(); if(s.r1===s.r2) return API.toast("Select a range with several rows");
      run(()=>{
        const seen = new Set(), keep = [];
        for(let r=s.r1;r<=s.r2;r++){
          const row = []; for(let c=s.c1;c<=s.c2;c++) row.push(JSON.parse(JSON.stringify(sh().data[id(c,r)]||{v:"",f:{}})));
          const key = row.map(x=>x.v).join('\u0001');
          if(!seen.has(key)){ seen.add(key); keep.push(row); }
        }
        const removed = (s.r2-s.r1+1)-keep.length;
        for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++){
          const k = r-s.r1; if(k<keep.length) sh().data[id(c,r)] = keep[k][c-s.c1]; else delete sh().data[id(c,r)];
        }
        API.toast(`Removed ${removed} duplicate row(s)`);
      });
    },
    trim(){ run(()=>{ eachSel((c,r)=>{ const x = sh().data[id(c,r)]; if(x && typeof x.v==='string' && !isF(x.v)) x.v = x.v.replace(/\s+/g,' ').trim(); }); },"Trimmed extra spaces"); },
    upper(){ caseConv(v=>v.toUpperCase(),"UPPERCASE applied"); },
    lower(){ caseConv(v=>v.toLowerCase(),"lowercase applied"); },
    proper(){ caseConv(v=>v.toLowerCase().replace(/(^|\s)\S/g,t=>t.toUpperCase()),"Proper Case applied"); },
    zebra(){
      const s = sel();
      run(()=>{ for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++){ const x = cell(c,r); x.f = x.f||{}; x.f.bg = (r-s.r1)%2 ? 'rgba(108,140,255,0.10)' : 'transparent'; } },"Zebra stripes applied");
    },
    headerStyle(){
      const s = sel();
      run(()=>{ for(let c=s.c1;c<=s.c2;c++){ const x = cell(c,s.r1); x.f = x.f||{}; x.f.bold = true; x.f.bg = '#2b3a6b'; x.f.color = '#ffffff'; x.f.align = 'center'; } },"Header styled");
    },
    highlightDupes(){
      const s = sel(), count = {};
      run(()=>{
        eachSel((c,r)=>{ const v = raw(c,r); if(v!=="") count[v] = (count[v]||0)+1; });
        let n = 0;
        eachSel((c,r)=>{ const v = raw(c,r), x = cell(c,r); x.f = x.f||{};
          if(v!=="" && count[v]>1){ x.f.bg = 'rgba(255,107,107,0.35)'; n++; } else if(x.f.bg==='rgba(255,107,107,0.35)') x.f.bg = 'transparent'; });
        API.toast(n ? `${n} duplicate cell(s) highlighted` : "No duplicates found");
      });
    },
    transpose(){
      const s = sel(), h = s.r2-s.r1+1, w = s.c2-s.c1+1, top = s.r2+2;
      if(top + w >= sh().rows) return API.toast("Not enough rows below the selection");
      if(s.c1 + h > COLS) return API.toast("Result would exceed 26 columns");
      run(()=>{ for(let r=0;r<h;r++) for(let c=0;c<w;c++){
        const src = sh().data[id(s.c1+c, s.r1+r)]; if(!src) continue;
        sh().data[id(s.c1+r, top+c)] = JSON.parse(JSON.stringify(src)); } },"Transposed below the selection");
    },
    autofit(){
      run(()=>{
        const widths = {};
        Object.keys(sh().data).forEach(k=>{ const m = k.match(/^([A-Z]+)(\d+)$/); if(!m) return;
          let c = 0; for(const ch of m[1]) c = c*26 + ch.charCodeAt(0)-64; c--;
          const x = sh().data[k]; if(!x || x.v==="" ) return;
          const len = isF(x.v) ? 10 : String(x.v).length;
          widths[c] = Math.max(widths[c]||0, len); });
        Object.keys(widths).forEach(c=>{ sh().colWidths[c] = Math.min(320, Math.max(60, widths[c]*8 + 24)); });
      },"Columns auto-fitted");
    },
    totalsRow(){
      const s = sel(), tr = s.r2+1;
      if(tr >= sh().rows) return API.toast("No room below the selection");
      run(()=>{ for(let c=s.c1;c<=s.c2;c++){
        let numeric = false; for(let r=s.r1;r<=s.r2;r++){ const v = raw(c,r); if(v!=="" && !isNaN(parseFloat(v)) && !isF(v)) numeric = true; }
        if(numeric){ const x = cell(c,tr); x.v = `=SUM(${id(c,s.r1)}:${id(c,s.r2)})`; x.f = Object.assign({}, x.f, {bold:true}); }
      } cell(s.c1, tr).v === "" && (cell(s.c1,tr).v = "Total", cell(s.c1,tr).f = {bold:true}); },"Totals row added");
    },
    fillDown(){
      const s = sel();
      if(s.r1===s.r2) return API.toast("Select 2+ rows to fill down");
      run(()=>{ for(let c=s.c1;c<=s.c2;c++){ const src = sh().data[id(c,s.r1)]; if(!src) continue;
        for(let r=s.r1+1;r<=s.r2;r++) sh().data[id(c,r)] = JSON.parse(JSON.stringify(src)); } },"Filled down");
    },
    clearFormats(){ run(()=>{ eachSel((c,r)=>{ const x = sh().data[id(c,r)]; if(x) x.f = {}; }); },"Formatting cleared"); },
    textToNumber(){
      run(()=>{ let n=0; eachSel((c,r)=>{ const x = sh().data[id(c,r)]; if(!x||typeof x.v!=='string'||isF(x.v)) return;
        const t = x.v.replace(/[₹$,\s]/g,''); if(t!=="" && !isNaN(Number(t)) && t!==x.v){ x.v = t; n++; } });
        API.toast(`Converted ${n} cell(s) to numbers`); });
    },
    insertDate(){
      run(()=>{ cell(API.getActive().c, API.getActive().r).v = new Date().toISOString().slice(0,10); },"Date inserted");
    },
    summary(){
      const s = sel(), nums = [];
      eachSel((c,r)=>{ const n = parseFloat(raw(c,r)); if(!isNaN(n) && !isF(raw(c,r))) nums.push(n); });
      if(!nums.length) return API.toast("No numbers in selection");
      nums.sort((a,b)=>a-b);
      const sum = nums.reduce((a,b)=>a+b,0), avg = sum/nums.length, med = nums.length%2 ? nums[(nums.length-1)/2] : (nums[nums.length/2-1]+nums[nums.length/2])/2;
      const f = n=>Math.round(n*100)/100;
      API.toast(`n=${nums.length} · Σ ${f(sum)} · avg ${f(avg)} · median ${f(med)} · min ${nums[0]} · max ${nums[nums.length-1]}`);
    }
  };
  function eachSel(cb){ const s = sel(); for(let r=s.r1;r<=s.r2;r++) for(let c=s.c1;c<=s.c2;c++) cb(c,r); }
  function caseConv(fn,msg){ run(()=>{ eachSel((c,r)=>{ const x = sh().data[id(c,r)]; if(x && typeof x.v==='string' && !isF(x.v)) x.v = fn(x.v); }); },msg); }

  /* ---------- UI: Add-ons menu ---------- */
  const css = document.createElement('style');
  css.textContent = `
  #addonBtn{position:relative}
  #addonMenu{position:absolute;top:calc(100% + 6px);right:0;width:250px;background:var(--surface-2);border:1px solid var(--border);border-radius:10px;padding:6px;box-shadow:0 14px 36px rgba(0,0,0,.55);display:none;z-index:300;max-height:70vh;overflow:auto;text-align:left}
  #addonMenu.show{display:block}
  #addonMenu .ah{font:600 9.5px 'DM Mono',monospace;color:var(--text-faint);text-transform:uppercase;letter-spacing:.6px;padding:8px 10px 4px}
  #addonMenu .ai{padding:7px 10px;font-size:12.5px;border-radius:6px;cursor:pointer;display:flex;justify-content:space-between;gap:10px}
  #addonMenu .ai:hover{background:var(--accent-soft);color:var(--accent)}
  #addonMenu .ai kbd{font:10px 'DM Mono',monospace;color:var(--text-faint)}
  html.ps-light{--bg:#f4f6fb;--surface:#ffffff;--surface-2:#f0f2f8;--surface-3:#e4e8f2;--border:#d5dae8;--border-soft:#e6e9f2;--text:#1b2030;--text-dim:#5b6278;--text-faint:#8a90a5;--header-bg:#fafbfe;--grid-line:#e3e7f0;--accent-soft:rgba(108,140,255,.16)}
  html.ps-light td.cell.active-cell{background:#fff}
  html.ps-light .colhead.sel,html.ps-light .rowhead.sel{color:#fff}`;
  document.head.appendChild(css);

  const groups = [
    ["Clean up", [["Remove duplicate rows","removeDupes"],["Trim extra spaces","trim"],["Text → numbers (strip ₹ , )","textToNumber"],["UPPERCASE","upper"],["lowercase","lower"],["Proper Case","proper"]]],
    ["Analyze", [["Highlight duplicates","highlightDupes"],["Quick summary (Σ avg median)","summary"],["Add totals row","totalsRow"]]],
    ["Layout & style", [["Style header row","headerStyle"],["Zebra stripes","zebra"],["Auto-fit columns","autofit"],["Transpose below","transpose"],["Clear formatting","clearFormats"]]],
    ["Quick insert", [["Fill down","fillDown","Ctrl+D"],["Insert today's date","insertDate","Ctrl+;"]]]
  ];
  const wrap = document.createElement('div');
  wrap.style.position = 'relative';
  wrap.innerHTML = `<button class="btn" id="addonBtn" title="Extra tools">Add-ons ▾</button><div id="addonMenu"></div>`;
  const menu = wrap.querySelector('#addonMenu');
  groups.forEach(([title,items])=>{
    menu.insertAdjacentHTML('beforeend', `<div class="ah">${title}</div>`);
    items.forEach(([label,key,kbd])=>{
      const el = document.createElement('div'); el.className = 'ai';
      el.innerHTML = `<span>${label}</span>${kbd?`<kbd>${kbd}</kbd>`:''}`;
      el.onclick = (e)=>{ e.stopPropagation(); menu.classList.remove('show'); tools[key](); };
      menu.appendChild(el);
    });
  });
  menu.insertAdjacentHTML('beforeend', `<div class="ah">View</div>`);
  const themeItem = document.createElement('div'); themeItem.className = 'ai';
  const setTheme = light=>{ document.documentElement.classList.toggle('ps-light', light); themeItem.innerHTML = `<span>${light?'Switch to dark theme':'Switch to light theme'}</span>`; try{ localStorage.setItem('prosheet_theme', light?'light':'dark'); }catch(e){} };
  themeItem.onclick = (e)=>{ e.stopPropagation(); menu.classList.remove('show'); setTheme(!document.documentElement.classList.contains('ps-light')); };
  menu.appendChild(themeItem);
  let saved = null; try{ saved = localStorage.getItem('prosheet_theme'); }catch(e){}
  setTheme(saved === 'light');

  const chartBtn = document.getElementById('chartBtn');
  chartBtn.parentNode.insertBefore(wrap, chartBtn);
  wrap.querySelector('#addonBtn').onclick = (e)=>{ e.stopPropagation(); menu.classList.toggle('show'); };
  document.addEventListener('click', ()=>menu.classList.remove('show'));

  /* ---------- shortcuts ---------- */
  document.addEventListener('keydown', (e)=>{
    const a = document.activeElement;
    if(a && (a.tagName==='INPUT' || a.tagName==='SELECT' || a.isContentEditable)) return;
    const ctrl = e.ctrlKey || e.metaKey;
    if(ctrl && e.key.toLowerCase()==='d'){ e.preventDefault(); tools.fillDown(); }
    else if(ctrl && e.key===';'){ e.preventDefault(); tools.insertDate(); }
  });
}
if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot); else boot();
})();
