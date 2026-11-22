/* Pro Sheet home (dashboard.html) — Excel desktop "Home" look.
   Drop-in replacement. Keep <link rel="stylesheet" href="dashboard-excel.css"> as the LAST item in <head>. */

:root{
  --bg:#f3f3f3;--surface:#fff;--s2:#f3f3f3;--s3:#e1dfdd;--border:#e1dfdd;
  --text:#201f1e;--dim:#605e5c;--faint:#797775;
  --accent:#217346;--accent2:#185c37;--soft:#e6f2ec;
  --rail:240px;
}
html,body{font-family:'Segoe UI','Segoe UI Web',Calibri,sans-serif}
body{background:#f3f3f3;color:#201f1e;font-size:14px}
button,input,select{font-family:inherit}
.hero h1,.sec-title,.logo,.stat b,.empty h3,.dialog h3{font-family:'Segoe UI Light','Segoe UI',Calibri,sans-serif}

/* ---- loading screen ---- */
#loadingScreen{background:#fff}
.load-mark{background:#217346;color:#fff;border-radius:4px;font-family:Calibri,'Segoe UI',sans-serif;animation:none}
.load-bar{background:#e1dfdd;border-radius:0}.load-bar i{background:#217346}
.load-sub{color:#797775}

/* ---- left green backstage pane ---- */
body::after{content:"Home\A\A New\A\A Open";white-space:pre;position:fixed;left:0;top:0;bottom:0;width:var(--rail);
  background:#217346;color:#fff;z-index:30;padding:96px 0 0 24px;font:400 15px/1.2 'Segoe UI',Calibri,sans-serif;box-sizing:border-box;pointer-events:none}
body::before{content:"Pro Sheet";position:fixed;left:0;top:0;width:var(--rail);height:64px;z-index:50;background:#217346;
  color:#fff;font:600 20px/64px 'Segoe UI',Calibri,sans-serif;padding-left:24px;box-sizing:border-box;pointer-events:none}

/* ---- top bar (light, right of the pane) ---- */
.topnav{margin-left:var(--rail);background:#f3f3f3;border-bottom:1px solid #e1dfdd;backdrop-filter:none;color:#201f1e;padding:10px 32px;height:64px;box-sizing:border-box}
.topnav .logo{display:none}
.search{background:#fff;border:1px solid #c8c6c4;border-radius:2px;max-width:420px;padding:7px 10px}
.search:focus-within{border-color:#217346;box-shadow:0 0 0 1px #217346}
.search svg{stroke:#605e5c}.search input{color:#201f1e}.search input::placeholder{color:#797775}
.btn{background:#fff;border:1px solid #8a8886;color:#201f1e;border-radius:2px;font-weight:400}
.btn:hover{background:#f3f2f1}
.btn.primary{background:#217346;border:1px solid #217346;color:#fff;font-weight:600}
.btn.primary:hover{background:#185c37;filter:none}
.sync{color:#605e5c}.sync.ok{color:#217346;background:#e6f2ec}
.chip{background:#fff;border-color:#c8c6c4;border-radius:2px}

/* ---- main ---- */
.main{margin:0 0 0 var(--rail);max-width:none;padding:28px 40px 90px}
.hero{margin-bottom:22px}
.hero h1{font-weight:300;font-size:32px;letter-spacing:0;margin:0 0 4px}
.hero p{color:#605e5c}
.stats{gap:0}
.stat{background:transparent;border:0;border-left:1px solid #e1dfdd;border-radius:0;padding:4px 20px}
.stat b{color:#217346;font-weight:400;font-size:24px}
.stat span{text-transform:none;letter-spacing:0;font-size:12px;color:#605e5c}

.sec-title{font-weight:300;font-size:22px;color:#201f1e}
.sec-title .count{font:400 12px 'Segoe UI';color:#797775}
.cat{border:0;border-radius:0;padding:6px 12px;font-size:13px;color:#605e5c;background:transparent}
.cat:hover{background:#e1dfdd;color:#201f1e}
.cat.on{background:transparent;color:#217346;font-weight:600;box-shadow:inset 0 -2px 0 #217346}

/* ---- "New" template tiles ---- */
.templates{grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:18px;margin-bottom:40px}
.tpl{background:transparent;border:1px solid transparent;border-radius:0;padding:6px}
.tpl:hover{background:#fff;border-color:#c8c6c4;box-shadow:none;transform:none}
.thumb{height:130px;position:relative;border:1px solid #c8c6c4;
  background:
    linear-gradient(var(--tc),var(--tc)) 0 0/100% 16px no-repeat,
    linear-gradient(#f3f3f3,#f3f3f3) 0 16px/28px 100% no-repeat,
    repeating-linear-gradient(0deg,transparent 0 17px,#e1dfdd 17px 18px) 0 16px/100% 100% no-repeat,
    repeating-linear-gradient(90deg,transparent 0 40px,#e1dfdd 40px 41px) 28px 16px/100% 100% no-repeat,
    #fff}
.tpl:hover .thumb{border-color:#217346}
.thumb .tag{top:24px;right:6px;background:#fff;color:var(--tc);border:1px solid var(--tc);border-radius:2px}
.tpl .lbl{padding:8px 2px 0;font-size:13px;font-weight:400}
.tpl .lbl small{color:#797775;margin-top:1px}
.tpl[data-t="blank"] .thumb{border-color:#217346}

/* ---- Recent files ---- */
.grid{grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:14px}
.doc{background:#fff;border:1px solid #e1dfdd;border-radius:0}
.doc:hover{border-color:#217346;box-shadow:0 2px 6px rgba(0,0,0,.15);transform:none}
.doc .thumb{border:0;border-bottom:1px solid #e1dfdd;height:110px}
.doc-name{font-weight:400}
.grid.list{gap:0;background:#fff;border:1px solid #e1dfdd}
.grid.list .doc{border:0;border-bottom:1px solid #edebe9;box-shadow:none}
.grid.list .doc:hover{background:#e6f2ec}
.grid.list .thumb{width:32px;height:32px;margin:8px 6px 8px 16px;border:0;border-radius:2px;background:#217346;display:grid;place-items:center;flex-shrink:0}
.grid.list .thumb::after{content:'X';color:#fff;font:700 18px Calibri,'Segoe UI',sans-serif}
.grid.list .doc-name{font-weight:400;font-size:14px}
.grid.list .doc-meta{font:400 12px 'Segoe UI';color:#605e5c}

.cloud{background:#fff;color:#217346;border:1px solid #c8c6c4;border-radius:2px}
.mbtn{background:#fff;border:1px solid #c8c6c4;color:#201f1e;border-radius:2px}
.menu{background:#fff;border:1px solid #c8c6c4;border-radius:2px;box-shadow:0 3.2px 7.2px rgba(0,0,0,.25)}
.mi{border-radius:0}.mi:hover{background:#e6f2ec;color:#201f1e}.mi.danger:hover{background:#fde7e9;color:#a4262c}
.seg,select.sort{background:#fff;border:1px solid #8a8886;border-radius:2px;color:#201f1e}
.seg button{color:#605e5c}.seg button.on{background:#e6f2ec;color:#217346}
.empty{background:#fff;border:1px dashed #c8c6c4;border-radius:0;color:#605e5c}
.empty h3{color:#201f1e}

/* ---- dialogs & toast ---- */
.overlay{background:rgba(255,255,255,.6);backdrop-filter:none}
.dialog{background:#fff;border:0;border-radius:2px;box-shadow:0 6.4px 14.4px rgba(0,0,0,.3)}
.dialog input{background:#fff;border:1px solid #8a8886;border-radius:2px;color:#201f1e}
.dialog input:focus{outline:0;border-color:#217346}
.toast{background:#323130;color:#fff;border:0;border-radius:2px}

::-webkit-scrollbar-thumb{background:#c8c6c4}

@media(max-width:800px){
  body::after,body::before{display:none}
  .topnav{margin-left:0;padding:10px 14px}.topnav .logo{display:flex;color:#217346;font-size:0}
  .main{margin-left:0;padding:22px 14px 80px}
}
