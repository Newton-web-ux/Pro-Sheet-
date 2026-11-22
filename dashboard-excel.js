/* Pro Sheet home screen (dashboard.html) — Excel look.
   Add as the LAST thing in <head>, after the inline <style>:
   <link rel="stylesheet" href="dashboard-excel.css"> */

:root{
  --bg:#f3f3f3;--surface:#ffffff;--s2:#f3f3f3;--s3:#e1e1e1;--border:#d4d4d4;
  --text:#222;--dim:#555;--faint:#777;
  --accent:#217346;--accent2:#1a5c38;--soft:rgba(33,115,70,.12);
}
body{background:#f3f3f3;color:#222;font-family:'Segoe UI',Calibri,'DM Sans',sans-serif}
.hero h1,.sec-title,.logo,.stat b,.empty h3,.dialog h3{font-family:'Segoe UI',Calibri,'DM Sans',sans-serif}
.hero h1{font-weight:600;font-size:26px}
.sec-title{font-weight:600;font-size:15px}

/* loading screen */
.load-mark{background:#217346;color:#fff;font-family:Calibri,'Segoe UI',sans-serif}
.load-bar i{background:#217346}

/* green title bar like Excel */
.topnav{background:#217346;border-bottom:0;backdrop-filter:none;color:#fff}
.logo{color:#fff}
.mark{background:#fff;color:#217346;border-radius:4px;font-family:Calibri,'Segoe UI',sans-serif}
.search{background:rgba(255,255,255,.18);border-color:transparent;border-radius:3px}
.search input{color:#fff}
.search input::placeholder{color:rgba(255,255,255,.78)}
.search svg{stroke:#fff}
.search:focus-within{background:#fff;box-shadow:none;border-color:#fff}
.search:focus-within input{color:#222}
.search:focus-within svg{stroke:#555}
.topnav .btn{background:rgba(255,255,255,.16);border-color:transparent;color:#fff;border-radius:3px}
.topnav .btn:hover{background:rgba(255,255,255,.28)}
.topnav .btn.primary{background:#fff;color:#217346;font-weight:600}
.topnav .btn.primary:hover{background:#eef7f1;filter:none}
.topnav .chip{background:rgba(255,255,255,.16);border-color:transparent;color:#fff}
.topnav .sync{color:rgba(255,255,255,.85)}
.topnav .sync.ok{background:rgba(255,255,255,.18);color:#fff}
.btn{border-radius:3px}
.btn.primary{background:#217346;color:#fff;border:0}
.btn.primary:hover{background:#1a5c38;filter:none}

/* stats */
.stat{border-radius:2px;background:#fff}
.stat b{color:#217346;font-weight:600}
.stat span{text-transform:none;letter-spacing:0}

/* section chips become flat tabs */
.cat{border-radius:2px;border-color:transparent;background:transparent}
.cat:hover{background:#e1e1e1}
.cat.on{background:#fff;border-color:#d4d4d4;color:#217346;font-weight:600;box-shadow:inset 0 -2px 0 #217346}

/* template tiles + document previews look like little worksheets */
.thumb{
  background:
    linear-gradient(var(--tc),var(--tc)) 0 0/100% 14px no-repeat,
    repeating-linear-gradient(0deg,transparent 0 17px,#dcdcdc 17px 18px) 0 14px/100% 100% no-repeat,
    repeating-linear-gradient(90deg,transparent 0 44px,#dcdcdc 44px 45px) 0 14px/100% 100% no-repeat,
    #fff;
  border-bottom:1px solid #d4d4d4}
.thumb .tag{background:#fff;color:var(--tc);border:1px solid var(--tc);top:22px}
.tpl,.doc{border-radius:2px;background:#fff}
.tpl:hover,.doc:hover{box-shadow:0 2px 8px rgba(0,0,0,.18);transform:none;border-color:#217346}
.tpl .lbl small{color:#777}

/* recent files: Excel "Recent" list (use the list view toggle) */
.grid.list{gap:0;background:#fff;border:1px solid #d4d4d4}
.grid.list .doc{border:0;border-bottom:1px solid #e6e6e6;border-radius:0}
.grid.list .doc:hover{background:#eef7f1;box-shadow:none}
.grid.list .thumb{width:34px;height:34px;margin:8px 4px 8px 14px;border:0;border-radius:2px;background:#217346;display:grid;place-items:center}
.grid.list .thumb::after{content:'X';color:#fff;font:700 19px Calibri,'Segoe UI',sans-serif}
.grid.list .doc-name{font-weight:500}

/* small controls */
.mbtn{background:#fff;color:#444;border:1px solid #ccc}
.cloud{background:#fff;color:#217346;border:1px solid #d4d4d4}
.menu{background:#fff;border-radius:2px;box-shadow:0 4px 14px rgba(0,0,0,.22)}
.mi:hover{background:#e8f3ec;color:#217346}
.seg,select.sort{border-radius:2px;background:#fff}
.seg button.on{background:#e8f3ec;color:#217346}
.empty{background:#fff;border-radius:2px}
.dialog{background:#fff;border-radius:4px}
.dialog input{border-radius:2px}
.toast{background:#323232;color:#fff;border-color:#323232;border-radius:3px}
