/*
 * Pro Sheet Site Builder
 * An Apps Script style editor (Code.gs + HTML files) that deploys the sheet as a public site.
 * Deployed sites live at site.html?d=<id>. Visitor writes are queued in Firestore and applied
 * to the sheet while it is open in Pro Sheet.
 */
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js';
import {
  getAuth, GoogleAuthProvider, signInWithCredential, signInWithPopup, onAuthStateChanged
} from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js';
import {
  getFirestore, doc, setDoc, updateDoc, deleteDoc, getDocs, collection, onSnapshot,
  query, orderBy, writeBatch, serverTimestamp
} from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js';
import { firebaseConfig } from './firebase-config.js';
import { createRuntime } from './site-runtime.js';

const API = window.ProSheetAPI;
if (!API) throw new Error('ProSheetAPI is missing — apply the Pro Sheet patch first.');

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

/* ================= STARTER FILES ================= */
const STARTER_GS = [
  '/**',
  ' * Code.gs: the server side of your site.',
  ' * Supported: SpreadsheetApp, HtmlService, ContentService, Utilities, Logger, Session, LockService.',
  ' * Functions ending with "_" are private and cannot be called from the page.',
  ' */',
  '',
  'function doGet(e) {',
  "  return HtmlService.createTemplateFromFile('index')",
  '    .evaluate()',
  "    .setTitle('Pro Sheet Site');",
  '}',
  '',
  'function getRows() {',
  '  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();',
  '  var last = sheet.getLastRow();',
  '  if (!last) return [];',
  '  return sheet.getRange(1, 1, last, Math.max(1, sheet.getLastColumn())).getValues();',
  '}',
  '',
  'function addRow(name, email) {',
  "  name = String(name || '').trim().slice(0, 100);",
  "  email = String(email || '').trim().slice(0, 150);",
  "  if (!name || !email) throw new Error('Name and email are required.');",
  '  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()',
  '    .appendRow([name, email, new Date().toISOString()]);',
  '  return getRows();',
  '}',
  ''
].join('\n');

const STARTER_HTML = [
  '<!DOCTYPE html>',
  '<html>',
  '<head>',
  '  <meta charset="UTF-8">',
  '  <meta name="viewport" content="width=device-width, initial-scale=1">',
  '  <style>',
  '    :root { color-scheme: light dark; --accent: #4f6bff; }',
  '    * { box-sizing: border-box; }',
  '    body { margin: 0; font: 14px/1.5 system-ui, sans-serif; background: Canvas; color: CanvasText; }',
  '    main { max-width: 720px; margin: 0 auto; padding: 32px 20px; }',
  '    h1 { font-size: 22px; margin: 0 0 20px; }',
  '    form { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 24px; }',
  '    input { flex: 1 1 160px; padding: 9px 11px; border: 1px solid #8884; border-radius: 8px; font: inherit; background: transparent; color: inherit; }',
  '    button { padding: 9px 16px; border: 0; border-radius: 8px; background: var(--accent); color: #fff; font: inherit; font-weight: 600; cursor: pointer; }',
  '    table { width: 100%; border-collapse: collapse; }',
  '    td { padding: 9px 10px; border-bottom: 1px solid #8883; }',
  '    .empty { opacity: .6; }',
  '  </style>',
  '</head>',
  '<body>',
  '<main>',
  '  <h1>Contact list</h1>',
  '  <form id="f">',
  '    <input name="name" placeholder="Name" required>',
  '    <input name="email" type="email" placeholder="Email" required>',
  '    <button>Add</button>',
  '  </form>',
  '  <table id="t"></table>',
  '</main>',
  '<script>',
  "  var form = document.getElementById('f'), table = document.getElementById('t');",
  '',
  '  function render(rows) {',
  "    table.textContent = '';",
  '    if (!rows.length) {',
  "      table.innerHTML = '<tr><td class=\"empty\">No entries yet.</td></tr>';",
  '      return;',
  '    }',
  '    rows.forEach(function (row) {',
  "      var tr = document.createElement('tr');",
  '      row.slice(0, 3).forEach(function (cell) {',
  "        var td = document.createElement('td');",
  '        td.textContent = cell;',
  '        tr.appendChild(td);',
  '      });',
  '      table.appendChild(tr);',
  '    });',
  '  }',
  '',
  '  function load() { google.script.run.withSuccessHandler(render).getRows(); }',
  '',
  "  form.addEventListener('submit', function (e) {",
  '    e.preventDefault();',
  '    var d = new FormData(form);',
  '    google.script.run',
  '      .withSuccessHandler(function (rows) { render(rows); form.reset(); })',
  '      .withFailureHandler(function (err) { alert(err.message); })',
  "      .addRow(d.get('name'), d.get('email'));",
  '  });',
  '',
  '  load();',
  '  ProSheet.onChange(load); // refresh when the sheet changes',
  '</' + 'script>',
  '</body>',
  '</html>',
  ''
].join('\n');

/* ================= STATE ================= */
let files = [];
let activeName = '';
let previewRt = null;
let previewTimer = null;
let pubTimer = null;
let unsubOps = null;
let opsChain = Promise.resolve();
let saveTimer = null;
let busy = false;

function loadFiles() {
  const site = API.getSite();
  files = site && Array.isArray(site.files) && site.files.length
    ? site.files.map((f) => ({ name: f.name, type: f.type, content: f.content }))
    : [
        { name: 'Code.gs', type: 'gs', content: STARTER_GS },
        { name: 'index.html', type: 'html', content: STARTER_HTML }
      ];
  if (!files.some((f) => f.name === activeName)) activeName = files[0].name;
}
function persistFiles() {
  const site = Object.assign({}, API.getSite() || {});
  site.files = files.map((f) => ({ name: f.name, type: f.type, content: f.content }));
  API.setSite(site);
}
function currentFile() { return files.find((f) => f.name === activeName); }
function esc(s) { return String(s).replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])); }
function newId() {
  const b = crypto.getRandomValues(new Uint8Array(10));
  return 's' + Array.from(b, (x) => x.toString(36).padStart(2, '0')).join('');
}
function siteUrl(id) { return new URL('site.html?d=' + encodeURIComponent(id), location.href).href; }

/* ================= UI ================= */
const css = document.createElement('style');
css.textContent = `
.pss{position:fixed;inset:0;z-index:450;background:var(--bg);display:none;flex-direction:column;user-select:text;font-family:'DM Sans',sans-serif;color:var(--text)}
.pss.show{display:flex}
.pss *{user-select:text}
.pss button,.pss select{font-family:inherit}
.pss-top{display:flex;align-items:center;gap:10px;padding:8px 14px;background:var(--header-bg);border-bottom:1px solid var(--border);flex-shrink:0}
.pss-title{font-family:'Syne',sans-serif;font-weight:800;font-size:15px;margin-right:6px}
.pss-sp{flex:1}
.pss-sel{background:var(--surface-2);border:1px solid var(--border);color:var(--text);font-size:12px;padding:6px 8px;border-radius:7px;min-width:130px}
.pss-body{flex:1;display:flex;min-height:0}
.pss-files{width:190px;background:var(--surface);border-right:1px solid var(--border);display:flex;flex-direction:column;flex-shrink:0}
.pss-files h4{margin:0;padding:12px 14px 6px;font-size:10.5px;letter-spacing:.6px;text-transform:uppercase;color:var(--text-faint);font-weight:700}
.pss-file{display:flex;align-items:center;gap:8px;padding:7px 14px;font-size:12.5px;cursor:pointer;color:var(--text-dim)}
.pss-file:hover{background:var(--surface-2);color:var(--text)}
.pss-file.active{background:var(--accent-soft);color:var(--accent)}
.pss-file .badge{font-family:'DM Mono',monospace;font-size:9.5px;padding:1px 5px;border-radius:4px;background:var(--surface-3);color:var(--text-dim)}
.pss-file .nm{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.pss-file .x{opacity:0;color:var(--text-faint);padding:0 3px}
.pss-file:hover .x{opacity:1}
.pss-file .x:hover{color:var(--danger)}
.pss-addrow{display:flex;gap:6px;padding:10px 12px;margin-top:auto;border-top:1px solid var(--border)}
.pss-addrow .btn{flex:1;justify-content:center;font-size:11.5px;padding:6px 8px}
.pss-editor{flex:1;display:flex;min-width:0;background:var(--bg)}
.pss-gutter{width:46px;padding:12px 8px 12px 0;text-align:right;overflow:hidden;font-family:'DM Mono',monospace;font-size:12.5px;line-height:20px;color:var(--text-faint);background:var(--surface);border-right:1px solid var(--border-soft);white-space:pre;flex-shrink:0}
.pss-code{flex:1;min-width:0;border:0;outline:0;resize:none;background:transparent;color:var(--text);font-family:'DM Mono',monospace;font-size:12.5px;line-height:20px;padding:12px 14px;white-space:pre;overflow:auto;tab-size:2}
.pss-side{width:460px;border-left:1px solid var(--border);display:flex;flex-direction:column;background:var(--surface);flex-shrink:0}
.pss-tabs{display:flex;border-bottom:1px solid var(--border);flex-shrink:0}
.pss-tab{padding:10px 16px;font-size:12px;font-weight:500;color:var(--text-dim);cursor:pointer;border-bottom:2px solid transparent}
.pss-tab:hover{color:var(--text)}
.pss-tab.active{color:var(--accent);border-bottom-color:var(--accent)}
.pss-pane{flex:1;min-height:0;display:none;flex-direction:column}
.pss-pane.show{display:flex}
.pss-console{flex:1;overflow:auto;padding:12px 14px;font-family:'DM Mono',monospace;font-size:12px;line-height:1.6}
.pss-console .dim{color:var(--text-faint)}
.pss-console .err{color:var(--danger)}
.pss-console .ok{color:var(--success)}
.pss-note{padding:9px 14px;font-size:11.5px;color:var(--text-dim);background:var(--surface-2);border-bottom:1px solid var(--border);line-height:1.5}
.pss-previewbox{flex:1;background:#fff;position:relative;min-height:0}
.pss-deploy{padding:16px;overflow:auto;display:flex;flex-direction:column;gap:14px;font-size:12.5px}
.pss-deploy h5{margin:0 0 6px;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:var(--text-faint)}
.pss-url{display:flex;gap:6px}
.pss-url input{flex:1;min-width:0;background:var(--bg);border:1px solid var(--border);color:var(--text);border-radius:7px;padding:8px 9px;font-family:'DM Mono',monospace;font-size:11.5px}
.pss-checks{display:flex;flex-direction:column;gap:6px}
.pss-checks label{display:flex;align-items:center;gap:8px;color:var(--text);cursor:pointer}
.pss-hint{color:var(--text-dim);line-height:1.55;font-size:12px}
.pss-pill{display:inline-flex;align-items:center;gap:6px;padding:3px 9px;border-radius:20px;font-size:11px;font-weight:600;background:var(--surface-3);color:var(--text-dim)}
.pss-pill.live{background:rgba(74,222,128,.14);color:var(--success)}
.pss-pill i{width:7px;height:7px;border-radius:50%;background:currentColor;display:inline-block}
`;
document.head.appendChild(css);

const root = document.createElement('div');
root.className = 'pss';
root.innerHTML = `
  <div class="pss-top">
    <div class="pss-title">Site Builder</div>
    <select class="pss-sel" id="pssFn" title="Function to run"></select>
    <button class="btn" id="pssRun">Run</button>
    <div class="pss-sp"></div>
    <button class="btn" id="pssSave">Save</button>
    <button class="btn primary" id="pssDeploy">Deploy</button>
    <button class="btn" id="pssClose" title="Close">Close</button>
  </div>
  <div class="pss-body">
    <aside class="pss-files">
      <h4>Files</h4>
      <div id="pssFileList"></div>
      <div class="pss-addrow">
        <button class="btn" id="pssAddGs">+ Script</button>
        <button class="btn" id="pssAddHtml">+ HTML</button>
      </div>
    </aside>
    <section class="pss-editor">
      <div class="pss-gutter" id="pssGutter"></div>
      <textarea class="pss-code" id="pssCode" spellcheck="false" wrap="off"></textarea>
    </section>
    <aside class="pss-side">
      <div class="pss-tabs">
        <div class="pss-tab active" data-tab="console">Console</div>
        <div class="pss-tab" data-tab="preview">Preview</div>
        <div class="pss-tab" data-tab="deploy">Deployment</div>
      </div>
      <div class="pss-pane show" id="pssPaneConsole"><div class="pss-console" id="pssConsole"></div></div>
      <div class="pss-pane" id="pssPanePreview">
        <div class="pss-note">The preview runs against your live sheet. Anything the site writes is written to this sheet.</div>
        <div class="pss-previewbox" id="pssPreviewBox"></div>
      </div>
      <div class="pss-pane" id="pssPaneDeploy"><div class="pss-deploy" id="pssDeployBody"></div></div>
    </aside>
  </div>`;
document.body.appendChild(root);

// Pro Sheet binds global keyboard and paste handlers on the document. Keep every editing event
// inside the builder so typing here never edits cells.
['keydown', 'keyup', 'keypress', 'paste', 'copy', 'cut'].forEach((ev) =>
  root.addEventListener(ev, (e) => e.stopPropagation()));

const $ = (id) => document.getElementById(id);
const el = {
  fn: $('pssFn'), code: $('pssCode'), gutter: $('pssGutter'), list: $('pssFileList'),
  console: $('pssConsole'), preview: $('pssPreviewBox'), deploy: $('pssDeployBody')
};

function toast(msg) { API.toast(msg); }

/* ---- files ---- */
function renderFiles() {
  el.list.innerHTML = '';
  files.forEach((f) => {
    const row = document.createElement('div');
    row.className = 'pss-file' + (f.name === activeName ? ' active' : '');
    row.innerHTML = `<span class="badge">${f.type === 'gs' ? 'GS' : 'HTML'}</span><span class="nm">${esc(f.name)}</span><span class="x" title="Delete file">✕</span>`;
    row.onclick = () => { activeName = f.name; renderFiles(); loadEditor(); };
    row.ondblclick = () => renameFile(f);
    row.querySelector('.x').onclick = (e) => { e.stopPropagation(); deleteFile(f); };
    el.list.appendChild(row);
  });
}
function addFile(type) {
  const ext = type === 'gs' ? '.gs' : '.html';
  let name = (prompt('File name', type === 'gs' ? 'Utils' : 'style') || '').trim();
  if (!name) return;
  if (!name.toLowerCase().endsWith(ext)) name += ext;
  if (files.some((f) => f.name === name)) { toast('A file with that name already exists'); return; }
  files.push({ name, type, content: type === 'gs' ? '// ' + name + '\n' : '<!-- ' + name + ' -->\n' });
  activeName = name;
  persistFiles(); renderFiles(); loadEditor(); refreshFnList();
}
function renameFile(f) {
  const ext = f.type === 'gs' ? '.gs' : '.html';
  let name = (prompt('Rename file', f.name) || '').trim();
  if (!name || name === f.name) return;
  if (!name.toLowerCase().endsWith(ext)) name += ext;
  if (files.some((x) => x.name === name)) { toast('A file with that name already exists'); return; }
  if (activeName === f.name) activeName = name;
  f.name = name;
  persistFiles(); renderFiles();
}
function deleteFile(f) {
  if (files.length <= 1) { toast('At least one file is required'); return; }
  if (!confirm('Delete ' + f.name + '?')) return;
  files = files.filter((x) => x !== f);
  if (activeName === f.name) activeName = files[0].name;
  persistFiles(); renderFiles(); loadEditor(); refreshFnList();
}

/* ---- editor ---- */
function updateGutter() {
  const n = el.code.value.split('\n').length;
  let s = '';
  for (let i = 1; i <= n; i++) s += i + '\n';
  el.gutter.textContent = s;
  el.gutter.scrollTop = el.code.scrollTop;
}
function loadEditor() {
  const f = currentFile();
  el.code.value = f ? f.content : '';
  updateGutter();
}
el.code.addEventListener('scroll', () => { el.gutter.scrollTop = el.code.scrollTop; });
el.code.addEventListener('input', () => {
  const f = currentFile();
  if (!f) return;
  f.content = el.code.value;
  updateGutter();
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => { persistFiles(); refreshFnList(); }, 600);
});
el.code.addEventListener('keydown', (e) => {
  const mod = e.ctrlKey || e.metaKey;
  if (e.key === 'Tab') {
    e.preventDefault();
    const s = el.code.selectionStart, t = el.code.selectionEnd;
    el.code.setRangeText('  ', s, t, 'end');
    el.code.dispatchEvent(new Event('input'));
  } else if (mod && e.key.toLowerCase() === 's') {
    e.preventDefault(); saveNow();
  } else if (mod && e.key === 'Enter') {
    e.preventDefault(); runSelected();
  }
});
function saveNow() {
  clearTimeout(saveTimer);
  persistFiles();
  toast('Saved');
}

/* ---- function list ---- */
function listFunctions() {
  const names = [];
  files.filter((f) => f.type === 'gs').forEach((f) => {
    const re = /^\s*(?:async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/gm;
    let m;
    while ((m = re.exec(f.content))) if (!m[1].endsWith('_') && !names.includes(m[1])) names.push(m[1]);
  });
  return names;
}
function refreshFnList() {
  const prev = el.fn.value;
  const names = listFunctions();
  el.fn.innerHTML = names.map((n) => `<option>${esc(n)}</option>`).join('') || '<option value="">No functions</option>';
  if (names.includes(prev)) el.fn.value = prev;
}

/* ---- console / run ---- */
function log(text, cls) {
  const d = document.createElement('div');
  if (cls) d.className = cls;
  d.textContent = text;
  el.console.appendChild(d);
  el.console.scrollTop = el.console.scrollHeight;
}
function preview(v) {
  let s;
  try { s = typeof v === 'string' ? v : JSON.stringify(v); } catch (e) { s = String(v); }
  if (s === undefined) s = 'undefined';
  return s.length > 400 ? s.slice(0, 400) + '…' : s;
}
async function runSelected() {
  const fn = el.fn.value;
  if (!fn) { toast('No function to run'); return; }
  showTab('console');
  persistFiles();
  log('▶ ' + fn + '()', 'dim');
  const rt = createRuntime({
    files,
    snapshot: API.snapshot(),
    onOps: (ops) => API.applyOps(ops),
    onLog: (lines) => lines.forEach((l) => log(l))
  });
  try {
    (await rt.ready).forEach((e) => log(e, 'err'));
    const r = await rt.call(fn, []);
    if (r.ok) log('✓ ' + (r.result === null || r.result === undefined ? 'done' : 'returned ' + preview(r.result)), 'ok');
    else log('✗ ' + r.error, 'err');
  } finally {
    rt.destroy();
  }
}

/* ---- preview ---- */
function stopPreview() { if (previewRt) { previewRt.destroy(); previewRt = null; } el.preview.replaceChildren(); }
function startPreview() {
  stopPreview();
  persistFiles();
  previewRt = createRuntime({
    files,
    snapshot: API.snapshot(),
    mount: el.preview,
    onOps: async (ops) => {
      API.applyOps(ops);
      if (previewRt) previewRt.setSnapshot(API.snapshot(), []);
    },
    onLog: (lines) => lines.forEach((l) => log(l))
  });
  previewRt.ready.then((errs) => errs.forEach((e) => log(e, 'err')));
  previewRt.render();
}
function schedulePreviewRefresh() {
  if (!previewRt) return;
  clearTimeout(previewTimer);
  previewTimer = setTimeout(() => { if (previewRt) previewRt.setSnapshot(API.snapshot(), []); }, 300);
}

/* ---- tabs ---- */
let activeTab = 'console';
function showTab(name) {
  activeTab = name;
  root.querySelectorAll('.pss-tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === name));
  $('pssPaneConsole').classList.toggle('show', name === 'console');
  $('pssPanePreview').classList.toggle('show', name === 'preview');
  $('pssPaneDeploy').classList.toggle('show', name === 'deploy');
  if (name === 'preview') startPreview();
  else stopPreview();
  if (name === 'deploy') renderDeploy();
}
root.querySelectorAll('.pss-tab').forEach((t) => { t.onclick = () => showTab(t.dataset.tab); });

/* ================= FIREBASE: AUTH ================= */
async function ensureAuth() {
  await auth.authStateReady();
  if (auth.currentUser) return auth.currentUser;
  try {
    const saved = JSON.parse(localStorage.getItem('prosheet_auth') || 'null');
    if (saved && saved.idToken) {
      const cred = await signInWithCredential(auth, GoogleAuthProvider.credential(saved.idToken));
      return cred.user;
    }
  } catch (e) { /* expired token or client id not allowed in Firebase: fall back to the popup */ }
  const cred = await signInWithPopup(auth, new GoogleAuthProvider());
  return cred.user;
}

/* ================= FIREBASE: DEPLOY ================= */
function siteState() {
  return Object.assign({ writeMode: 'append', sheets: null, applied: [] }, API.getSite() || {});
}
function publishedSnapshot(site) {
  const s = API.snapshot();
  if (Array.isArray(site.sheets)) {
    s.order = s.order.filter((n) => site.sheets.includes(n));
    Object.keys(s.sheets).forEach((n) => { if (!s.order.includes(n)) delete s.sheets[n]; });
    if (!s.order.includes(s.current)) s.current = s.order[0] || '';
  }
  return s;
}
function serialiseData(site) {
  const json = JSON.stringify(publishedSnapshot(site));
  if (json.length > 900000) throw new Error('The sheet data is too large to publish (limit about 900 KB).');
  return json;
}
function setBusy(b) {
  busy = b;
  $('pssDeploy').disabled = b;
  $('pssDeploy').textContent = b ? 'Working…' : (API.getSite() && API.getSite().deployId ? 'Update deployment' : 'Deploy');
}
async function deploy() {
  if (busy) return;
  setBusy(true);
  try {
    persistFiles();
    const user = await ensureAuth();
    const site = siteState();
    site.files = files.map((f) => ({ name: f.name, type: f.type, content: f.content }));
    site.deployId = site.deployId || newId();
    const dataJson = serialiseData(site);
    await setDoc(doc(db, 'deployments', site.deployId), {
      ownerUid: user.uid,
      docId: API.getDocId(),
      name: API.getName(),
      enabled: true,
      writeMode: site.writeMode,
      files: site.files,
      codeVersion: Date.now(),
      dataJson,
      appliedOps: (site.applied || []).slice(-200),
      updatedAt: serverTimestamp()
    });
    site.deployedAt = Date.now();
    API.setSite(site);
    startOps();
    showTab('deploy');
    toast('Deployed');
  } catch (err) {
    log('Deploy failed: ' + (err.code || err.message || err), 'err');
    toast('Deploy failed: ' + (err.message || err.code || 'unknown error'));
    showTab('console');
  } finally {
    setBusy(false);
    renderDeploy();
  }
}
async function undeploy() {
  const site = siteState();
  if (!site.deployId) return;
  if (!confirm('Take this site offline? The link will stop working.')) return;
  try {
    await ensureAuth();
    stopOps();
    const snap = await getDocs(collection(db, 'deployments', site.deployId, 'ops'));
    if (!snap.empty) { const b = writeBatch(db); snap.forEach((d) => b.delete(d.ref)); await b.commit(); }
    await deleteDoc(doc(db, 'deployments', site.deployId));
    delete site.deployId; delete site.deployedAt;
    API.setSite(site);
    toast('Site taken offline');
  } catch (err) {
    toast('Could not take the site offline: ' + (err.message || err.code));
    startOps();
  }
  setBusy(false);
  renderDeploy();
}
async function publish() {
  const site = API.getSite();
  if (!site || !site.deployId || !auth.currentUser) return;
  try {
    await updateDoc(doc(db, 'deployments', site.deployId), {
      dataJson: serialiseData(siteState()),
      appliedOps: (site.applied || []).slice(-200),
      updatedAt: serverTimestamp()
    });
  } catch (err) { console.warn('Publish failed', err); }
}
function schedulePublish() {
  const site = API.getSite();
  if (!site || !site.deployId || !auth.currentUser) return;
  clearTimeout(pubTimer);
  pubTimer = setTimeout(publish, 1500);
}

/* ================= FIREBASE: VISITOR WRITES ================= */
function opAllowed(site, op) {
  if (!op || typeof op.sheet !== 'string' || typeof op.payload !== 'string' || op.payload.length > 20000) return false;
  if (Array.isArray(site.sheets) && !site.sheets.includes(op.sheet)) return false;
  if (site.writeMode === 'full') return ['appendRow', 'setValues', 'clear'].includes(op.type);
  if (site.writeMode === 'append') return op.type === 'appendRow';
  return false;
}
function stopOps() { if (unsubOps) { unsubOps(); unsubOps = null; } }
function startOps() {
  stopOps();
  const site = API.getSite();
  if (!site || !site.deployId || !auth.currentUser) return;
  unsubOps = onSnapshot(
    query(collection(db, 'deployments', site.deployId, 'ops'), orderBy('at')),
    (snap) => {
      const added = snap.docChanges().filter((c) => c.type === 'added').map((c) => ({ id: c.doc.id, ...c.doc.data() }));
      if (added.length) processOps(added);
    },
    (err) => console.warn('Ops listener stopped', err)
  );
}
function processOps(ops) {
  opsChain = opsChain.then(async () => {
    const site = siteState();
    if (!site.deployId) return;
    const applied = new Set(site.applied || []);
    const fresh = ops.filter((o) => !applied.has(o.id) && opAllowed(site, o));
    if (fresh.length) {
      API.applyOps(fresh);
      site.applied = (site.applied || []).concat(fresh.map((o) => o.id)).slice(-300);
      API.setSite(site);
    }
    const batch = writeBatch(db);
    ops.forEach((o) => batch.delete(doc(db, 'deployments', site.deployId, 'ops', o.id)));
    await batch.commit();
    await publish();
    if (fresh.length) toast(fresh.length + ' update' + (fresh.length > 1 ? 's' : '') + ' received from your site');
  }).catch((err) => console.warn('Could not process site writes', err));
}

/* ================= DEPLOYMENT PANEL ================= */
function renderDeploy() {
  const site = siteState();
  const live = !!site.deployId;
  const snap = API.snapshot();
  const user = auth.currentUser;
  const selected = Array.isArray(site.sheets) ? site.sheets : snap.order;
  el.deploy.innerHTML = `
    <div>
      <span class="pss-pill ${live ? 'live' : ''}"><i></i>${live ? 'Live' : 'Not deployed'}</span>
    </div>
    ${live ? `
    <div>
      <h5>Site link</h5>
      <div class="pss-url">
        <input id="pssUrl" readonly value="${esc(siteUrl(site.deployId))}">
        <button class="btn" id="pssCopy">Copy</button>
        <button class="btn primary" id="pssOpen">Open</button>
      </div>
    </div>` : `<div class="pss-hint">Deploy to get a public link. Anyone with the link can open your site; your sheet data stays in sync with it.</div>`}
    <div>
      <h5>Visitor writes</h5>
      <select class="pss-sel" id="pssWrite" style="width:100%">
        <option value="append" ${site.writeMode === 'append' ? 'selected' : ''}>Append rows only (recommended)</option>
        <option value="full" ${site.writeMode === 'full' ? 'selected' : ''}>Full edit (set and clear cells)</option>
        <option value="readonly" ${site.writeMode === 'readonly' ? 'selected' : ''}>Read-only</option>
      </select>
    </div>
    <div>
      <h5>Sheets published to the site</h5>
      <div class="pss-checks" id="pssSheets">
        ${snap.order.map((n) => `<label><input type="checkbox" value="${esc(n)}" ${selected.includes(n) ? 'checked' : ''}> ${esc(n)}</label>`).join('')}
      </div>
    </div>
    <div class="pss-hint">
      ${user ? 'Signed in as ' + esc(user.email || user.displayName || 'Google account') + '.' : 'You will be asked to sign in with Google when you deploy.'}
      The site's code and the published sheets are readable by anyone with the link. Writes from the site
      are applied to this sheet while it is open in Pro Sheet and wait in a queue otherwise.
    </div>
    ${live ? '<div><button class="btn" id="pssUndeploy" style="color:var(--danger)">Take site offline</button></div>' : ''}
  `;
  if (live) {
    $('pssCopy').onclick = async () => { try { await navigator.clipboard.writeText($('pssUrl').value); toast('Link copied'); } catch (e) { $('pssUrl').select(); } };
    $('pssOpen').onclick = () => window.open($('pssUrl').value, '_blank', 'noopener');
    $('pssUndeploy').onclick = undeploy;
  }
  $('pssWrite').onchange = async (e) => {
    const s = siteState();
    s.writeMode = e.target.value;
    API.setSite(s);
    if (s.deployId && auth.currentUser) {
      try { await updateDoc(doc(db, 'deployments', s.deployId), { writeMode: s.writeMode }); toast('Write access updated'); }
      catch (err) { toast('Could not update write access'); }
    }
  };
  $('pssSheets').onchange = () => {
    const s = siteState();
    s.sheets = Array.from($('pssSheets').querySelectorAll('input:checked')).map((i) => i.value);
    API.setSite(s);
    publish();
  };
  setBusy(busy);
}

/* ================= WIRING ================= */
function open() {
  loadFiles();
  root.classList.add('show');
  renderFiles(); loadEditor(); refreshFnList(); renderDeploy();
  setBusy(false);
}
function close() {
  persistFiles();
  stopPreview();
  root.classList.remove('show');
}
$('siteBtn').onclick = open;
$('pssClose').onclick = close;
$('pssSave').onclick = saveNow;
$('pssRun').onclick = runSelected;
$('pssDeploy').onclick = deploy;
$('pssAddGs').onclick = () => addFile('gs');
$('pssAddHtml').onclick = () => addFile('html');

// Pro Sheet calls this after every change to the workbook.
window.__psOnChange = () => { schedulePublish(); schedulePreviewRefresh(); };

onAuthStateChanged(auth, (user) => {
  if (user) { startOps(); schedulePublish(); } else stopOps();
  if (root.classList.contains('show') && activeTab === 'deploy') renderDeploy();
});
