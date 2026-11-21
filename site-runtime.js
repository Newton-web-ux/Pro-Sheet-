/*
 * Pro Sheet site runtime.
 * Runs Code.gs (an Apps Script compatible subset) and serves index.html, each inside a
 * sandboxed iframe with an opaque origin. The host page mediates all sheet data and writes.
 */

const SANDBOX_SERVER = 'allow-scripts';
const SANDBOX_SITE = 'allow-scripts allow-forms allow-modals allow-popups allow-popups-to-escape-sandbox allow-downloads';

/* ---------- Runs inside the sandboxed "server" iframe (serialised with toString) ---------- */
function serverMain() {
  var FILES = {}, DATA = { name: '', sheets: {}, order: [], current: '' }, OPS = [], LOGS = [];
  var currentFile = '', loadErrors = [];

  function str(v) { if (typeof v === 'string') return v; try { return JSON.stringify(v); } catch (e) { return String(v); } }
  function logFn() { LOGS.push(Array.prototype.map.call(arguments, str).join(' ')); }
  window.console = { log: logFn, info: logFn, warn: logFn, error: logFn, debug: logFn };
  window.Logger = { log: logFn };
  window.addEventListener('error', function (e) {
    loadErrors.push((currentFile || 'script') + ': ' + e.message + (e.lineno ? ' (line ' + e.lineno + ')' : ''));
  });

  /* ---- grid helpers (1-indexed rows/cols) ---- */
  function nonEmpty(v) { return v !== '' && v !== null && v !== undefined; }
  function lastRow(g) { for (var r = g.length - 1; r >= 0; r--) { if (g[r].some(nonEmpty)) return r + 1; } return 0; }
  function lastCol(g) {
    var m = 0;
    g.forEach(function (row) { for (var c = row.length - 1; c >= 0; c--) { if (nonEmpty(row[c])) { if (c + 1 > m) m = c + 1; break; } } });
    return m;
  }
  function get(g, r, c) { var row = g[r - 1]; var v = row ? row[c - 1] : undefined; return v === undefined ? '' : v; }
  function put(g, r, c, v) {
    while (g.length < r) g.push([]);
    var row = g[r - 1];
    while (row.length < c) row.push('');
    row[c - 1] = v;
  }
  function clean(v) {
    if (v instanceof Date) return v.toISOString();
    if (v === null || v === undefined) return '';
    if (typeof v === 'number' || typeof v === 'boolean' || typeof v === 'string') return v;
    return String(v);
  }
  function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }

  /* ---- write operations: applied locally right away and reported to the host ---- */
  function applyLocal(op) {
    var g = DATA.sheets[op.sheet];
    if (!g) return;
    var p = JSON.parse(op.payload);
    if (op.type === 'appendRow') {
      var r = lastRow(g) + 1;
      p.values.forEach(function (v, i) { put(g, r, i + 1, v); });
    } else if (op.type === 'setValues') {
      p.values.forEach(function (row, ri) { row.forEach(function (v, ci) { put(g, p.row + ri, p.col + ci, v); }); });
    } else if (op.type === 'clear') {
      for (var a = 0; a < p.numRows; a++) for (var b = 0; b < p.numCols; b++) put(g, p.row + a, p.col + b, '');
    }
  }
  function emit(type, sheet, payload) {
    var op = { id: uid(), type: type, sheet: sheet, payload: JSON.stringify(payload) };
    OPS.push(op);
    applyLocal(op);
  }

  /* ---- A1 notation ---- */
  function colNum(s) { var n = 0; s = s.toUpperCase(); for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; }
  function parseA1(a) {
    var m = /^\s*(?:.*!)?\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?\s*$/.exec(a);
    if (!m) throw new Error('Range not found: ' + a);
    var c1 = colNum(m[1]), r1 = +m[2], c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    return [Math.min(r1, r2), Math.min(c1, c2), Math.abs(r2 - r1) + 1, Math.abs(c2 - c1) + 1];
  }

  /* ---- Range ---- */
  function Range(sheet, r, c, nr, nc) { this.sheet = sheet; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  Range.prototype.getValues = function () {
    var g = DATA.sheets[this.sheet], out = [];
    for (var i = 0; i < this.nr; i++) {
      var row = [];
      for (var j = 0; j < this.nc; j++) row.push(get(g, this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  };
  Range.prototype.getDisplayValues = function () { return this.getValues().map(function (row) { return row.map(String); }); };
  Range.prototype.getValue = function () { return get(DATA.sheets[this.sheet], this.r, this.c); };
  Range.prototype.setValues = function (v) {
    var nc = this.nc;
    if (!Array.isArray(v) || v.length !== this.nr) {
      throw new Error('The number of rows in the data does not match the number of rows in the range. The data has ' + (Array.isArray(v) ? v.length : 0) + ' but the range has ' + this.nr + '.');
    }
    v.forEach(function (row) {
      if (!Array.isArray(row) || row.length !== nc) {
        throw new Error('The number of columns in the data does not match the number of columns in the range. The data has ' + (Array.isArray(row) ? row.length : 0) + ' but the range has ' + nc + '.');
      }
    });
    emit('setValues', this.sheet, { row: this.r, col: this.c, values: v.map(function (row) { return row.map(clean); }) });
    return this;
  };
  Range.prototype.setValue = function (v) {
    var rows = [];
    for (var i = 0; i < this.nr; i++) { var row = []; for (var j = 0; j < this.nc; j++) row.push(v); rows.push(row); }
    return this.setValues(rows);
  };
  Range.prototype.clearContent = Range.prototype.clear = function () {
    emit('clear', this.sheet, { row: this.r, col: this.c, numRows: this.nr, numCols: this.nc });
    return this;
  };
  Range.prototype.getRow = function () { return this.r; };
  Range.prototype.getColumn = function () { return this.c; };
  Range.prototype.getNumRows = function () { return this.nr; };
  Range.prototype.getNumColumns = function () { return this.nc; };

  /* ---- Sheet / Spreadsheet ---- */
  function Sheet(name) { this.name = name; }
  Sheet.prototype = {
    getName: function () { return this.name; },
    getLastRow: function () { return lastRow(DATA.sheets[this.name]); },
    getLastColumn: function () { return lastCol(DATA.sheets[this.name]); },
    getRange: function (a, b, c, d) {
      if (typeof a === 'string') { var p = parseA1(a); return new Range(this.name, p[0], p[1], p[2], p[3]); }
      if (!(a >= 1) || !(b >= 1)) throw new Error('The coordinates of the range are outside the dimensions of the sheet.');
      return new Range(this.name, a, b, c || 1, d || 1);
    },
    getDataRange: function () {
      var g = DATA.sheets[this.name];
      return new Range(this.name, 1, 1, Math.max(1, lastRow(g)), Math.max(1, lastCol(g)));
    },
    getSheetValues: function (r, c, nr, nc) { return this.getRange(r, c, nr, nc).getValues(); },
    appendRow: function (values) {
      if (!Array.isArray(values)) throw new Error('Cannot convert to Object[].');
      emit('appendRow', this.name, { values: values.map(clean) });
      return this;
    },
    clear: function () {
      var g = DATA.sheets[this.name];
      emit('clear', this.name, { row: 1, col: 1, numRows: Math.max(1, lastRow(g)), numCols: Math.max(1, lastCol(g)) });
      return this;
    }
  };
  Sheet.prototype.clearContents = Sheet.prototype.clear;

  var spreadsheet = {
    getName: function () { return DATA.name || 'Spreadsheet'; },
    getId: function () { return 'prosheet'; },
    getUrl: function () { return ''; },
    getSheets: function () { return DATA.order.map(function (n) { return new Sheet(n); }); },
    getSheetByName: function (n) { return DATA.sheets[n] ? new Sheet(n) : null; },
    getActiveSheet: function () {
      var n = DATA.sheets[DATA.current] ? DATA.current : DATA.order[0];
      if (!n) throw new Error('This spreadsheet has no published sheets.');
      return new Sheet(n);
    }
  };
  window.SpreadsheetApp = {
    getActiveSpreadsheet: function () { return spreadsheet; },
    getActive: function () { return spreadsheet; },
    openById: function () { return spreadsheet; },
    getActiveSheet: function () { return spreadsheet.getActiveSheet(); }
  };

  /* ---- misc services ---- */
  window.Utilities = {
    getUuid: function () { return (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : uid(); },
    formatDate: function (d, tz, fmt) {
      d = new Date(d);
      function p(n) { return ('0' + n).slice(-2); }
      return fmt.replace(/yyyy|MM|dd|HH|mm|ss/g, function (t) {
        return { yyyy: d.getFullYear(), MM: p(d.getMonth() + 1), dd: p(d.getDate()), HH: p(d.getHours()), mm: p(d.getMinutes()), ss: p(d.getSeconds()) }[t];
      });
    },
    sleep: function () {}
  };
  window.Session = {
    getActiveUser: function () { return { getEmail: function () { return ''; } }; },
    getScriptTimeZone: function () { return Intl.DateTimeFormat().resolvedOptions().timeZone; }
  };
  var lock = { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {}, hasLock: function () { return true; } };
  window.LockService = { getScriptLock: function () { return lock; }, getDocumentLock: function () { return lock; }, getUserLock: function () { return lock; } };

  /* ---- ContentService ---- */
  function TextOutput(s) { this.__type = 'text'; this.content = String(s); this.mime = 'text/plain'; }
  TextOutput.prototype = {
    setMimeType: function (m) { this.mime = String(m); return this; },
    getContent: function () { return this.content; },
    append: function (s) { this.content += s; return this; }
  };
  window.ContentService = {
    createTextOutput: function (s) { return new TextOutput(s === undefined ? '' : s); },
    MimeType: { TEXT: 'text/plain', JSON: 'application/json', CSV: 'text/csv', XML: 'application/xml', JAVASCRIPT: 'application/javascript' }
  };

  /* ---- HtmlService + templates (<?= ?>, <?!= ?>, <? ?>) ---- */
  function esc(s) {
    return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (m) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m];
    });
  }
  function HtmlOutput(c) { this.__type = 'html'; this.content = c; this.title = ''; }
  HtmlOutput.prototype = {
    getContent: function () { return this.content; },
    setTitle: function (t) { this.title = String(t); return this; },
    append: function (s) { this.content += s; return this; },
    setXFrameOptionsMode: function () { return this; },
    addMetaTag: function () { return this; },
    setFaviconUrl: function () { return this; }
  };
  function fileContent(name) {
    var key = FILES[name] !== undefined ? name : (FILES[name + '.html'] !== undefined ? name + '.html' : null);
    if (key === null) throw new Error('No HTML file named ' + name);
    return FILES[key];
  }
  function runTemplate(src, tpl) {
    var re = /<\?(!=|=)?([\s\S]*?)\?>/g, code = 'var __o=[];', last = 0, m;
    while ((m = re.exec(src))) {
      code += '__o.push(' + JSON.stringify(src.slice(last, m.index)) + ');';
      if (m[1] === '=') code += '__o.push(__esc(' + m[2] + '));';
      else if (m[1] === '!=') code += '__o.push((' + m[2] + '));';
      else code += m[2] + '\n';
      last = re.lastIndex;
    }
    code += '__o.push(' + JSON.stringify(src.slice(last)) + ');return __o.join("");';
    return new Function('__esc', '__t', 'with(__t){' + code + '}')(esc, tpl);
  }
  function Template(src) { this.__src = src; }
  Template.prototype.evaluate = function () { return new HtmlOutput(runTemplate(this.__src, this)); };
  window.HtmlService = {
    createHtmlOutput: function (s) { return new HtmlOutput(String(s === undefined ? '' : s)); },
    createHtmlOutputFromFile: function (n) { return new HtmlOutput(fileContent(n)); },
    createTemplate: function (s) { return new Template(String(s)); },
    createTemplateFromFile: function (n) { return new Template(fileContent(n)); },
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' }
  };

  /* ---- messaging with the host ---- */
  function send(m) { parent.postMessage(m, '*'); }
  function inject(name, code) {
    currentFile = name;
    var s = document.createElement('script');
    s.textContent = code;
    document.head.appendChild(s);
    currentFile = '';
  }
  function serialize(res) {
    if (res instanceof HtmlOutput || res instanceof TextOutput) return res;
    if (res === undefined) return null;
    return JSON.parse(JSON.stringify(res));
  }

  window.addEventListener('message', function (e) {
    if (e.source !== parent) return;
    var m = e.data || {};
    if (m.ps === 'init') {
      FILES = m.html || {};
      DATA = m.data || DATA;
      loadErrors = [];
      (m.gs || []).forEach(function (f) { inject(f.name, f.content); });
      send({ ps: 'inited', errors: loadErrors.slice() });
    } else if (m.ps === 'data') {
      DATA = m.data;
      (m.pending || []).forEach(function (op) { try { applyLocal(op); } catch (x) {} });
    } else if (m.ps === 'call') {
      OPS = []; LOGS = [];
      var out = { ps: 'result', id: m.id, ok: true };
      try {
        if (!/^[A-Za-z_$][\w$]*$/.test(String(m.fn))) throw new Error('Invalid function name');
        var f;
        try { f = (0, eval)(m.fn); } catch (x) { f = undefined; }
        if (typeof f !== 'function') {
          if (m.fn === 'doGet') out.result = serialize(new Template(fileContent('index')).evaluate());
          else throw new Error('Script function not found: ' + m.fn);
        } else {
          out.result = serialize(f.apply(null, m.args || []));
        }
      } catch (err) {
        out.ok = false;
        out.error = (err && err.message) ? err.message : String(err);
      }
      out.ops = OPS; out.logs = LOGS;
      send(out);
    }
  });
  send({ ps: 'ready' });
}

/* ---------- Injected into the site iframe: google.script.run bridge ---------- */
function clientShim() {
  var cbs = {}, n = 0, listeners = [];
  function runner(ok, fail, uo) {
    return new Proxy({}, {
      get: function (_, prop) {
        if (prop === 'withSuccessHandler') return function (f) { return runner(f, fail, uo); };
        if (prop === 'withFailureHandler') return function (f) { return runner(ok, f, uo); };
        if (prop === 'withUserObject') return function (o) { return runner(ok, fail, o); };
        if (typeof prop !== 'string') return undefined;
        return function () {
          var id = ++n;
          cbs[id] = { ok: ok, fail: fail, uo: uo };
          parent.postMessage({ ps: 'run', id: id, fn: prop, args: Array.prototype.slice.call(arguments) }, '*');
        };
      }
    });
  }
  window.google = {
    script: {
      get run() { return runner(null, null, null); },
      host: { close: function () {}, setHeight: function () {}, setWidth: function () {} }
    }
  };
  window.ProSheet = { onChange: function (f) { listeners.push(f); } };
  window.addEventListener('message', function (e) {
    if (e.source !== parent) return;
    var d = e.data;
    if (!d || !d.ps) return;
    if (d.ps === 'result') {
      var cb = cbs[d.id]; delete cbs[d.id];
      if (!cb) return;
      if (d.ok) { if (cb.ok) cb.ok(d.result, cb.uo); }
      else if (cb.fail) cb.fail(new Error(d.error), cb.uo);
      else console.error(d.error);
    } else if (d.ps === 'sheet-changed') {
      listeners.forEach(function (f) { try { f(); } catch (x) { console.error(x); } });
      window.dispatchEvent(new CustomEvent('prosheet:change'));
    }
  });
}

/* ---------- Host side ---------- */
function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
}
function errorPage(msg, errs) {
  return '<!doctype html><meta charset="utf-8"><body style="font:14px system-ui,sans-serif;padding:24px;color:#b3261e">' +
    '<h3 style="margin:0 0 8px">Script error</h3><pre style="white-space:pre-wrap;margin:0">' +
    escapeHtml([msg].concat(errs || []).join('\n')) + '</pre>';
}
function textPage(text) {
  return '<!doctype html><meta charset="utf-8"><body style="margin:0;padding:16px;font:13px ui-monospace,monospace"><pre style="white-space:pre-wrap;margin:0">' +
    escapeHtml(text) + '</pre>';
}
function withShim(content, title) {
  const shim = '<script>(' + clientShim.toString() + ')();<\/script>';
  const t = title ? '<title>' + escapeHtml(title) + '</title>' : '';
  if (!/^\s*<!doctype/i.test(content)) content = '<!doctype html>' + content;
  if (/<head[^>]*>/i.test(content)) return content.replace(/<head[^>]*>/i, (m) => m + shim + t);
  return content.replace(/^(\s*<!doctype[^>]*>)/i, (m) => m + shim + t);
}

/**
 * createRuntime({ files, snapshot, params, mount, onOps, onLog, onTitle })
 *  files:    [{ name, type: 'gs' | 'html', content }]
 *  snapshot: { name, sheets: { [sheet]: any[][] }, order: string[], current: string }
 *  mount:    element to render the site into (omit for headless runs)
 * Returns { ready, call, render, setSnapshot, destroy }.
 */
export function createRuntime(opts) {
  const o = Object.assign({
    files: [], snapshot: { name: '', sheets: {}, order: [], current: '' }, params: {},
    mount: null, onOps() {}, onLog() {}, onTitle() {}
  }, opts);
  const waiting = new Map();
  let seq = 0, site = null, destroyed = false;

  const server = document.createElement('iframe');
  server.setAttribute('sandbox', SANDBOX_SERVER);
  server.style.display = 'none';
  server.srcdoc = '<!doctype html><meta charset="utf-8"><script>(' + serverMain.toString() + ')();<\/script>';

  const readyPromise = new Promise((res) => waiting.set('ready', res));

  function onMessage(e) {
    const m = e.data;
    if (!m || typeof m !== 'object') return;
    if (e.source === server.contentWindow) {
      if (m.ps === 'ready') { const r = waiting.get('ready'); waiting.delete('ready'); r && r(); }
      else if (m.ps === 'inited') { const r = waiting.get('init'); waiting.delete('init'); r && r(m); }
      else if (m.ps === 'result') { const r = waiting.get(m.id); waiting.delete(m.id); r && r(m); }
    } else if (site && e.source === site.contentWindow && m.ps === 'run') {
      handleRun(m);
    }
  }
  window.addEventListener('message', onMessage);
  document.body.appendChild(server);

  const inited = (async () => {
    await readyPromise;
    const gs = o.files.filter((f) => f.type === 'gs').map((f) => ({ name: f.name, content: f.content }));
    const html = {};
    o.files.filter((f) => f.type === 'html').forEach((f) => { html[f.name] = f.content; });
    const p = new Promise((res) => waiting.set('init', res));
    server.contentWindow.postMessage({ ps: 'init', gs, html, data: o.snapshot }, '*');
    return (await p).errors || [];
  })();

  async function call(fn, args) {
    await inited;
    if (destroyed) return { ok: false, error: 'Runtime stopped', ops: [], logs: [] };
    const id = ++seq;
    const p = new Promise((res) => waiting.set(id, res));
    server.contentWindow.postMessage({ ps: 'call', id, fn: String(fn), args: args || [] }, '*');
    const r = await p;
    if (r.logs && r.logs.length) o.onLog(r.logs);
    if (r.ops && r.ops.length) { try { await o.onOps(r.ops); } catch (err) { console.warn(err); } }
    return r;
  }

  async function handleRun(m) {
    const reply = (p) => site && site.contentWindow && site.contentWindow.postMessage(Object.assign({ ps: 'result', id: m.id }, p), '*');
    if (typeof m.fn !== 'string' || m.fn.endsWith('_')) return reply({ ok: false, error: 'Script function not found: ' + m.fn });
    const r = await call(m.fn, Array.isArray(m.args) ? m.args : []);
    reply(r.ok ? { ok: true, result: r.result } : { ok: false, error: r.error });
  }

  async function render() {
    if (!o.mount) return { ok: false, error: 'No mount element' };
    const params = o.params || {};
    const e = {
      parameter: Object.assign({}, params),
      parameters: Object.fromEntries(Object.entries(params).map(([k, v]) => [k, [v]])),
      queryString: new URLSearchParams(params).toString(),
      pathInfo: '', contextPath: ''
    };
    const errs = await inited;
    const r = await call('doGet', [e]);
    let html;
    if (!r.ok) html = errorPage(r.error, errs);
    else if (r.result && r.result.__type === 'html') { html = withShim(r.result.content, r.result.title); o.onTitle(r.result.title || ''); }
    else if (r.result && r.result.__type === 'text') html = textPage(r.result.content);
    else html = errorPage('doGet must return HtmlService or ContentService output.', errs);
    if (site) site.remove();
    site = document.createElement('iframe');
    site.setAttribute('sandbox', SANDBOX_SITE);
    site.style.cssText = 'width:100%;height:100%;border:0;display:block;background:#fff';
    site.srcdoc = html;
    o.mount.replaceChildren(site);
    return r;
  }

  async function setSnapshot(snap, pending) {
    await inited;
    if (destroyed) return;
    server.contentWindow.postMessage({ ps: 'data', data: snap, pending: pending || [] }, '*');
    if (site && site.contentWindow) site.contentWindow.postMessage({ ps: 'sheet-changed' }, '*');
  }

  function destroy() {
    destroyed = true;
    window.removeEventListener('message', onMessage);
    server.remove();
    if (site) site.remove();
  }

  return { ready: inited, call, render, setSnapshot, destroy };
}
