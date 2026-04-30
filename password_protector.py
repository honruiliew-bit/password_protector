#!/usr/bin/env python3
"""
Password Protect Fund & Ecap Folders — HTML Dashboard
Run: python3 password_protect_dashboard.py
Opens http://localhost:7788 automatically.
Auto-installs pyzipper on first run (requires internet).
"""

import http.server, json, os, re, subprocess, sys, threading, webbrowser
from urllib.parse import urlparse, parse_qs

PORT = 7788

# ── Auto-install pyzipper (AES-256 zip encryption) ───────────────────────────
def ensure_pyzipper():
    try:
        import pyzipper
        return True
    except ImportError:
        print("  Installing pyzipper for AES-256 encryption…")
        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'install', 'pyzipper', '--quiet'],
            capture_output=True
        )
        if result.returncode == 0:
            print("  pyzipper installed OK.")
            return True
        else:
            print("  Could not install pyzipper:", result.stderr.decode())
            return False

PYZIPPER_OK = ensure_pyzipper()

# ── Shared state ──────────────────────────────────────────────────────────────
log_lines = []
log_lock  = threading.Lock()
job_state = {"running": False, "ok": 0, "fail": 0}

def push(line):
    with log_lock:
        log_lines.append(line)

# ── Core logic ────────────────────────────────────────────────────────────────

def get_numbered_folders(path):
    out = []
    try:
        for e in os.scandir(path):
            if e.is_dir() and re.fullmatch(r'\d+', e.name):
                out.append({"num": int(e.name), "name": e.name})
    except Exception:
        pass
    return sorted(out, key=lambda x: x["num"])

def find_targets(folder_path):
    hits = []
    try:
        for e in os.scandir(folder_path):
            if e.is_dir() and ('fund' in e.name.lower() or 'ecap' in e.name.lower()):
                hits.append({"path": e.path, "name": e.name})
    except Exception:
        pass
    return hits

def zip_folder(source_dir, out_zip, password):
    """Create an AES-256 encrypted zip using pyzipper."""
    if os.path.exists(out_zip):
        os.remove(out_zip)
    try:
        import pyzipper
        base = os.path.dirname(source_dir)
        with pyzipper.AESZipFile(
            out_zip, 'w',
            compression=pyzipper.ZIP_DEFLATED,
            encryption=pyzipper.WZ_AES
        ) as zf:
            zf.setpassword(password.encode('utf-8'))
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    arcname = os.path.relpath(fp, base)
                    zf.write(fp, arcname)
        return True, f"Created: {os.path.basename(out_zip)}"
    except ImportError:
        return False, "pyzipper not installed — check your internet connection and restart."
    except Exception as ex:
        return False, str(ex)

def run_job(main_path, folder_names, password):
    job_state.update({"running": True, "ok": 0, "fail": 0})
    push(f"SEP=== Processing {len(folder_names)} folder(s) ===")
    ok = fail = 0
    for name in folder_names:
        fp = os.path.join(main_path, name)
        push(f"HDR📁  {name}")
        targets = find_targets(fp)
        if not targets:
            push("WRN  ⚠  No fund/ecap subfolders found — skipping"); continue
        for t in targets:
            zp = os.path.join(fp, f"Secure_{t['name']}.zip")
            success, msg = zip_folder(t['path'], zp, password)
            push(("OK " if success else "ERR") + f"  {'✓' if success else '✗'}  {msg}")
            if success: ok += 1
            else:       fail += 1
    push(f"SEP{'='*40}")
    push(f"OK Done — {ok} zip(s) created, {fail} failed.")
    job_state.update({"running": False, "ok": ok, "fail": fail})
    push("DONE")

# ── HTML ──────────────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Folder Lock</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f5f4f1;--sur:#fff;--sur2:#f1efe8;
  --bd:rgba(0,0,0,.11);--bdm:rgba(0,0,0,.2);
  --tx:#1a1a18;--txm:#5f5e5a;--txh:#888780;
  --acc:#185fa5;--acc-bg:#e6f1fb;--acc-tx:#0c447c;
  --ok-bg:#eaf3de;--ok-tx:#3b6d11;
  --err-bg:#fcebeb;--err-tx:#a32d2d;
  --warn-bg:#faeeda;--warn-tx:#854f0b;
  --r:8px;--rl:12px;
}
@media(prefers-color-scheme:dark){:root{
  --bg:#1c1c1a;--sur:#242422;--sur2:#2c2c2a;
  --bd:rgba(255,255,255,.1);--bdm:rgba(255,255,255,.18);
  --tx:#e8e6df;--txm:#b4b2a9;--txh:#888780;
  --acc:#378add;--acc-bg:#042c53;--acc-tx:#85b7eb;
  --ok-bg:#173404;--ok-tx:#c0dd97;
  --err-bg:#501313;--err-tx:#f7c1c1;
  --warn-bg:#412402;--warn-tx:#fac775;
}}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg);color:var(--tx);font-size:14px;line-height:1.6;padding:2rem 1rem;min-height:100vh}
.page{max-width:760px;margin:0 auto;display:flex;flex-direction:column;gap:14px}
header{display:flex;align-items:center;gap:12px;padding-bottom:14px;border-bottom:.5px solid var(--bd)}
.logo{width:34px;height:34px;background:var(--acc-bg);border-radius:var(--r);display:flex;align-items:center;justify-content:center;flex-shrink:0}
.logo svg{width:18px;height:18px;stroke:var(--acc);fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
h1{font-size:16px;font-weight:500}
.sub{font-size:12px;color:var(--txm)}
.card{background:var(--sur);border:.5px solid var(--bd);border-radius:var(--rl);padding:1.1rem 1.3rem}
.chd{display:flex;align-items:center;gap:8px;margin-bottom:10px}
.nbadge{width:20px;height:20px;border-radius:50%;background:var(--acc-bg);color:var(--acc-tx);font-size:11px;font-weight:500;display:flex;align-items:center;justify-content:center;flex-shrink:0}
.ctitle{font-size:13px;font-weight:500}

/* drop zone */
.dropzone{border:1.5px dashed var(--bdm);border-radius:var(--rl);padding:1.5rem;text-align:center;cursor:pointer;transition:background .15s,border-color .15s;position:relative}
.dropzone:hover,.dropzone.drag-over{background:var(--acc-bg);border-color:var(--acc)}
.dropzone svg{width:28px;height:28px;stroke:var(--txh);margin-bottom:8px}
.dropzone p{font-size:13px;color:var(--txm);margin-bottom:10px}
.dropzone .path-display{font-size:12px;font-family:'SF Mono','Fira Code',monospace;color:var(--acc-tx);background:var(--acc-bg);padding:4px 10px;border-radius:6px;margin-top:8px;word-break:break-all;display:none}
.btn-browse{padding:7px 16px;border-radius:var(--r);border:.5px solid var(--bdm);background:var(--sur2);color:var(--tx);font-size:12px;cursor:pointer;transition:background .12s}
.btn-browse:hover{background:var(--bdm)}

input[type=number],input[type=password]{padding:7px 10px;border:.5px solid var(--bdm);border-radius:var(--r);background:var(--sur2);color:var(--tx);font-size:13px;font-family:inherit;outline:none}
input:focus{border-color:var(--acc)}
.hint{font-size:11px;color:var(--txh);margin-top:5px}
.tags{display:flex;flex-wrap:wrap;gap:5px;margin-top:9px;min-height:22px}
.tag{padding:2px 9px;border-radius:100px;font-size:11px;background:var(--sur2);border:.5px solid var(--bd);color:var(--txm)}
.tag.sel{background:var(--acc-bg);border-color:var(--acc);color:var(--acc-tx)}
.range-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.lbl{font-size:12px;color:var(--txm)}
input[type=number]{width:84px}
.btn{padding:7px 14px;border-radius:var(--r);border:.5px solid var(--bdm);background:var(--sur2);color:var(--tx);font-size:12px;cursor:pointer}
.btn:hover{background:var(--bdm)}
.chip{font-size:11px;color:var(--acc-tx);background:var(--acc-bg);padding:2px 9px;border-radius:100px;display:none}
.pw-row{display:flex;gap:8px;align-items:center}
.pw-row input{flex:1}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-bottom:12px}
.stat{background:var(--sur2);border-radius:var(--r);padding:9px 12px}
.slbl{font-size:11px;color:var(--txh);margin-bottom:2px}
.sval{font-size:20px;font-weight:500}
.primary{padding:9px 22px;border-radius:var(--r);border:.5px solid var(--acc);background:var(--acc);color:#fff;font-size:13px;font-weight:500;cursor:pointer;transition:opacity .12s}
.primary:hover{opacity:.88}
.primary:disabled{opacity:.4;cursor:not-allowed}
.run-row{display:flex;align-items:center;gap:14px}
.spinner{display:none;width:16px;height:16px;border:2px solid var(--bd);border-top-color:var(--acc);border-radius:50%;animation:spin .7s linear infinite;flex-shrink:0}
@keyframes spin{to{transform:rotate(360deg)}}
.stxt{font-size:12px;color:var(--txm)}
.log-wrap{background:var(--sur2);border:.5px solid var(--bd);border-radius:var(--rl);overflow:hidden}
.log-bar{display:flex;justify-content:space-between;align-items:center;padding:5px 10px;border-bottom:.5px solid var(--bd)}
.log-lbl{font-size:11px;color:var(--txm);font-family:'SF Mono','Fira Code',monospace}
.log-clr{font-size:11px;color:var(--txh);background:none;border:none;cursor:pointer;padding:2px 6px}
#log{font-family:'SF Mono','Fira Code',monospace;font-size:12px;color:var(--txm);padding:10px 12px;height:220px;overflow-y:auto;white-space:pre-wrap;word-break:break-all;line-height:1.75}
.l-ok{color:var(--ok-tx)}.l-err{color:var(--err-tx)}.l-hdr{color:var(--tx);font-weight:500}.l-sep{color:var(--txh)}.l-inf{color:var(--acc-tx)}.l-wrn{color:var(--warn-tx)}
</style></head><body>
<div class="page">

<header>
  <div class="logo"><svg viewBox="0 0 24 24"><rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg></div>
  <div><h1>Folder lock</h1><p class="sub">Password-protect fund &amp; ecap subfolders in bulk</p></div>
</header>

<!-- Step 1 -->
<div class="card">
  <div class="chd"><div class="nbadge">1</div><div class="ctitle">Select main folder</div></div>
  <div class="dropzone" id="dropzone" onclick="browsePicker()">
    <svg viewBox="0 0 24 24" style="fill:none;stroke-width:1.5;stroke-linecap:round;stroke-linejoin:round"><path d="M3 7a2 2 0 0 1 2-2h3.17a2 2 0 0 1 1.42.59l1.41 1.41A2 2 0 0 0 12.41 7H19a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>
    <p>Drag your folder here, or click to browse</p>
    <button class="btn-browse" onclick="event.stopPropagation();browsePicker()">Browse…</button>
    <div class="path-display" id="pathDisplay"></div>
  </div>
  <div class="hint" id="folderHint"></div>
  <div class="tags" id="tagList"></div>
</div>

<!-- Step 2 -->
<div class="card">
  <div class="chd"><div class="nbadge">2</div><div class="ctitle">Choose folder range</div></div>
  <div class="range-row">
    <span class="lbl">From</span>
    <input type="number" id="rangeFrom" placeholder="2298">
    <span class="lbl">to</span>
    <input type="number" id="rangeTo" placeholder="2313">
    <button class="btn" onclick="previewRange()">Preview</button>
    <span class="chip" id="chip"></span>
  </div>
  <div class="tags" id="selTags"></div>
</div>

<!-- Step 3 -->
<div class="card">
  <div class="chd"><div class="nbadge">3</div><div class="ctitle">Set password</div></div>
  <div class="pw-row">
    <input type="password" id="password" placeholder="Enter a strong password">
    <button class="btn" id="showBtn" onclick="togglePw()">Show</button>
  </div>
</div>

<!-- Step 4 -->
<div class="card">
  <div class="chd"><div class="nbadge">4</div><div class="ctitle">Run</div></div>
  <div class="stats">
    <div class="stat"><div class="slbl">Folders selected</div><div class="sval" id="sFolders">—</div></div>
    <div class="stat"><div class="slbl">Zips created</div><div class="sval l-ok" id="sOk">—</div></div>
    <div class="stat"><div class="slbl">Failures</div><div class="sval l-err" id="sFail">—</div></div>
  </div>
  <div class="run-row">
    <button class="primary" id="runBtn" onclick="runJob()">Lock folders</button>
    <div class="spinner" id="spinner"></div>
    <span class="stxt" id="stxt"></span>
  </div>
</div>

<!-- Log -->
<div class="log-wrap">
  <div class="log-bar"><span class="log-lbl">Activity log</span><button class="log-clr" onclick="clearLog()">Clear</button></div>
  <div id="log"></div>
</div>

</div>
<script>
let allFolders = [], selFolders = [], logIdx = 0, polling = null;
let folderPath = '';

// ── Drag & drop ──────────────────────────────────────────────────────────────
const dz = document.getElementById('dropzone');
dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
dz.addEventListener('drop', async e => {
  e.preventDefault(); dz.classList.remove('drag-over');
  // Try file:// URI from text/uri-list (works on macOS Chrome/Safari)
  let uri = e.dataTransfer.getData('text/uri-list') || e.dataTransfer.getData('text/plain');
  if (uri && uri.startsWith('file://')) {
    const p = decodeURIComponent(uri.replace(/^file:\/\//, '').split('\\n')[0].trim());
    if (p) { setPath(p); return; }
  }
  // Fallback: try File System Access API entry name for hint
  const items = e.dataTransfer.items;
  if (items && items.length) {
    const entry = items[0].webkitGetAsEntry && items[0].webkitGetAsEntry();
    if (entry && entry.isDirectory) {
      appendLog('WRN Browser blocked the full path. Please use Browse… button instead.', true);
    }
  }
});

async function browsePicker() {
  const res = await fetch('/browse');
  const d = await res.json();
  if (d.path) setPath(d.path);
  else if (d.error) appendLog('ERR ' + d.error, true);
}

function setPath(p) {
  folderPath = p;
  const disp = document.getElementById('pathDisplay');
  disp.textContent = p; disp.style.display = 'block';
  document.getElementById('folderHint').textContent = 'Scanning…';
  scanFolder(p);
}

// ── Folder scan ───────────────────────────────────────────────────────────────
async function scanFolder(path) {
  const res = await fetch('/scan?path=' + encodeURIComponent(path));
  const d = await res.json();
  if (d.error) { document.getElementById('folderHint').textContent = '⚠ ' + d.error; return; }
  allFolders = d.folders;
  renderTags('tagList', allFolders, false);
  if (!allFolders.length) {
    document.getElementById('folderHint').textContent = 'No numeric subfolders found.'; return;
  }
  const nums = allFolders.map(f => f.num);
  document.getElementById('rangeFrom').value = Math.min(...nums);
  document.getElementById('rangeTo').value   = Math.max(...nums);
  document.getElementById('folderHint').textContent =
    `Found ${allFolders.length} numbered folders (${Math.min(...nums)} – ${Math.max(...nums)})`;
  previewRange();
}

// ── Range preview ─────────────────────────────────────────────────────────────
function previewRange() {
  const lo = parseInt(document.getElementById('rangeFrom').value);
  const hi = parseInt(document.getElementById('rangeTo').value);
  if (isNaN(lo) || isNaN(hi)) return;
  const mn = Math.min(lo,hi), mx = Math.max(lo,hi);
  selFolders = allFolders.filter(f => f.num >= mn && f.num <= mx);
  renderTags('selTags', selFolders, true);
  document.getElementById('sFolders').textContent = selFolders.length;
  const chip = document.getElementById('chip');
  chip.textContent = selFolders.length + ' folder' + (selFolders.length !== 1 ? 's' : '');
  chip.style.display = 'inline-block';
}

function renderTags(id, folders, highlight) {
  const el = document.getElementById(id); el.innerHTML = '';
  folders.forEach(f => {
    const t = document.createElement('span');
    t.className = 'tag' + (highlight ? ' sel' : '');
    t.textContent = f.name; el.appendChild(t);
  });
}

// ── Password toggle ───────────────────────────────────────────────────────────
function togglePw() {
  const inp = document.getElementById('password'), btn = document.getElementById('showBtn');
  inp.type = inp.type === 'password' ? 'text' : 'password';
  btn.textContent = inp.type === 'password' ? 'Show' : 'Hide';
}

// ── Run ───────────────────────────────────────────────────────────────────────
async function runJob() {
  if (!folderPath)       { appendLog('ERR Please select a folder first.', true); return; }
  if (!selFolders.length){ appendLog('ERR No folders in range.', true); return; }
  const pw = document.getElementById('password').value;
  if (!pw)               { appendLog('ERR Please enter a password.', true); return; }

  if (!confirm('Create password-protected zips for ' + selFolders.length +
      ' folder(s)?\\n' + selFolders.map(f=>f.name).join(', '))) return;

  document.getElementById('runBtn').disabled = true;
  document.getElementById('spinner').style.display = 'block';
  document.getElementById('stxt').textContent = 'Working…';
  document.getElementById('sOk').textContent = '0';
  document.getElementById('sFail').textContent = '0';
  logIdx = 0;

  await fetch('/run', {
    method: 'POST',
    headers: {'Content-Type':'application/json'},
    body: JSON.stringify({ main_path: folderPath, folders: selFolders.map(f=>f.name), password: pw })
  });

  polling = setInterval(pollLogs, 300);
}

async function pollLogs() {
  const res = await fetch('/logs?since=' + logIdx);
  const d = await res.json();
  d.lines.forEach(line => {
    appendLog(line);
    if (line === 'DONE') finishJob();
  });
  logIdx += d.lines.length;
}

function appendLog(raw, force) {
  const el = document.getElementById('log');
  const div = document.createElement('div');
  let cls = '', text = raw;
  if      (raw.startsWith('OK '))  { cls = 'l-ok';  text = raw.slice(3); }
  else if (raw.startsWith('ERR'))  { cls = 'l-err'; text = raw.slice(3); }
  else if (raw.startsWith('HDR'))  { cls = 'l-hdr'; text = raw.slice(3); }
  else if (raw.startsWith('SEP'))  { cls = 'l-sep'; text = raw.slice(3); }
  else if (raw.startsWith('WRN'))  { cls = 'l-wrn'; text = raw.slice(3); }
  else if (raw === 'DONE') return;
  if (cls) div.className = cls;
  div.textContent = text;
  el.appendChild(div);
  el.scrollTop = el.scrollHeight;
}

function finishJob() {
  clearInterval(polling); polling = null;
  fetch('/state').then(r=>r.json()).then(d => {
    document.getElementById('sOk').textContent   = d.ok;
    document.getElementById('sFail').textContent = d.fail;
    document.getElementById('runBtn').disabled   = false;
    document.getElementById('spinner').style.display = 'none';
    document.getElementById('stxt').textContent  = 'Finished.';
  });
}

function clearLog() {
  document.getElementById('log').innerHTML = '';
  fetch('/clear-log');
}
</script></body></html>"""

# ── HTTP handler ──────────────────────────────────────────────────────────────

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def send_json(self, data, status=200):
        body = json.dumps(data).encode()
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        p = urlparse(self.path)
        qs = parse_qs(p.query)

        if p.path == '/':
            body = HTML.encode()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', len(body))
            self.end_headers()
            self.wfile.write(body)

        elif p.path == '/browse':
            # Open native macOS folder picker via AppleScript
            try:
                r = subprocess.run(
                    ['osascript', '-e',
                     'POSIX path of (choose folder with prompt "Select main folder")'],
                    capture_output=True, text=True, timeout=120
                )
                path = r.stdout.strip()
                if path:
                    self.send_json({'path': path})
                else:
                    self.send_json({'path': '', 'error': 'No folder selected'})
            except Exception as ex:
                self.send_json({'path': '', 'error': str(ex)})

        elif p.path == '/scan':
            folder = qs.get('path', [''])[0]
            if not os.path.isdir(folder):
                self.send_json({'error': f'Not found: {folder}'}); return
            self.send_json({'folders': get_numbered_folders(folder)})

        elif p.path == '/logs':
            since = int(qs.get('since', ['0'])[0])
            with log_lock:
                lines = log_lines[since:]
            self.send_json({'lines': lines})

        elif p.path == '/state':
            self.send_json(dict(job_state))

        elif p.path == '/clear-log':
            with log_lock: log_lines.clear()
            self.send_json({'ok': True})

        else:
            self.send_response(404); self.end_headers()

    def do_POST(self):
        if urlparse(self.path).path == '/run':
            n = int(self.headers.get('Content-Length', 0))
            body = json.loads(self.rfile.read(n))
            with log_lock: log_lines.clear()
            threading.Thread(
                target=run_job,
                args=(body['main_path'], body['folders'], body['password']),
                daemon=True
            ).start()
            self.send_json({'ok': True})
        else:
            self.send_response(404); self.end_headers()


# ── Entry ─────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    server = http.server.ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    url = f'http://localhost:{PORT}'
    print(f'\n  Folder Lock  →  {url}\n  Press Ctrl+C to stop.\n')
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n  Stopped.')