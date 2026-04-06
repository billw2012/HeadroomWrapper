"""
Headroom Windows Service
========================
Manages the headroom proxy process and serves a local dashboard.

Usage:
    python headroom_service.py debug       # Run in foreground (no admin needed)
    python headroom_service.py install     # Install as Windows service (run as admin)
    python headroom_service.py start       # Start the service
    python headroom_service.py stop        # Stop the service
    python headroom_service.py remove      # Uninstall the service
    python headroom_service.py status      # Show service status

Dashboard: http://127.0.0.1:8788
"""

import sys
import os
import time
import subprocess
import threading
import json
import shutil
from collections import deque
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.request import urlopen, Request
from urllib.error import URLError

try:
    import win32serviceutil
    import win32service
    import win32event
    import servicemanager
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HEADROOM_PORT = 8787
DASHBOARD_PORT = 8788
LOG_BUFFER_SIZE = 500
SERVICE_NAME = "HeadroomProxy"
SERVICE_DISPLAY = "Headroom Proxy Service"
SERVICE_DESC = "Runs the Headroom context optimization proxy and web dashboard."

# ---------------------------------------------------------------------------
# Rolling log buffer (shared between manager and dashboard)
# ---------------------------------------------------------------------------

_log_buffer: deque = deque(maxlen=LOG_BUFFER_SIZE)
_log_lock = threading.Lock()


def _add_log(line: str) -> None:
    with _log_lock:
        _log_buffer.append(line)


def _get_logs() -> list:
    with _log_lock:
        return list(_log_buffer)


# ---------------------------------------------------------------------------
# Headroom process manager
# ---------------------------------------------------------------------------

class HeadroomManager(threading.Thread):
    def __init__(self, stop_event: threading.Event):
        super().__init__(daemon=True, name="HeadroomManager")
        self.stop_event = stop_event
        self.process: subprocess.Popen | None = None
        self.headroom_exe: str | None = shutil.which("headroom")

    def run(self) -> None:
        backoff = 2
        while not self.stop_event.is_set():
            if not self.headroom_exe:
                _add_log("[service] ERROR: 'headroom' not found in PATH. Retrying in 30s...")
                self.stop_event.wait(30)
                self.headroom_exe = shutil.which("headroom")
                continue

            _add_log(f"[service] Starting: {self.headroom_exe} proxy --port {HEADROOM_PORT}")
            try:
                flags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
                self.process = subprocess.Popen(
                    [self.headroom_exe, "proxy", "--port", str(HEADROOM_PORT)],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    encoding="utf-8",
                    errors="replace",
                    creationflags=flags,
                )
                backoff = 2  # reset on successful launch

                for line in self.process.stdout:
                    line = line.rstrip()
                    if line:
                        _add_log(line)
                    if self.stop_event.is_set():
                        break

                self.process.wait()
                if not self.stop_event.is_set():
                    _add_log(
                        f"[service] headroom exited (code {self.process.returncode}). "
                        f"Restarting in {backoff}s..."
                    )
                    self.stop_event.wait(backoff)
                    backoff = min(backoff * 2, 60)

            except Exception as exc:
                _add_log(f"[service] Failed to start headroom: {exc}")
                self.stop_event.wait(backoff)
                backoff = min(backoff * 2, 60)

    def stop(self) -> None:
        if self.process and self.process.poll() is None:
            _add_log("[service] Terminating headroom proxy...")
            self.process.terminate()
            try:
                self.process.wait(timeout=10)
            except subprocess.TimeoutExpired:
                _add_log("[service] Force-killing headroom proxy...")
                self.process.kill()


# ---------------------------------------------------------------------------
# Dashboard HTTP server
# ---------------------------------------------------------------------------

DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Headroom Dashboard</title>
<style>
  :root {
    --bg:#0f1117; --card:#1a1d27; --border:#2d3148;
    --accent:#6366f1; --text:#e2e8f0; --muted:#94a3b8;
    --green:#22c55e; --red:#ef4444; --yellow:#f59e0b;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;padding:24px;min-height:100vh}
  h1{font-size:1.4rem;font-weight:700;letter-spacing:-0.01em}
  .subtitle{color:var(--muted);font-size:0.82rem;margin-top:2px}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;flex-wrap:wrap;gap:12px}
  .header-actions{display:flex;gap:8px;align-items:center}
  /* lifetime savings banner */
  .banner{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:20px;background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px 20px}
  .bstat{text-align:center}
  .bstat .val{font-size:1.7rem;font-weight:700;font-variant-numeric:tabular-nums;color:var(--green)}
  .bstat .lbl{font-size:0.72rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-top:3px}
  .bstat .sub{font-size:0.75rem;color:var(--muted);margin-top:2px}
  .banner-title{font-size:0.65rem;text-transform:uppercase;letter-spacing:0.1em;color:var(--muted);font-weight:600;grid-column:1/-1;margin-bottom:4px}
  /* cards */
  .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:14px;margin-bottom:20px}
  .card{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:18px}
  .card-title{font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);margin-bottom:12px;font-weight:600}
  .big{font-size:2.2rem;font-weight:700;line-height:1;font-variant-numeric:tabular-nums}
  .big-label{color:var(--muted);font-size:0.78rem;margin-top:4px}
  .rows{margin-top:12px}
  .row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--border);font-size:0.82rem}
  .row:last-child{border-bottom:none}
  .row .k{color:var(--muted)}
  .row .v{font-weight:500;font-variant-numeric:tabular-nums}
  .dot{display:inline-block;width:8px;height:8px;border-radius:50%;margin-right:6px;flex-shrink:0}
  .up{background:var(--green);box-shadow:0 0 8px #22c55e88}
  .down{background:var(--red)}
  .btn{background:var(--accent);color:#fff;border:none;padding:6px 14px;border-radius:6px;cursor:pointer;font-size:0.8rem;font-weight:500}
  .btn:hover{opacity:.85}
  .btn.danger{background:var(--red)}
  .btn.sm{padding:4px 10px;font-size:0.75rem}
  .refresh-badge{color:var(--muted);font-size:0.75rem}
  .log-section{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:18px}
  .log-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
  .log-title{font-size:0.7rem;text-transform:uppercase;letter-spacing:0.08em;color:var(--muted);font-weight:600}
  .log-body{background:#0a0c13;border-radius:6px;height:340px;overflow-y:auto;padding:10px 12px;font-family:'Cascadia Code','Consolas',monospace;font-size:0.76rem;line-height:1.65}
  .line{white-space:pre-wrap;word-break:break-all}
  .line.err{color:#f87171}
  .line.warn{color:#fbbf24}
  .line.svc{color:#818cf8}
  .line.ok{color:#4ade80}
  .divider{width:1px;background:var(--border);margin:0 8px;align-self:stretch}
</style>
</head>
<body>
<div class="header">
  <div>
    <h1>&#x2728; Headroom Dashboard</h1>
    <div class="subtitle">Proxy: http://127.0.0.1:8787 &mdash; Dashboard: http://127.0.0.1:8788</div>
  </div>
  <div class="header-actions">
    <span class="refresh-badge" id="ts">Loading...</span>
    <button class="btn danger sm" onclick="clearCache()">Clear Cache</button>
  </div>
</div>

<div class="banner" id="banner">
  <div class="banner-title">&#x1F4BE; Lifetime Savings</div>
  <div class="bstat"><div class="val" id="lt-tokens">—</div><div class="lbl">Tokens Saved</div></div>
  <div class="bstat"><div class="val" id="lt-usd">—</div><div class="lbl">Cost Saved</div></div>
  <div class="bstat"><div class="val" id="sess-tokens">—</div><div class="lbl">Session Tokens Saved</div><div class="sub" id="sess-pct"></div></div>
  <div class="bstat"><div class="val" id="sess-usd">—</div><div class="lbl">Session Cost Saved</div></div>
</div>

<div class="grid">
  <div class="card">
    <div class="card-title">Status</div>
    <div id="health-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
  <div class="card">
    <div class="card-title">Compression (Session)</div>
    <div id="comp-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
  <div class="card">
    <div class="card-title">Requests</div>
    <div id="req-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
  <div class="card">
    <div class="card-title">Prefix Cache</div>
    <div id="pcache-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
  <div class="card">
    <div class="card-title">TOIN Intelligence</div>
    <div id="toin-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
  <div class="card">
    <div class="card-title">Per-Model (Session)</div>
    <div id="model-body"><span style="color:var(--muted)">Loading...</span></div>
  </div>
</div>

<div class="log-section">
  <div class="log-header">
    <span class="log-title">Proxy Console Output</span>
    <span class="refresh-badge" id="log-count"></span>
  </div>
  <div class="log-body" id="log-el"></div>
</div>

<script>
async function api(path, opts) {
  const r = await fetch('/api/' + path.replace(/^\//,''), opts);
  if (!r.ok) throw new Error(r.status);
  return r.json();
}

const fmtNum = n => (n == null ? '—' : Number(n).toLocaleString());
const fmtUsd = n => (n == null ? '—' : '$' + Number(n).toLocaleString(undefined, {minimumFractionDigits:2,maximumFractionDigits:2}));
const fmtPct = n => (n == null ? '—' : Number(n).toFixed(1) + '%');
const fmtMs  = n => (n == null ? '—' : Number(n).toFixed(0) + ' ms');
const esc = s => String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const row = (k, v) => `<div class="row"><span class="k">${k}</span><span class="v">${esc(v)}</span></div>`;

async function refresh() {
  const [health, stats] = await Promise.allSettled([
    api('/health'), api('/stats')
  ]);

  // --- Status card ---
  const hEl = document.getElementById('health-body');
  if (health.status === 'fulfilled') {
    const h = health.value;
    hEl.innerHTML = `
      <div style="display:flex;align-items:center;margin-bottom:12px">
        <span class="dot up"></span><strong>Online</strong>
      </div>
      <div class="rows">
        ${row('Version', h.version)}
        ${row('Optimisation', h.config?.optimize ? 'Enabled' : 'Disabled')}
        ${row('Cache', h.config?.cache ? 'Enabled' : 'Disabled')}
        ${row('Rate Limiting', h.config?.rate_limit ? 'Enabled' : 'Disabled')}
      </div>`;
  } else {
    hEl.innerHTML = '<div style="display:flex;align-items:center"><span class="dot down"></span><span style="color:var(--red)">Proxy offline</span></div>';
  }

  if (stats.status !== 'fulfilled') {
    ['comp-body','req-body','pcache-body','toin-body','model-body'].forEach(id =>
      document.getElementById(id).innerHTML = '<span style="color:var(--muted)">Unavailable</span>');
    document.getElementById('ts').textContent = 'Updated ' + new Date().toLocaleTimeString();
    refreshLogs();
    return;
  }

  const s = stats.value;

  // --- Lifetime banner ---
  const lt = s.persistent_savings?.lifetime;
  const sess = s.tokens;
  const cost = s.cost;
  if (lt) {
    document.getElementById('lt-tokens').textContent = fmtNum(lt.tokens_saved);
    document.getElementById('lt-usd').textContent    = fmtUsd(lt.compression_savings_usd);
  }
  if (sess) {
    document.getElementById('sess-tokens').textContent = fmtNum(sess.saved);
    document.getElementById('sess-pct').textContent    = sess.savings_percent != null ? fmtPct(sess.savings_percent) + ' reduction' : '';
  }
  document.getElementById('sess-usd').textContent = fmtUsd(cost?.savings_usd);

  // --- Compression card ---
  document.getElementById('comp-body').innerHTML = `
    <div class="big">${fmtPct(s.tokens?.savings_percent)}</div>
    <div class="big-label">avg token reduction this session</div>
    <div class="rows">
      ${row('Input tokens', fmtNum(s.tokens?.input))}
      ${row('Tokens saved', fmtNum(s.tokens?.saved))}
      ${row('Compressions run', fmtNum(s.summary?.compression?.requests_compressed))}
      ${row('CLI tokens avoided', fmtNum(s.tokens?.cli_tokens_avoided))}
    </div>`;

  // --- Requests card ---
  const req = s.requests || {};
  document.getElementById('req-body').innerHTML = `
    <div class="big">${fmtNum(req.total)}</div>
    <div class="big-label">total requests</div>
    <div class="rows">
      ${row('Failed', fmtNum(req.failed))}
      ${row('Rate limited', fmtNum(req.rate_limited))}
      ${row('Avg latency', fmtMs(s.latency?.average_ms))}
      ${row('Avg overhead', fmtMs(s.overhead?.average_ms))}
      ${row('Avg TTFB', fmtMs(s.ttfb?.average_ms))}
    </div>`;

  // --- Prefix cache card ---
  const pc = s.prefix_cache?.totals || {};
  document.getElementById('pcache-body').innerHTML = `
    <div class="big">${fmtPct(pc.hit_rate)}</div>
    <div class="big-label">cache hit rate</div>
    <div class="rows">
      ${row('Hit requests', fmtNum(pc.hit_requests) + ' / ' + fmtNum(pc.requests))}
      ${row('Tokens read', fmtNum(pc.cache_read_tokens))}
      ${row('Tokens written', fmtNum(pc.cache_write_tokens))}
      ${row('Net savings', fmtUsd(pc.net_savings_usd))}
      ${row('Cache busts', fmtNum(pc.bust_count))}
    </div>`;

  // --- TOIN card ---
  const toin = s.toin || {};
  document.getElementById('toin-body').innerHTML = `
    <div class="big">${fmtNum(toin.patterns_tracked)}</div>
    <div class="big-label">patterns tracked</div>
    <div class="rows">
      ${row('Compressions', fmtNum(toin.total_compressions))}
      ${row('Retrievals', fmtNum(toin.total_retrievals))}
      ${row('Retrieval rate', fmtPct(toin.global_retrieval_rate))}
      ${row('With recommendations', fmtNum(toin.patterns_with_recommendations))}
    </div>`;

  // --- Per-model card ---
  const perModel = cost?.per_model || {};
  const modelRows = Object.entries(perModel).map(([m, v]) =>
    `<div style="margin-bottom:10px">
      <div style="font-size:0.78rem;font-weight:600;margin-bottom:4px;color:var(--text)">${esc(m.replace('claude-',''))}</div>
      <div class="rows">
        ${row('Requests', fmtNum(v.requests))}
        ${row('Tokens saved', fmtNum(v.tokens_saved))}
        ${row('Reduction', fmtPct(v.reduction_pct))}
      </div>
    </div>`
  ).join('');
  document.getElementById('model-body').innerHTML = modelRows || '<span style="color:var(--muted)">No data</span>';

  document.getElementById('ts').textContent = 'Updated ' + new Date().toLocaleTimeString();
  refreshLogs();
}

async function refreshLogs() {
  try {
    const lines = await api('/log');
    const el = document.getElementById('log-el');
    const atBottom = el.scrollHeight - el.scrollTop <= el.clientHeight + 60;
    el.innerHTML = lines.map(l => {
      let cls = 'line';
      const ll = l.toLowerCase();
      if (/error|fail|exception|traceback/.test(ll)) cls += ' err';
      else if (/warn/.test(ll)) cls += ' warn';
      else if (/\[service\]/.test(l)) cls += ' svc';
      else if (/started|ready|listening|online/.test(ll)) cls += ' ok';
      return `<div class="${cls}">${esc(l)}</div>`;
    }).join('');
    document.getElementById('log-count').textContent = lines.length + ' lines';
    if (atBottom) el.scrollTop = el.scrollHeight;
  } catch {}
}

async function clearCache() {
  try {
    await api('/cache/clear', {method:'POST'});
    alert('Cache cleared.');
    refresh();
  } catch { alert('Failed to clear cache.'); }
}

refresh();
setInterval(refresh, 30000);
</script>
</body>
</html>
"""


class DashboardHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args) -> None:
        pass  # suppress server access logs from console

    def do_GET(self) -> None:
        if self.path in ("/", "/index.html"):
            self._html()
        elif self.path.startswith("/api/"):
            self._proxy_get(self.path[5:])
        else:
            self.send_error(404)

    def do_POST(self) -> None:
        if self.path.startswith("/api/"):
            self._proxy_post(self.path[5:])
        else:
            self.send_error(404)

    def _html(self) -> None:
        data = DASHBOARD_HTML.encode()
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _proxy_get(self, path: str) -> None:
        if path == "log":
            data = json.dumps(_get_logs()).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return
        url = f"http://127.0.0.1:{HEADROOM_PORT}/{path}"
        try:
            with urlopen(url, timeout=5) as r:
                data = r.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except URLError as e:
            self.send_error(502, str(e))

    def _proxy_post(self, path: str) -> None:
        url = f"http://127.0.0.1:{HEADROOM_PORT}/{path}"
        try:
            req = Request(url, method="POST", data=b"")
            with urlopen(req, timeout=5) as r:
                data = r.read()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except URLError as e:
            self.send_error(502, str(e))


def run_dashboard(stop_event: threading.Event) -> None:
    server = HTTPServer(("127.0.0.1", DASHBOARD_PORT), DashboardHandler)
    server.timeout = 1
    _add_log(f"[service] Dashboard at http://127.0.0.1:{DASHBOARD_PORT}")
    while not stop_event.is_set():
        server.handle_request()
    server.server_close()


# ---------------------------------------------------------------------------
# Windows service class
# ---------------------------------------------------------------------------

if HAS_WIN32:
    class HeadroomService(win32serviceutil.ServiceFramework):
        _svc_name_ = SERVICE_NAME
        _svc_display_name_ = SERVICE_DISPLAY
        _svc_description_ = SERVICE_DESC

        def __init__(self, args):
            win32serviceutil.ServiceFramework.__init__(self, args)
            self._wait_event = win32event.CreateEvent(None, 0, 0, None)
            self._stop_event = threading.Event()
            self._manager: HeadroomManager | None = None

        def SvcStop(self):
            self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
            self._stop_event.set()
            if self._manager:
                self._manager.stop()
            win32event.SetEvent(self._wait_event)

        def SvcDoRun(self):
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, ""),
            )
            self._manager = HeadroomManager(self._stop_event)
            self._manager.start()
            dash = threading.Thread(
                target=run_dashboard, args=(self._stop_event,), daemon=True
            )
            dash.start()
            win32event.WaitForSingleObject(self._wait_event, win32event.INFINITE)


# ---------------------------------------------------------------------------
# Debug (foreground) mode
# ---------------------------------------------------------------------------

def run_debug() -> None:
    print(f"Headroom Service [debug mode]")
    print(f"  Proxy port : {HEADROOM_PORT}")
    print(f"  Dashboard  : http://127.0.0.1:{DASHBOARD_PORT}")
    print(f"  Ctrl+C to stop\n")
    stop_event = threading.Event()
    manager = HeadroomManager(stop_event)
    manager.start()
    dash = threading.Thread(target=run_dashboard, args=(stop_event,), daemon=True)
    dash.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping...")
        stop_event.set()
        manager.stop()


def print_status() -> None:
    if not HAS_WIN32:
        print("pywin32 not available")
        return
    try:
        status = win32serviceutil.QueryServiceStatus(SERVICE_NAME)
        states = {
            win32service.SERVICE_STOPPED: "STOPPED",
            win32service.SERVICE_START_PENDING: "START PENDING",
            win32service.SERVICE_STOP_PENDING: "STOP PENDING",
            win32service.SERVICE_RUNNING: "RUNNING",
            win32service.SERVICE_CONTINUE_PENDING: "CONTINUE PENDING",
            win32service.SERVICE_PAUSE_PENDING: "PAUSE PENDING",
            win32service.SERVICE_PAUSED: "PAUSED",
        }
        state = states.get(status[1], f"UNKNOWN ({status[1]})")
        print(f"{SERVICE_NAME}: {state}")
    except Exception as e:
        print(f"Service not found or error: {e}")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    args = sys.argv[1:]

    if args and args[0] == "debug":
        run_debug()
    elif args and args[0] == "status":
        print_status()
    elif HAS_WIN32:
        win32serviceutil.HandleCommandLine(HeadroomService)
    else:
        print("pywin32 is not installed. Install it first:")
        print("  pip install pywin32")
        print()
        print("To test without the service, run:")
        print("  python headroom_service.py debug")
        sys.exit(1)
