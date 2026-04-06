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
    <span class="log-title">Token Usage per API Call</span>
    <span class="refresh-badge" id="graph-call-count"></span>
    <button onclick="clearGraph()" style="margin-left:auto;font-size:0.7rem;padding:2px 8px;background:var(--card);border:1px solid var(--border);color:var(--muted);border-radius:4px;cursor:pointer">Clear</button>
  </div>
  <div style="position:relative;height:220px">
    <canvas id="token-canvas"></canvas>
  </div>
  <div id="turn-summary" style="margin-top:10px;font-size:0.74rem;line-height:1.8;min-height:1.2em"></div>
  <div style="margin-top:6px;font-size:0.7rem;color:var(--muted)">
    Each bar = one API call. Bars of the same colour belong to the same user turn (prompt).
    Input tokens (solid) + output tokens (faded) stacked. Polling every 2 s.
  </div>
</div>

<div class="log-section" style="margin-top:14px">
  <div class="log-header">
    <span class="log-title">Proxy Console Output</span>
    <span class="refresh-badge" id="log-count"></span>
  </div>
  <div class="log-body" id="log-el"></div>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
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

// ---- Token Usage Graph ----
const TURN_GAP_MS = 12000; // gap between API calls that marks a new user turn
const TURN_PALETTE = [
  '#6366f1','#22c55e','#f59e0b','#06b6d4',
  '#a855f7','#ec4899','#f97316','#84cc16','#ef4444','#14b8a6'
];
let callHistory = [];
let prevInput = null, prevOutput = null;
let lastCallTime = 0, turnNum = 0;
let tokenChart = null;

function saveGraphState() {
  try {
    localStorage.setItem('hrTokenHistory', JSON.stringify({ callHistory, turnNum, lastCallTime }));
  } catch {}
}

function loadGraphState() {
  try {
    const saved = JSON.parse(localStorage.getItem('hrTokenHistory') || 'null');
    if (saved && Array.isArray(saved.callHistory)) {
      callHistory  = saved.callHistory;
      turnNum      = saved.turnNum      || 0;
      lastCallTime = saved.lastCallTime || 0;
    }
  } catch {}
}

function initTokenChart() {
  const ctx = document.getElementById('token-canvas').getContext('2d');
  tokenChart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: [],
      datasets: [
        { label: 'Input tokens',  data: [], backgroundColor: [], borderRadius: 3, stack: 'tok' },
        { label: 'Output tokens', data: [], backgroundColor: [], borderRadius: 3, stack: 'tok' }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: false,
      plugins: {
        legend: { labels: { color: '#e2e8f0', font: { size: 11 }, boxWidth: 12 } },
        tooltip: {
          callbacks: {
            title(items) {
              const c = callHistory[items[0].dataIndex];
              return `Call #${c.seq} \u00b7 Turn ${c.turn} \u00b7 ${c.ts}`;
            },
            label(item) {
              return `${item.dataset.label}: ${Number(item.raw).toLocaleString()}`;
            },
            afterBody(items) {
              const c = callHistory[items[0].dataIndex];
              return [`Total: ${Number(c.inputDelta + c.outputDelta).toLocaleString()}`];
            }
          }
        }
      },
      scales: {
        x: {
          stacked: true,
          ticks: { color: '#94a3b8', font: { size: 10 } },
          grid: { color: '#2d3148' },
          title: { display: true, text: 'API call (chronological)', color: '#94a3b8', font: { size: 10 } }
        },
        y: {
          stacked: true,
          ticks: { color: '#94a3b8', font: { size: 10 } },
          grid: { color: '#2d3148' },
          title: { display: true, text: 'Tokens', color: '#94a3b8', font: { size: 10 } }
        }
      }
    }
  });
}

function renderTokenChart() {
  if (!tokenChart) return;
  const inBg = [], outBg = [], labels = [];
  callHistory.forEach(c => {
    const col = TURN_PALETTE[(c.turn - 1) % TURN_PALETTE.length];
    inBg.push(col + 'dd');
    outBg.push(col + '55');
    labels.push('#' + c.seq);
  });
  tokenChart.data.labels = labels;
  tokenChart.data.datasets[0].data = callHistory.map(c => c.inputDelta);
  tokenChart.data.datasets[0].backgroundColor = inBg;
  tokenChart.data.datasets[1].data = callHistory.map(c => c.outputDelta);
  tokenChart.data.datasets[1].backgroundColor = outBg;
  tokenChart.update('none');

  // Turn summary row
  const turns = {};
  callHistory.forEach(c => {
    if (!turns[c.turn]) turns[c.turn] = { calls: 0, tokens: 0 };
    turns[c.turn].calls++;
    turns[c.turn].tokens += c.inputDelta + c.outputDelta;
  });
  document.getElementById('turn-summary').innerHTML =
    Object.entries(turns).map(([t, v]) => {
      const col = TURN_PALETTE[(t - 1) % TURN_PALETTE.length];
      return `<span style="color:${col};margin-right:16px">\u25a0 Turn ${t}: ` +
             `${v.calls} call${v.calls > 1 ? 's' : ''}, ` +
             `${Number(v.tokens).toLocaleString()} tokens</span>`;
    }).join('');

  document.getElementById('graph-call-count').textContent =
    callHistory.length + ' call' + (callHistory.length !== 1 ? 's' : '') +
    ' \u00b7 ' + turnNum + ' turn' + (turnNum !== 1 ? 's' : '');
}

async function pollTokenMetrics() {
  try {
    const text = await fetch('/api/metrics').then(r => r.text());
    const inM  = text.match(/^headroom_tokens_input_total\s+([\d]+)/m);
    const outM = text.match(/^headroom_tokens_output_total\s+([\d]+)/m);
    if (!inM || !outM) return;
    const curIn  = parseInt(inM[1],  10);
    const curOut = parseInt(outM[1], 10);
    if (prevInput === null) { prevInput = curIn; prevOutput = curOut; return; }
    const dIn  = curIn  - prevInput;
    const dOut = curOut - prevOutput;
    if (dIn > 0 || dOut > 0) {
      const now = Date.now();
      if (now - lastCallTime > TURN_GAP_MS) turnNum++;
      lastCallTime = now;
      callHistory.push({
        seq: callHistory.length + 1,
        inputDelta:  Math.max(0, dIn),
        outputDelta: Math.max(0, dOut),
        ts:   new Date().toLocaleTimeString(),
        turn: turnNum
      });
      saveGraphState();
      renderTokenChart();
    }
    prevInput  = curIn;
    prevOutput = curOut;
  } catch {}
}

function clearGraph() {
  callHistory = []; turnNum = 0; lastCallTime = 0;
  saveGraphState();
  renderTokenChart();
}

loadGraphState();
initTokenChart();
renderTokenChart();
pollTokenMetrics();
setInterval(pollTokenMetrics, 2000);
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
# Task Scheduler management (works with Microsoft Store Python)
# ---------------------------------------------------------------------------

TASK_NAME = SERVICE_NAME


def _pythonw() -> str:
    """Return path to pythonw.exe next to the current python.exe."""
    pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if os.path.exists(pythonw):
        return pythonw
    # Store Python puts executables in a different location; fall back to python
    return sys.executable


def _script() -> str:
    return os.path.abspath(__file__)


def task_install() -> None:
    pythonw = _pythonw()
    script = _script()
    cmd = f'"{pythonw}" "{script}" debug'
    # Build schtasks XML so we can set restart-on-failure and hidden window
    xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>{SERVICE_DESC}</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <RestartOnFailure>
      <Interval>PT1M</Interval>
      <Count>999</Count>
    </RestartOnFailure>
    <Enabled>true</Enabled>
  </Settings>
  <Actions>
    <Exec>
      <Command>{pythonw}</Command>
      <Arguments>"{script}" debug</Arguments>
      <WorkingDirectory>{os.path.dirname(script)}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>"""
    xml_path = os.path.join(os.path.dirname(script), "_task.xml")
    with open(xml_path, "w", encoding="utf-16") as f:
        f.write(xml)
    try:
        result = subprocess.run(
            ["schtasks", "/Create", "/TN", TASK_NAME, "/XML", xml_path, "/F"],
            capture_output=True, text=True
        )
        if result.returncode == 0:
            print(f"Installed: {TASK_NAME}")
            print(f"  Runs at logon: {cmd}")
            print(f"  Dashboard will be at http://127.0.0.1:{DASHBOARD_PORT}")
        else:
            print("Failed to install task:")
            print(result.stderr or result.stdout)
            sys.exit(1)
    finally:
        try:
            os.remove(xml_path)
        except OSError:
            pass


def task_start() -> None:
    result = subprocess.run(
        ["schtasks", "/Run", "/TN", TASK_NAME],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(f"Started: {TASK_NAME}")
    else:
        print("Failed to start task:")
        print(result.stderr or result.stdout)
        sys.exit(1)


def task_stop() -> None:
    result = subprocess.run(
        ["schtasks", "/End", "/TN", TASK_NAME],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(f"Stopped: {TASK_NAME}")
    else:
        print("Failed to stop task:")
        print(result.stderr or result.stdout)
        sys.exit(1)


def task_remove() -> None:
    result = subprocess.run(
        ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(f"Removed: {TASK_NAME}")
    else:
        print("Failed to remove task:")
        print(result.stderr or result.stdout)
        sys.exit(1)


def task_status() -> None:
    result = subprocess.run(
        ["schtasks", "/Query", "/TN", TASK_NAME, "/FO", "LIST"],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        print(result.stdout)
    else:
        print(f"Task '{TASK_NAME}' not found.")


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


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

COMMANDS = {
    "debug":   (run_debug,    "Run in foreground (Ctrl+C to stop)"),
    "install": (task_install, "Register as a startup task"),
    "start":   (task_start,   "Start the task now"),
    "stop":    (task_stop,    "Stop the running task"),
    "remove":  (task_remove,  "Unregister the startup task"),
    "status":  (task_status,  "Show task status"),
}

if __name__ == "__main__":
    args = sys.argv[1:]
    if not args or args[0] not in COMMANDS:
        print("Usage: python headroom_service.py <command>\n")
        for name, (_, desc) in COMMANDS.items():
            print(f"  {name:<10} {desc}")
        sys.exit(0 if not args else 1)
    COMMANDS[args[0]][0]()
