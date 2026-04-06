# HeadroomWrapper

Windows service wrapper for [Headroom](https://github.com/chopratejas/headroom) — runs the proxy on startup with a hidden console and serves a local dashboard.

## Features

- Runs `headroom proxy` as a Windows service (auto-start, no console window)
- Restarts headroom automatically if it crashes
- Dashboard at `http://127.0.0.1:8788` showing:
  - Lifetime and session token/cost savings
  - Compression stats, request counts, latency
  - Prefix cache hit rate and savings
  - TOIN intelligence patterns
  - Per-model breakdown
  - Token usage graph — tokens per API call, grouped by user turn (persisted across page refreshes)
  - Live proxy console output

## Requirements

```
pip install pywin32
```

Headroom must already be installed and on PATH (`pip install "headroom-ai[all]"`).

## Usage

### Test in foreground (no admin required)

```
python headroom_service.py debug
```

Open `http://127.0.0.1:8788` to view the dashboard.

### Install as a startup task (run as administrator)

```
python headroom_service.py install
python headroom_service.py start
```

Registers a Windows Task Scheduler task that runs at logon with no console window, and restarts automatically on failure.

### Other commands

```
python headroom_service.py stop     # Stop the running task
python headroom_service.py remove   # Unregister the startup task
python headroom_service.py status   # Show task status
```

## Ports

| Port | Purpose |
|------|---------|
| 8787 | Headroom proxy (set `ANTHROPIC_BASE_URL=http://127.0.0.1:8787`) |
| 8788 | Dashboard |
