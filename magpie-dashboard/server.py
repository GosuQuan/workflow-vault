# -*- coding: utf-8 -*-
"""Magpie 本地模型 Token 用量看板服务（纯标准库，端口 8788）"""
import json, sqlite3, threading
from datetime import datetime, timedelta
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

ROOT = Path(__file__).parent
DB = ROOT / "usage.db"
DASH = ROOT / "dashboard.html"
LOCK = threading.Lock()

def db():
    c = sqlite3.connect(DB)
    c.execute("""CREATE TABLE IF NOT EXISTS events(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ts TEXT NOT NULL,
        model TEXT NOT NULL DEFAULT '',
        prompt_tokens INTEGER NOT NULL DEFAULT 0,
        completion_tokens INTEGER NOT NULL DEFAULT 0
    )""")
    return c

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _json(self, obj, code=200):
        body = json.dumps(obj).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/" or self.path.startswith("/index"):
            body = DASH.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        elif self.path.startswith("/api/ping"):
            self._json({"ok": True})
        elif self.path.startswith("/api/summary"):
            days = 30
            if "days=" in self.path:
                try: days = int(self.path.split("days=")[1].split("&")[0])
                except ValueError: pass
            since = (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
            with LOCK, db() as c:
                tot = c.execute("SELECT COALESCE(SUM(prompt_tokens),0),COALESCE(SUM(completion_tokens),0),COUNT(*) FROM events WHERE ts>=?", (since,)).fetchone()
                daily = c.execute("SELECT substr(ts,1,10) d,SUM(prompt_tokens),SUM(completion_tokens),COUNT(*) FROM events WHERE ts>=? GROUP BY d ORDER BY d", (since,)).fetchall()
                models = c.execute("SELECT model,SUM(prompt_tokens),SUM(completion_tokens),COUNT(*) FROM events WHERE ts>=? GROUP BY model ORDER BY 2+3 DESC", (since,)).fetchall()
                recent = c.execute("SELECT ts,model,prompt_tokens,completion_tokens FROM events ORDER BY id DESC LIMIT 50").fetchall()
            self._json({
                "days": days,
                "totals": {"prompt_tokens": tot[0], "completion_tokens": tot[1], "requests": tot[2]},
                "daily": [{"date": d[0], "prompt": d[1], "completion": d[2], "requests": d[3]} for d in daily],
                "models": [{"model": m[0], "prompt": m[1], "completion": m[2], "requests": m[3]} for m in models],
                "recent": [{"ts": r[0], "model": r[1], "prompt": r[2], "completion": r[3]} for r in recent],
            })
        else:
            self._json({"error": "not found"}, 404)

    def do_POST(self):
        if not self.path.startswith("/api/events"):
            self._json({"error": "not found"}, 404); return
        n = int(self.headers.get("Content-Length") or 0)
        try:
            data = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            self._json({"error": "bad json"}, 400); return
        ev = {
            "ts": data.get("ts") or datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "model": str(data.get("model") or "unknown"),
            "p": int(data.get("prompt_tokens") or 0),
            "c": int(data.get("completion_tokens") or 0),
        }
        with LOCK, db() as c:
            c.execute("INSERT INTO events(ts,model,prompt_tokens,completion_tokens) VALUES(?,?,?,?)",
                      (ev["ts"], ev["model"], ev["p"], ev["c"]))
        self._json({"ok": True})

if __name__ == "__main__":
    print("dashboard: http://127.0.0.1:8788")
    HTTPServer(("127.0.0.1", 8788), Handler).serve_forever()
