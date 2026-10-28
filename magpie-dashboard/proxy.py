# -*- coding: utf-8 -*-
"""OpenAI 兼容统计代理：Magpie 指向本代理(8789)，自动转发到本地模型后端并记录 token 用量。
用法: python proxy.py [target_base]  例如 python proxy.py http://127.0.0.1:11434/v1
转发响应中的 usage 字段；流式响应按字符数/4 估算并标记 model 名加 "(估)"。"""
import json, sys, urllib.request
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler

TARGET = (sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:11434/v1").rstrip("/")
DASH = "http://127.0.0.1:8788/api/events"

def report(model, p, c):
    try:
        body = json.dumps({"model": model, "prompt_tokens": p, "completion_tokens": c}).encode()
        req = urllib.request.Request(DASH, data=body, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=3).read()
    except Exception:
        pass

class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def _forward(self):
        url = TARGET + self.path
        n = int(self.headers.get("Content-Length") or 0)
        body = self.rfile.read(n) if n else None
        headers = {k: v for k, v in self.headers.items() if k.lower() not in ("host", "content-length")}
        req = urllib.request.Request(url, data=body, headers=headers, method=self.command)
        try:
            resp = urllib.request.urlopen(req, timeout=600)
            data = resp.read()
            code, rheaders = resp.code, resp.headers
        except urllib.error.HTTPError as e:
            data, code, rheaders = e.read(), e.code, e.headers
        except Exception as e:
            msg = json.dumps({"error": str(e)}).encode()
            self.send_response(502); self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(msg))); self.end_headers(); self.wfile.write(msg); return
        ctype = rheaders.get("Content-Type", "application/json")
        model, p, c = "", 0, 0
        if "json" in ctype and (b'"usage"' in data or b"usage" in data[:200000]):
            try:
                j = json.loads(data)
                u = j.get("usage") or {}
                p, c = u.get("prompt_tokens") or 0, u.get("completion_tokens") or 0
                model = j.get("model") or ""
            except Exception:
                pass
        if code == 200 and self.path.rstrip("/").endswith(("completions", "generate")) and (p or c) == 0:
            # 流式/无 usage：粗略估算（约4字符=1 token）
            model = model or (json.loads(body).get("model", "") if body else "")
            p = max(1, len(body or b"") // 8); c = max(1, len(data) // 4)
            model = (model or "unknown") + "(估)"
        if (p or c) and code == 200:
            report(model or "unknown", p, c)
        self.send_response(code)
        for k in ("Content-Type",):
            if rheaders.get(k): self.send_header(k, rheaders[k])
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    do_GET = do_POST = do_PUT = do_DELETE = _forward

if __name__ == "__main__":
    print("proxy: http://127.0.0.1:8789  ->  " + TARGET)
    HTTPServer(("127.0.0.1", 8789), Handler).serve_forever()

