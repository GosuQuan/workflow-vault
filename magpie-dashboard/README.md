# Magpie Token 用量看板（本地）

监控本地模型的 token 消耗，无需任何第三方依赖（纯 Python 标准库）。

## 组成
- `server.py` — 看板+记账服务（端口 **8788**），数据存 SQLite（`usage.db`）
- `proxy.py` — 统计代理（端口 **8789**）：转发请求到本地模型后端，自动记录 usage
- `dashboard.html` — 看板页面（总览卡片、每日用量图、按模型汇总、最近请求）
- `start.bat` — 一键启动看板并打开浏览器

## 用法

1. 启动看板：双击 `start.bat`（或 `python server.py`）
2. 启动统计代理，指向你的本地模型后端（Ollama=11434，vLLM=8000，LM Studio=1234）：
   ```
   python proxy.py http://127.0.0.1:11434/v1
   ```
3. 把 Magpie 的 API base 指向代理：
   ```
   OPENAI_API_BASE=http://127.0.0.1:8789/v1
   ```
   Magpie 正常调用即可，token 用量自动上看板。

## 说明
- 非 JSON / 流式响应无法取到 usage 时，按字符数估算（模型名带「(估)」标记）
- 也可不经代理，直接上报：`POST http://127.0.0.1:8788/api/events`，
  body: `{"model":"x","prompt_tokens":10,"completion_tokens":5}`
- 数据都在本地 `usage.db`，删掉即清零
