# workflow-vault · 个人工作流沉淀库

把日常课程工作区里"跑通且反复在用"的工具与流程，沉淀成一个开源、可复用的仓库。
每天早上由 agent 扫描工作区增量，做脱敏与版权检查后归档到这里（流程见 [docs/daily-distill.md](docs/daily-distill.md)）。

## 目录

| 目录 | 是什么 | 依赖 |
| --- | --- | --- |
| [apgen/](apgen/) | AP 双语课程材料生成框架：PPTX 课件 + DOCX 讲义/大纲，公式用 `^{...}` / `_{...}` 标记上下标（支持嵌套），含校验→批量构建→回归校验流水线 | python-pptx, python-docx |
| [exam-cropper/](exam-cropper/) | 把渲染好的试卷页面按坐标裁出节选图（仅个人使用，不再分发） | Pillow |
| [deck-verifier/](deck-verifier/) | 生成后自动校验课件：页数、图片数、关键文本是否在场 | python-pptx |
| [md-to-docx/](md-to-docx/) | 中文 Markdown → 打印友好 DOCX（A4、中文字体、表格、等宽块），一条命令 | python-docx |
| [magpie-dashboard/](magpie-dashboard/) | 本地 LLM token 用量看板：统计代理 + SQLite + 网页，纯标准库零依赖 | 无 |
| [docs/](docs/) | 多机 Git 同步约定、每日沉淀 SOP | - |

## 快速上手

```bash
# 1. 校验一个单元内容文件（内容与版式分离，先卡格式再渲染）
python apgen/validate_unit.py apgen/content/calc_u01.py

# 2. 批量构建全部单元材料，并对生成物做回归校验
python apgen/build_all.py && python apgen/verify_all.py

# 3. 中文 Markdown 一键转 DOCX
python md-to-docx/md_to_docx.py 手册.md

# 4. 启动本地 LLM token 用量看板（零依赖）
cd magpie-dashboard && python server.py   # http://127.0.0.1:8788
```

## 约定

- **不收录**：密钥/凭据、受版权保护的第三方内容（真题原文、截图、曲谱等）、个人数据、大二进制（素材走网盘）。
- 工具尽量零依赖或最少依赖；路径走参数/argv，不硬编码个人目录。
- 提交信息：中文一行，`init:` / `feat:` / `fix:` / `docs:` 前缀。

## License

[MIT](LICENSE)
