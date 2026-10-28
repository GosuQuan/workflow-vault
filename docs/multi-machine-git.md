# 多机 Git 同步约定

> 源自课程工作区（AGENTS.md 与 .gitignore）跑了一段时间的实践。

## 原则

- 所有电脑对同一个仓库 `git pull` / `git push` 完成多机同步。
- 密钥、凭据**绝不入库**；大型二进制（PDF/图片/zip）**不入库**，转存网盘或用 Git LFS。

## .gitignore 策略

- 通用段：OS（Thumbs.db/.DS_Store）、编辑器（.idea/.vscode）、环境（.env/.venv/__pycache__/node_modules）、构建产物（dist/build）。
- 业务段：按目录屏蔽大二进制（如 `某课程/**/*.pdf`），文档（md/docx/pptx）与索引正常提交。
- 生成产物整目录屏蔽（如渲染页与裁剪图所在的 `assets/`）。

## 提交习惯

- 中文一行式提交信息，带 `init:` / `feat:` / `fix:` / `docs:` 前缀，一眼能看出这批提交干了什么。
- 网络受限环境给仓库配代理：`git config http.proxy http://127.0.0.1:7897`。

## 新机器接入

1. `git clone <仓库>`；
2. 按需从网盘把被屏蔽的大文件放回对应目录；
3. 大文件从不入库，所以多机之间永远不会因它们冲突。
