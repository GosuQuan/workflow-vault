# deck-verifier · 课件生成结果校验

生成 PPTX 后做"回归测试"：自动检查页数、图片数、关键文本（标题 / 公式 / 图表说明）是否在场，防止模板化生成悄悄崩坏。

## 用法

把 `check()` 调用里的路径与期望值换成你自己的项目：

```python
check("课件.pptx", 期望页数, 期望图片数, [必须出现的关键文本, ...])
```

运行 `python verify_decks.py`，逐项输出 `OK` / `MISS`。

这个思路适用于任何"模板化生成 Office 文档"的流水线：生成 → 校验 → 不过就修。

## 依赖

```bash
pip install python-pptx
```
