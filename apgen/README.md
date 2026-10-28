# apgen · AP 双语课程材料生成框架

统一版式地批量生成 AP 课程的双语材料：PPTX 课件、DOCX 讲义/作业/综合测试、教学大纲。
公式不用 LaTeX，用 `^{...}` / `_{...}` 标记上下标，渲染成真实排版（PPTX 的 rPr baseline、DOCX 的 font.superscript/subscript），中文走 eastAsia 字体，不乱码。

## 组成

- `framework.py` — 生成函数库：三门课（calc / phys / em）的主题配色、考试结构与分数分布数据（EXAM_INFO）、课件排版、`build_syllabus()` 教学大纲生成。
- `validate_unit.py` — 内容文件校验器：字段完整性、公式标记合法性、可疑字符与 LaTeX 检查、数量约束（每小节 ≥4 条概念、≥2 道例题且例题 ≥3 步解析、课后练习恰 5 题且必填评分步骤、复习卷 5 选择 + 2 大题 + 1 挑战）。
- `content/calc_u01.py` — 示例内容文件（AP Calculus BC Unit 1 极限与连续）。

## 用法

1. 照 `content/calc_u01.py` 的 `UNIT` 字典结构写单元内容（所有字符串遵守标记规范）。
2. 校验：`python validate_unit.py content/calc_u01.py` → 输出 `OK calc U01 Limits and Continuity`。
3. 用 `framework.py` 的函数把 UNIT 渲染成 PPTX 课件 / DOCX 讲义 / 综合测试 / 教学大纲（在你的 build 脚本里 import 使用）。

## 依赖

```bash
pip install python-pptx python-docx
```

## 设计要点

- **内容与版式分离**：内容是纯 Python 字典，校验器先卡住格式，再渲染，杜绝"渲染完才发现格式崩"。
- 每份材料自动追加考试结构与分数分布页（数据在 `EXAM_INFO`，按年份更新即可）。
- 中英双语一一对应：所有文本字段都是 `(en, zh)` 二元组，同页并排展示。
