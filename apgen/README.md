# apgen · AP 双语课程材料生成框架

统一版式地批量生成 AP 课程的双语材料：PPTX 课件、DOCX 讲义/作业/综合测试、教学大纲。
公式不用 LaTeX，用 `^{...}` / `_{...}` 标记上下标，渲染成真实排版（PPTX 的 rPr baseline、DOCX 的 font.superscript/subscript），中文走 eastAsia 字体，不乱码。

## 组成

- `framework.py` — 生成函数库：三门课（calc / phys / em）主题配色、考试结构与分数分布数据（EXAM_INFO）、课件排版、`build_syllabus()` 教学大纲生成。标记解析支持**嵌套**：`e^{5x^{2}}`、`∫_{0}^{x^{2}}` 都能正确拍平渲染。
- `validate_unit.py` — 内容文件校验器：字段完整性、公式标记合法性、可疑字符与 LaTeX 检查、数量约束（每小节 ≥4 条概念、≥2 道例题且例题 ≥3 步解析、课后练习恰 5 题且必填评分步骤、复习卷 5 选择 + 2 大题 + 1 挑战）。
- `build_all.py` — 批量流水线：扫描 content/ 下所有单元文件，**先校验后构建**，一次产出每个单元的课件 + 课后练习（题目/答案分册）+ 章节测试（卷/答案）。`python build_all.py calc` 只构建单门课，校验不过的单元自动跳过并在结尾汇总报错（非零码退出）。
- `build_syllabi.py` — 示例：汇总各单元元数据，调用 `build_syllabus()` 生成整年教学大纲（课程简介文案按自己的课程替换）。
- `verify_all.py` — 生成物回归校验：所有 PPTX 能打开、页数/图片数统计、检查 `^{` `_{` `\frac` 等标记是否泄漏进成稿（MARK-LEAK），任何异常非零码退出，适合接 CI。
- `content/calc_u01.py` — 示例内容文件（AP Calculus BC Unit 1 极限与连续）。

## 用法

```bash
# 1. 照 content/calc_u01.py 的 UNIT 字典写内容，先校验
python validate_unit.py content/calc_u01.py

# 2. 批量构建
python build_all.py            # 全部课程
python build_all.py calc       # 只构建 calc

# 3. 生成教学大纲（编辑 build_syllabi.py 里的课程简介文案）
python build_syllabi.py

# 4. 对生成物做回归校验（在课程根目录运行，glob 扫描课件目录）
python verify_all.py
```

注意：`build_all.py` 的输出目录映射（`COURSE_FOLDER`）、`build_syllabi.py` 的输出路径与 `verify_all.py` 的扫描目录都按自己的课程目录调整；`framework.py` 本身不依赖目录结构。

## 依赖

```bash
pip install python-pptx python-docx
```

## 设计要点

- **内容与版式分离**：内容是纯 Python 字典，校验器先卡住格式，再渲染，杜绝"渲染完才发现格式崩"。
- **构建 → 回归闭环**：构建前有 validate（内容层），构建后有 verify_all（成稿层，抓标记泄漏），双层防线。
- 每份材料自动追加考试结构与分数分布页（数据在 `EXAM_INFO`，按年份更新即可）。
- 中英双语一一对应：所有文本字段都是 `(en, zh)` 二元组，同页并排展示。
