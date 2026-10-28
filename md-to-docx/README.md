# md-to-docx · 中文 Markdown 转打印友好 DOCX

一条命令把中文 Markdown 转成 A4 打印友好的 Word 文档，纯 python-docx 实现，无 pandoc 依赖。

支持：标题层级、无序/有序列表、表格（表头底纹、自动全宽）、行内代码与链接、引用块、` ``` ` 等宽块（灰底 Consolas、空格保留）、分隔线、页脚居中页码；正文宋体 + 标题微软雅黑，附打印前自检（重新打开统计段落数与表格数）。

## 用法

```bash
python md_to_docx.py 输入.md [输出.docx]
```

省略输出路径时自动取同名 `.docx`。

## 依赖

```bash
pip install python-docx
```
