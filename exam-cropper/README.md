# exam-cropper · 试卷节选裁剪

把渲染成 PNG 的试卷页面按坐标裁出题目 / 评分标准节选，用作自编课件的配图。

> 注意：裁出的真题图仅供个人教学使用，请勿再分发（本仓库只收工具，不收素材）。

## 用法

1. 先把 PDF 页面渲染成 PNG（任意渲染器，例如 PyMuPDF 按 144 dpi 导出）。
2. 在脚本里按 144 dpi 坐标写裁剪调用：`crop(源图, (left, top, right, bottom), 输出路径)`。脚本用 `S = 1870/1224` 把 144 dpi 坐标自动缩放到目标分辨率。
3. 运行 `python crop_exam_excerpts.py`，输出图与尺寸会打印出来。

坐标怎么定：先用看图软件在 144 dpi 版面上读出大致框选范围，跑一次看输出尺寸，微调一两次即可。

批量裁剪时，每个渲染批次写一个批处理脚本（复用同样的 `crop()` 帮手与缩放因子），每组调用前用注释标明用途与所属单元，例如：

```python
# Calc 2024: Q1 (p3, Riemann/accumulation -> Unit 8), Q2 (p4, parametric motion -> Unit 9)
crop("assets/render/calchi2/page-0001.png", (55, 15, 1165, 950), "assets/crops/calc_q1.png")
```

## 依赖

```bash
pip install Pillow
```
