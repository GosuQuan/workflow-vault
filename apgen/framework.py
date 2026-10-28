# -*- coding: utf-8 -*-
"""AP 双语课程材料生成框架:统一版式 + 公式上下标标记 ^{...} / _{...}。

PPTX:标记解析为真实上下标(rPr baseline),不依赖花哨 Unicode。
DOCX:标记解析为 run.font.superscript / subscript,中文设 eastAsia 字体。
"""
import math
import os
import re

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from docx import Document
from docx.shared import Pt as DPt, Inches as DIn, RGBColor as DRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

EN_FONT = "Calibri"
ZH_FONT = "Microsoft YaHei"
MATH_FONT = "Cambria Math"
DOC_ZH_EA = "微软雅黑"
DOC_MATH = "Cambria Math"

MARK_RE = re.compile(r"(\^\{[^}]*\}|_\{[^}]*\})")

THEMES = {
    "calc": {"dark": RGBColor(0x1F, 0x4E, 0x79), "accent": RGBColor(0x2E, 0x75, 0xB6),
             "light": RGBColor(0xDD, 0xEB, 0xF7), "ans": RGBColor(0xC0, 0x50, 0x4D)},
    "phys": {"dark": RGBColor(0x1E, 0x6F, 0x50), "accent": RGBColor(0x2E, 0x8B, 0x6A),
             "light": RGBColor(0xE2, 0xF0, 0xE4), "ans": RGBColor(0xC0, 0x50, 0x4D)},
    "em": {"dark": RGBColor(0x5B, 0x2E, 0x7E), "accent": RGBColor(0x7B, 0x4F, 0xA6),
           "light": RGBColor(0xE9, 0xDF, 0xF2), "ans": RGBColor(0xC0, 0x50, 0x4D)},
}

COURSE_TAG = {"calc": "AP Calculus BC", "phys": "AP Physics C: Mechanics", "em": "AP Physics C: Electricity & Magnetism"}

# 每门课的考试结构与分数页数据(自动追加到每份课件末尾)
EXAM_INFO = {
    "calc": {
        "struct_title": ("Exam Structure & Calculator 考试结构与计算器政策",
                         "总时长 3 小时 15 分钟 · 每年 5 月全球统考"),
        "struct_head": ["部分 Section", "题量 Questions", "时间 Time", "计算器 Calculator"],
        "struct_rows": [
            ["Section I · Part A  选择题", "30 题", "60 分钟", "不允许 No"],
            ["Section I · Part B  选择题", "15 题", "45 分钟", "允许 Permitted"],
            ["Section II · Part A  大题 FRQ", "2 题", "30 分钟", "必须使用 Required"],
            ["Section II · Part B  大题 FRQ", "4 题", "60 分钟", "不允许 No"],
        ],
        "struct_notes": [
            ("Each FRQ is worth 9 points; MCQ 50% + FRQ 50%. Graphing calculator required in Section II Part A.",
             "每道大题 9 分;选择与大题各占 50%。Section II Part A 必须使用图形计算器,平时就用考试同型号练习。"),
        ],
        "score_title": ("Scores & Where This Fits 分数分布与 AP 体系定位", "难度、平均分与学分兑换 What the numbers say"),
        "score_head": ["年份 Year", "5 分率", "3 分以上 3+", "平均分 Mean", "考生人数"],
        "score_rows": [["2024", "45.0%", "80.0%", "3.92", "148,191"],
                        ["2025", "44.0%", "79.0%", "3.82", "160,436"]],
        "score_notes": [
            ("Among the highest 5-rates in AP (self-selected cohort), but the fastest pace of all AP math: BC = all of AB + series, parametric & polar.",
             "5 分率全 AP 最高之列(选课群体强),但 BC 是节奏最快的 AP 数学 = AB 全部内容 + 级数、参数方程、极坐标。"),
            ("AP system: global exam each May, 1–5; 3+ usually earns college credit (top schools 4–5). Calc BC ≈ college Calculus I+II.",
             "AP 体系:每年 5 月统考,1–5 分制;3 分以上多数大学可换学分(顶尖校常要求 4–5)。微积分 BC ≈ 大学微积分 I+II。"),
        ],
        "score_source": "数据 Source: College Board official score distributions (apstudents.collegeboard.org)",
    },
    "phys": {
        "struct_title": ("Exam Structure & Calculator 考试结构与计算器政策", "总时长 3 小时 · 每年 5 月全球统考"),
        "struct_head": ["部分 Section", "题量 Questions", "时间 Time", "计算器 Calculator"],
        "struct_rows": [
            ["Section I  选择题 MCQ", "40 题", "80 分钟", "全程允许(科学/图形)"],
            ["Section II  大题 FRQ", "4 题", "100 分钟", "全程允许(科学/图形)"],
        ],
        "struct_notes": [
            ("2025 format: 40 MCQ + 4 FRQ (previously 35 + 3 with reading period). Calculator allowed on the ENTIRE exam.",
             "2025 年起改革:40 道选择 + 4 道大题(原 35+3 含读卷时间);物理 C 全程允许使用计算器。"),
        ],
        "score_title": ("Scores & Where This Fits 分数分布与 AP 体系定位", "难度、平均分与学分兑换 What the numbers say"),
        "score_head": ["年份 Year", "5 分率", "3 分以上 3+", "平均分 Mean", "考生人数"],
        "score_rows": [["2024", "28.5%", "76.3%", "3.50", "61,252"],
                        ["2025 (新格式)", "21.7%", "73.2%", "3.30", "65,980"]],
        "score_notes": [
            ("One of the hardest AP sciences: calculus on every page (F = dp/dt). 5-rate dropped after the 2025 change; 3+ still ~73%.",
             "公认最难的 AP 理科之一:每题都要用微积分。2025 新格式后 5 分率下降,但 3 分以上仍约 73%。"),
            ("AP system: 1–5; 3+ usually earns credit (engineering majors 4–5). Physics C ≈ calculus-based university Physics I.",
             "AP 体系:1–5 分制,3 分以上通常可换学分(工程专业常要求 4–5)。Physics C ≈ 大学基于微积分的物理 I。"),
        ],
        "score_source": "数据 Source: College Board official score distributions (apstudents.collegeboard.org)",
    },
    "em": {
        "struct_title": ("Exam Structure & Calculator 考试结构与计算器政策", "总时长 3 小时 · 每年 5 月全球统考"),
        "struct_head": ["部分 Section", "题量 Questions", "时间 Time", "计算器 Calculator"],
        "struct_rows": [
            ["Section I  选择题 MCQ", "40 题", "80 分钟", "全程允许(科学/图形)"],
            ["Section II  大题 FRQ", "4 题", "100 分钟", "全程允许(科学/图形)"],
        ],
        "struct_notes": [
            ("E&M shares the 2025 format with Mechanics: 40 MCQ + 4 FRQ; calculator allowed throughout.",
             "电磁学与力学同用 2025 新格式:40 选择 + 4 大题;全程允许计算器。力学与电磁学可同年报考。"),
        ],
        "score_title": ("Scores & Where This Fits 分数分布与 AP 体系定位", "难度、平均分与学分兑换 What the numbers say"),
        "score_head": ["年份 Year", "5 分率", "3 分以上 3+", "平均分 Mean", "考生人数"],
        "score_rows": [["2024", "25.3%", "70.4%", "3.40", "48,375"],
                        ["2025 (新格式)", "18.0%", "68.2%", "3.20", "52,000"]],
        "score_notes": [
            ("E&M is typically scored slightly below Mechanics; charge distributions, Gauss's law and RC transients are the main gates.",
             "电磁学均分通常略低于力学:电荷分布积分、高斯定律与 RC 暂态是最主要的关口。"),
            ("AP system: 1–5; 3+ usually earns credit (engineering majors 4–5). E&M ≈ calculus-based university Physics II (E&M part).",
             "AP 体系:1–5 分制,3 分以上通常可换学分。电磁学 ≈ 大学基于微积分的物理 II(电磁部分)。"),
        ],
        "score_source": "数据 Source: College Board score distributions; 2025 近似值,以官方公布为准",
    },
}


# ---------- 标记解析 ----------

def split_marked(text):
    """'x^{n-1}+a_{1}' -> [('x', None), ('n-1','sup'), ('+a', None), ('1','sub')]"""
    parts = MARK_RE.split(str(text))
    out = []
    for p in parts:
        if not p:
            continue
        if p.startswith("^{") and p.endswith("}"):
            out.append((p[2:-1], "sup"))
        elif p.startswith("_{") and p.endswith("}"):
            out.append((p[2:-1], "sub"))
        else:
            out.append((p, None))
    return out


def _apply_baseline(run, kind):
    if kind is None:
        return
    rPr = run._r.get_or_add_rPr()
    rPr.set("baseline", "30000" if kind == "sup" else "-25000")


def add_rich_runs(p, text, size=14, bold=False, color=None, font=None, italic=False):
    for seg, kind in split_marked(text):
        r = p.add_run()
        r.text = seg
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.italic = italic
        if color is not None:
            r.font.color.rgb = color
        r.font.name = font or EN_FONT
        _apply_baseline(r, kind)
    return p


def para(tf, first=False):
    return tf.paragraphs[0] if first else tf.add_paragraph()


# ---------- PPTX 版式 ----------

W, H = Inches(13.333), Inches(7.5)


def new_deck():
    prs = Presentation()
    prs.slide_width = W
    prs.slide_height = H
    return prs


def add_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def header_bar(slide, en, zh, tag, th):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(1.0))
    bar.fill.solid(); bar.fill.fore_color.rgb = th["dark"]; bar.line.fill.background()
    tf = bar.text_frame
    tf.margin_left = Inches(0.5); tf.margin_top = Inches(0.08); tf.word_wrap = True
    p = para(tf, True)
    add_rich_runs(p, en, 24, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    p2 = tf.add_paragraph()
    r = p2.add_run(); r.text = zh; r.font.size = Pt(13)
    r.font.color.rgb = RGBColor(0xD5, 0xE5, 0xF0); r.font.name = ZH_FONT
    tbox = slide.shapes.add_textbox(Inches(9.6), Inches(0.28), Inches(3.4), Inches(0.5))
    p3 = para(tbox.text_frame, True); p3.alignment = PP_ALIGN.RIGHT
    add_rich_runs(p3, tag, 14, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))


def footer(slide, th, page, tag):
    box = slide.shapes.add_textbox(Inches(0.5), Inches(7.05), Inches(11.9), Inches(0.35))
    p = para(box.text_frame, True)
    r = p.add_run(); r.text = tag + "  |  双语课件 Bilingual Courseware"
    r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x7A, 0x84, 0x8E)
    pbox = slide.shapes.add_textbox(Inches(12.55), Inches(7.05), Inches(0.6), Inches(0.35))
    p3 = para(pbox.text_frame, True); p3.alignment = PP_ALIGN.RIGHT
    add_rich_runs(p3, str(page), 10, color=RGBColor(0x7A, 0x84, 0x8E))


def bullets(slide, items, x, y, w, h, th, en_size=16, zh_size=13):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for en, zh in items:
        p = para(tf, first); first = False
        p.space_after = Pt(4)
        add_rich_runs(p, "• ", en_size, bold=True, color=th["dark"])
        add_rich_runs(p, en, en_size)
        if zh:
            p2 = tf.add_paragraph(); p2.space_after = Pt(10)
            r = p2.add_run(); r.text = zh
            r.font.size = Pt(zh_size); r.font.color.rgb = RGBColor(0x6B, 0x72, 0x7B); r.font.name = ZH_FONT
    return box


def formula_box(slide, text, x, y, w, h, th, size=20, sub=None):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb = th["light"]
    shp.line.color.rgb = th["accent"]; shp.line.width = Pt(1.2)
    tf = shp.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = para(tf, True); p.alignment = PP_ALIGN.CENTER
    add_rich_runs(p, text, size, bold=True, color=th["dark"], font=MATH_FONT)
    if sub:
        p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
        r = p2.add_run(); r.text = sub
        r.font.size = Pt(12); r.font.color.rgb = RGBColor(0x6B, 0x72, 0x7B); r.font.name = ZH_FONT
    return shp


def make_table(slide, header, data, x, y, w, h, th, widths=None, font_size=13):
    rows, cols = len(data) + 1, len(header)
    tbl = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    if widths:
        for i, cw in enumerate(widths):
            tbl.columns[i].width = Inches(cw)
    for j, txt in enumerate(header):
        cell = tbl.cell(0, j)
        cell.fill.solid(); cell.fill.fore_color.rgb = th["dark"]
        add_rich_runs(cell.text_frame.paragraphs[0], txt, font_size, bold=True,
                      color=RGBColor(0xFF, 0xFF, 0xFF), font=ZH_FONT)
    for i, row in enumerate(data, start=1):
        for j, txt in enumerate(row):
            cell = tbl.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xF6, 0xFA, 0xFE) if i % 2 else RGBColor(0xFF, 0xFF, 0xFF)
            add_rich_runs(cell.text_frame.paragraphs[0], txt, font_size, font=MATH_FONT)
    return tbl


def example_slide(prs, th, tag, num, kind, problem_en, problem_zh, steps, answer, page):
    slide = add_slide(prs)
    header_bar(slide, f"例题 Example {num}", kind, tag, th)
    panel = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.15), Inches(12.33), Inches(1.35))
    panel.fill.solid(); panel.fill.fore_color.rgb = th["light"]; panel.line.color.rgb = th["accent"]
    tf = panel.text_frame; tf.word_wrap = True; tf.margin_left = Inches(0.25); tf.margin_top = Inches(0.08)
    add_rich_runs(para(tf, True), problem_en, 15, bold=True)
    p2 = tf.add_paragraph()
    r = p2.add_run(); r.text = problem_zh
    r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x44, 0x4C, 0x55); r.font.name = ZH_FONT
    box = slide.shapes.add_textbox(Inches(0.6), Inches(2.7), Inches(12.1), Inches(3.4))
    tf2 = box.text_frame; tf2.word_wrap = True
    first = True
    for label, line in steps:
        p = para(tf2, first); first = False; p.space_after = Pt(6)
        if label:
            r = p.add_run(); r.text = label + "  "
            r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = th["accent"]; r.font.name = ZH_FONT
        add_rich_runs(p, line, 15, font=MATH_FONT)
    ans = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.15), Inches(12.1), Inches(0.65))
    ans.fill.solid(); ans.fill.fore_color.rgb = RGBColor(0xFD, 0xEF, 0xEE); ans.line.color.rgb = th["ans"]
    atf = ans.text_frame; atf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ap = para(atf, True); ap.alignment = PP_ALIGN.CENTER
    add_rich_runs(ap, "答案 Answer:  " + answer, 16, bold=True, color=th["ans"], font=MATH_FONT)
    footer(slide, th, page, tag)
    return slide


def excerpt_slide(prs, th, tag, title_en, title_zh, caption, source, img_path, img_w_in, page):
    from PIL import Image
    with Image.open(img_path) as im:
        ratio = im.width / im.height
    slide = add_slide(prs)
    header_bar(slide, title_en, title_zh, tag, th)
    cap = slide.shapes.add_textbox(Inches(0.6), Inches(1.08), Inches(12.1), Inches(0.75))
    tf = cap.text_frame; tf.word_wrap = True
    r = para(tf, True).add_run(); r.text = caption
    r.font.size = Pt(12.5); r.font.bold = True; r.font.color.rgb = th["dark"]; r.font.name = ZH_FONT
    r2 = tf.add_paragraph().add_run(); r2.text = source
    r2.font.size = Pt(9.5); r2.font.italic = True; r2.font.color.rgb = RGBColor(0x7A, 0x84, 0x8E)
    img_h = img_w_in / ratio
    x = (13.333 - img_w_in) / 2.0
    slide.shapes.add_picture(img_path, Inches(x), Inches(1.86), width=Inches(img_w_in))
    footer(slide, th, page, tag)
    return slide


def simple_list_slide(prs, th, tag, en, zh, items, page, en_size=15, zh_size=12.5):
    slide = add_slide(prs)
    header_bar(slide, en, zh, tag, th)
    bullets(slide, items, 0.7, 1.35, 12.0, 5.4, th, en_size=en_size, zh_size=zh_size)
    footer(slide, th, page, tag)
    return slide


def practice_slide(prs, th, tag, items, answers, page, title_en, title_zh, hint_zh):
    slide = add_slide(prs)
    header_bar(slide, title_en, title_zh, tag, th)
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.25), Inches(12.1), Inches(4.9))
    tf = box.text_frame; tf.word_wrap = True
    first = True
    for i, (en, zh) in enumerate(items, 1):
        p = para(tf, first); first = False; p.space_after = Pt(2)
        add_rich_runs(p, f"{i}. ", 15, bold=True, color=th["dark"])
        add_rich_runs(p, en, 15, font=MATH_FONT)
        p2 = tf.add_paragraph(); p2.space_after = Pt(12)
        r = p2.add_run(); r.text = "   " + zh
        r.font.size = Pt(12.5); r.font.color.rgb = RGBColor(0x6B, 0x72, 0x7B); r.font.name = ZH_FONT
    abox = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.25), Inches(12.1), Inches(0.6))
    abox.fill.solid(); abox.fill.fore_color.rgb = th["light"]; abox.line.color.rgb = th["accent"]
    atf = abox.text_frame; atf.vertical_anchor = MSO_ANCHOR.MIDDLE
    ap = para(atf, True); ap.alignment = PP_ALIGN.CENTER
    add_rich_runs(ap, "参考答案 Key:  " + answers, 13, bold=True, color=th["dark"], font=MATH_FONT)
    footer(slide, th, page, tag)
    return slide


def cover_slide(prs, th, course_line, unit_line_en, unit_line_zh):
    s = add_slide(prs)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(7.5))
    band.fill.solid(); band.fill.fore_color.rgb = th["dark"]; band.line.fill.background()
    t = s.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(2.6))
    tf = t.text_frame; tf.word_wrap = True
    p = para(tf, True); p.alignment = PP_ALIGN.CENTER
    add_rich_runs(p, course_line, 38, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    p = tf.add_paragraph(); p.alignment = PP_ALIGN.CENTER; p.space_before = Pt(10)
    add_rich_runs(p, unit_line_en, 26, color=RGBColor(0xE6, 0xEF, 0xF6))
    p = tf.add_paragraph(); p.alignment = PP_ALIGN.CENTER; p.space_before = Pt(6)
    r = p.add_run(); r.text = unit_line_zh
    r.font.size = Pt(22); r.font.color.rgb = RGBColor(0xE6, 0xEF, 0xF6); r.font.name = ZH_FONT
    sub = s.shapes.add_textbox(Inches(1.0), Inches(5.4), Inches(11.3), Inches(0.9))
    p = para(sub.text_frame, True); p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = "双语课件 Bilingual Courseware  |  含 AP 风格例题与真题节选 with AP-style examples & released-exam excerpts"
    r.font.size = Pt(13); r.font.color.rgb = RGBColor(0xBF, 0xD7, 0xEE); r.font.name = ZH_FONT
    return s


def lesson_slide(prs, th, tag, lesson, page):
    slide = add_slide(prs)
    header_bar(slide, f"Lesson {lesson['no']}  {lesson['title_en']}", lesson["title_zh"], tag, th)
    y = 1.3
    if lesson.get("concepts"):
        n = len(lesson["concepts"])
        bullets(slide, [(c["en"], c["zh"]) for c in lesson["concepts"]],
                0.7, y, 12.0, 2.1, th, en_size=14.5, zh_size=12)
        y += min(2.15 + 0.62 * n, 2.9)
    for f in lesson.get("formulas", []):
        formula_box(slide, f["f"], 1.2, y, 10.9, 0.78, th, size=19, sub=f.get("note"))
        y += 0.95
    for t in lesson.get("tables", []):
        hh = 0.5 + 0.42 * len(t["rows"])
        make_table(slide, t["header"], t["rows"], 1.2, y, 10.9, hh, th,
                   widths=t.get("widths"), font_size=12)
        y += hh + 0.25
    footer(slide, th, page, tag)
    return slide


def build_deck(unit, out_path):
    th = THEMES[unit["course"]]
    tag = COURSE_TAG[unit["course"]]
    prs = new_deck()
    page = 1
    cover_slide(prs, th, tag,
                f"Unit {unit['unit_no']}  {unit['unit_title_en']}",
                unit["unit_title_zh"])
    page = 2
    simple_list_slide(prs, th, tag, "Learning Objectives 学习目标",
                      "对应 AP 课程与考试描述 (CED)", unit["objectives"], 2)
    page = 3
    for lesson in unit["lessons"]:
        lesson_slide(prs, th, tag, lesson, page); page += 1
        for i, ex in enumerate(lesson["examples"], 1):
            example_slide(prs, th, tag, f"{lesson['no']}-{i}", ex.get("kind", "AP 风格"),
                          ex["en"], ex["zh"], ex["steps"], ex["answer"], page)
            page += 1
    if unit.get("excerpt"):
        e = unit["excerpt"]
        excerpt_slide(prs, th, tag, e.get("title_en", "真题节选 Released Exam Excerpt"),
                      e.get("title_zh", "官方真题对应本单元内容"), e["caption"], e["source"],
                      e["img"], e.get("width", 8.6), page)
        page += 1
    simple_list_slide(prs, th, tag, "Common Mistakes 常见错误",
                      "阅卷中最常扣分的点 Most-penalized errors", unit["mistakes"], page); page += 1
    simple_list_slide(prs, th, tag, "AP Exam Connection AP 考试链接",
                      "本单元如何考、如何拿分 How this unit earns points", unit["exam_link"], page); page += 1
    practice_slide(prs, th, tag, unit["deck_practice"], unit["deck_answers"], page,
                   "Class Practice 课堂练习", "限时完成 · 完成后对照参考答案", "Timed practice, then check the key")
    page += 1
    simple_list_slide(prs, th, tag, "Summary 小结", "本单元要点回顾 Key takeaways", unit["summary"], page); page += 1
    info = EXAM_INFO[unit["course"]]
    # 考试结构页
    s = add_slide(prs)
    header_bar(s, info["struct_title"][0], info["struct_title"][1], tag, th)
    make_table(s, info["struct_head"], info["struct_rows"], 1.0, 1.35, 11.3,
               0.55 + 0.55 * len(info["struct_rows"]), th,
               widths=[4.3, 2.2, 2.2, 2.6], font_size=13)
    bullets(s, info["struct_notes"], 1.0, 4.3, 11.3, 2.4, th, en_size=14, zh_size=12)
    footer(s, th, page, tag); page += 1
    # 分数页
    s = add_slide(prs)
    header_bar(s, info["score_title"][0], info["score_title"][1], tag, th)
    make_table(s, info["score_head"], info["score_rows"], 1.0, 1.35, 11.3, 1.6, th,
               widths=[2.4, 2.2, 2.4, 2.3, 2.0], font_size=13)
    bullets(s, info["score_notes"], 1.0, 3.25, 11.3, 3.2, th, en_size=14, zh_size=12)
    cap = s.shapes.add_textbox(Inches(1.0), Inches(6.55), Inches(11.3), Inches(0.35))
    r = para(cap.text_frame, True).add_run()
    r.text = info["score_source"]; r.font.size = Pt(9.5); r.font.italic = True
    r.font.color.rgb = RGBColor(0x7A, 0x84, 0x8E)
    footer(s, th, page, tag)
    prs.save(out_path)
    return page


# ---------- DOCX ----------

def _set_ea(run, ea=DOC_ZH_EA, latin=None):
    if latin:
        run.font.name = latin
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = rPr.makeelement(qn("w:rFonts"), {})
        rPr.append(rf)
    rf.set(qn("w:eastAsia"), ea)
    if latin and not rf.get(qn("w:ascii")):
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)


def docx_rich(p, text, size=11, bold=False, color=None, math=False, italic=False):
    for seg, kind in split_marked(text):
        r = p.add_run(seg)
        r.font.size = DPt(size)
        r.font.bold = bold
        r.font.italic = italic
        if color:
            r.font.color.rgb = color
        if math:
            _set_ea(r, DOC_ZH_EA, DOC_MATH)
        else:
            _set_ea(r, DOC_ZH_EA, EN_FONT)
        if kind == "sup":
            r.font.superscript = True
        elif kind == "sub":
            r.font.subscript = True
    return p


def _docx_header(doc, title, subtitle):
    h = doc.add_heading("", level=0)
    docx_rich(h, title, size=22, bold=True)
    p = doc.add_paragraph()
    docx_rich(p, subtitle, size=11, color=DRGB(0x59, 0x59, 0x59), italic=True)


def _docx_h(doc, text, level=1):
    h = doc.add_heading("", level=level)
    docx_rich(h, text, size=15 if level == 1 else 13, bold=True)


def build_workbooks(unit, hw_dir, test_dir, base):
    """生成 课后练习/答案(hw_dir) 与 综合测试/答案(test_dir) 四个 docx。"""
    tag = COURSE_TAG[unit["course"]]
    unit_name = f"Unit {unit['unit_no']}  {unit['unit_title_en']}  {unit['unit_title_zh']}"
    os.makedirs(hw_dir, exist_ok=True)
    os.makedirs(test_dir, exist_ok=True)

    # ---- 课后练习(题目卷) ----
    doc = Document()
    _docx_header(doc, f"{tag} 课后练习 · {unit_name}",
                 "Homework Problems · 题目卷(不含答案,答案另册) · 完成后对照《课后练习答案》")
    for hw in unit["hw"]:
        _docx_h(doc, f"Lesson {hw['lesson']}  {hw['title_en']}  {hw['title_zh']}")
        for i, pr in enumerate(hw["problems"], 1):
            p = doc.add_paragraph()
            docx_rich(p, f"{hw['lesson'].split('.')[-1]}.{i}  ", size=11, bold=True)
            docx_rich(p, pr["q"], size=11, math=True)
        doc.add_paragraph()
    doc.save(os.path.join(hw_dir, f"{base}_课后练习.docx"))

    # ---- 课后练习答案 ----
    doc = Document()
    _docx_header(doc, f"{tag} 课后练习答案 · {unit_name}",
                 "Answer Key · 按官方评分规范书写:每步标注得分点 Scoring guideline style, each step labeled with the point it earns")
    for hw in unit["hw"]:
        _docx_h(doc, f"Lesson {hw['lesson']}  {hw['title_en']}  {hw['title_zh']}")
        for i, pr in enumerate(hw["problems"], 1):
            p = doc.add_paragraph()
            docx_rich(p, f"{hw['lesson'].split('.')[-1]}.{i}  答案 Answer: ", size=11, bold=True)
            docx_rich(p, pr["ans"], size=11, math=True, bold=True)
            sol_lines = pr.get("sol") or []
            if isinstance(sol_lines, str):
                sol_lines = [sol_lines]
            for line in sol_lines:
                p2 = doc.add_paragraph()
                p2.paragraph_format.left_indent = DIn(0.35)
                p2.paragraph_format.space_after = DPt(2)
                docx_rich(p2, "· " + line, size=10.5, math=True, color=DRGB(0x40, 0x48, 0x50))
    doc.save(os.path.join(hw_dir, f"{base}_课后练习答案.docx"))

    # ---- 综合测试(题目卷) ----
    rv = unit["review"]
    doc = Document()
    _docx_header(doc, f"{tag} 章节综合测试 · {unit_name}",
                 "Unit Review Test · 建议限时 40 分钟 · Section A 选择题 + Section B 大题 + Challenge")
    if rv.get("img"):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(rv["img"], width=DIn(5.8))
        cap = doc.add_paragraph(); cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        docx_rich(cap, rv.get("img_caption", ""), size=9, italic=True, color=DRGB(0x59, 0x59, 0x59))
    _docx_h(doc, "Section A · 选择题 Multiple Choice(每题 4 分)")
    for i, q in enumerate(rv["mcq"], 1):
        p = doc.add_paragraph()
        docx_rich(p, f"A{i}.  ", size=11, bold=True)
        docx_rich(p, q["q"], size=11, math=True)
        for ch in q["choices"]:
            pc = doc.add_paragraph()
            pc.paragraph_format.left_indent = DIn(0.35)
            docx_rich(pc, ch, size=11, math=True)
    _docx_h(doc, "Section B · 大题 Free Response(每题 15 分,写出全部步骤)")
    for i, q in enumerate(rv["frq"], 1):
        p = doc.add_paragraph()
        docx_rich(p, f"B{i}.  ", size=11, bold=True)
        docx_rich(p, q["q"], size=11, math=True)
        for part in q["parts"]:
            pp = doc.add_paragraph()
            pp.paragraph_format.left_indent = DIn(0.35)
            docx_rich(pp, part, size=11, math=True)
    _docx_h(doc, "Challenge · 挑战题(10 分)")
    p = doc.add_paragraph()
    docx_rich(p, "C1.  ", size=11, bold=True)
    docx_rich(p, rv["bonus"]["q"], size=11, math=True)
    if rv.get("img_caption"):
        src = doc.add_paragraph()
        docx_rich(src, "真题出处: " + rv.get("img_source", "见课件真题节选页"), size=9,
                  italic=True, color=DRGB(0x59, 0x59, 0x59))
    doc.save(os.path.join(test_dir, f"{base}_综合测试.docx"))

    # ---- 综合测试答案 ----
    doc = Document()
    _docx_header(doc, f"{tag} 章节综合测试答案 · {unit_name}",
                 "Answer Key · 按官方评分规范:每步标注得分点 Rubric style, each scoring step labeled")
    _docx_h(doc, "Section A · 选择题")
    for i, q in enumerate(rv["mcq"], 1):
        p = doc.add_paragraph()
        docx_rich(p, f"A{i}.  ", size=11, bold=True)
        docx_rich(p, q["ans"], size=11, bold=True, math=True, color=DRGB(0xB0, 0x30, 0x2E))
        sol_lines = q.get("sol") or []
        if isinstance(sol_lines, str):
            sol_lines = [sol_lines]
        for line in sol_lines:
            p2 = doc.add_paragraph()
            p2.paragraph_format.left_indent = DIn(0.35)
            docx_rich(p2, "· " + line, size=10.5, math=True, color=DRGB(0x40, 0x48, 0x50))
    _docx_h(doc, "Section B · 大题")
    for i, q in enumerate(rv["frq"], 1):
        p = doc.add_paragraph()
        docx_rich(p, f"B{i}.  ", size=11, bold=True)
        docx_rich(p, q["ans"], size=11, bold=True, math=True, color=DRGB(0xB0, 0x30, 0x2E))
        sol_lines = q.get("sol") or []
        if isinstance(sol_lines, str):
            sol_lines = [sol_lines]
        for line in sol_lines:
            p2 = doc.add_paragraph()
            p2.paragraph_format.left_indent = DIn(0.35)
            docx_rich(p2, "· " + line, size=10.5, math=True, color=DRGB(0x40, 0x48, 0x50))
    _docx_h(doc, "Challenge")
    p = doc.add_paragraph()
    docx_rich(p, "C1.  ", size=11, bold=True)
    docx_rich(p, rv["bonus"]["ans"], size=11, bold=True, math=True, color=DRGB(0xB0, 0x30, 0x2E))
    sol_lines = rv["bonus"].get("sol") or []
    if isinstance(sol_lines, str):
        sol_lines = [sol_lines]
    for line in sol_lines:
        p2 = doc.add_paragraph()
        p2.paragraph_format.left_indent = DIn(0.35)
        docx_rich(p2, "· " + line, size=10.5, math=True, color=DRGB(0x40, 0x48, 0x50))
    doc.save(os.path.join(test_dir, f"{base}_综合测试答案.docx"))


def build_syllabus(doc_path, title, intro_en, intro_zh, course_key, units_meta, extra_notes):
    """units_meta: list of (unit_no, title_en, title_zh, lessons[(no,en,zh)], hours)"""
    info = EXAM_INFO[course_key]
    doc = Document()
    _docx_header(doc, title, "AP Course Syllabus · 教学大纲 · 双语 Bilingual")
    _docx_h(doc, "一、课程简介 Course Overview")
    p = doc.add_paragraph(); docx_rich(p, intro_en, size=11)
    p = doc.add_paragraph(); docx_rich(p, intro_zh, size=11)
    _docx_h(doc, "二、考试结构与计算器政策 Exam Structure & Calculator Policy")
    t = doc.add_table(rows=1 + len(info["struct_rows"]), cols=4)
    t.style = "Table Grid"
    for j, htxt in enumerate(info["struct_head"]):
        c = t.cell(0, j).paragraphs[0]
        docx_rich(c, htxt, size=10.5, bold=True)
    for i, row in enumerate(info["struct_rows"], 1):
        for j, v in enumerate(row):
            docx_rich(t.cell(i, j).paragraphs[0], v, size=10.5)
    _docx_h(doc, "三、分数分布 Score Distributions")
    t = doc.add_table(rows=1 + len(info["score_rows"]), cols=5)
    t.style = "Table Grid"
    for j, htxt in enumerate(info["score_head"]):
        docx_rich(t.cell(0, j).paragraphs[0], htxt, size=10.5, bold=True)
    for i, row in enumerate(info["score_rows"], 1):
        for j, v in enumerate(row):
            docx_rich(t.cell(i, j).paragraphs[0], v, size=10.5)
    p = doc.add_paragraph()
    docx_rich(p, info["score_source"], size=9, italic=True, color=DRGB(0x59, 0x59, 0x59))
    _docx_h(doc, "四、单元与课时安排 Units & Schedule")
    t = doc.add_table(rows=1 + len(units_meta), cols=4)
    t.style = "Table Grid"
    for j, htxt in enumerate(["单元 Unit", "内容 Content", "小节 Lessons", "课时 Hours"]):
        docx_rich(t.cell(0, j).paragraphs[0], htxt, size=10.5, bold=True)
    for i, (no, ten, tzh, lessons, hours) in enumerate(units_meta, 1):
        docx_rich(t.cell(i, 0).paragraphs[0], f"Unit {no}", size=10.5, bold=True)
        docx_rich(t.cell(i, 1).paragraphs[0], f"{ten} {tzh}", size=10.5)
        docx_rich(t.cell(i, 2).paragraphs[0], "; ".join(l[0] + " " + l[1] for l in lessons), size=9.5)
        docx_rich(t.cell(i, 3).paragraphs[0], str(hours), size=10.5)
    _docx_h(doc, "五、每单元配套材料 Materials per Unit")
    for note in extra_notes:
        p = doc.add_paragraph(style="List Bullet")
        docx_rich(p, note, size=10.5)
    _docx_h(doc, "六、学习建议 Study Advice")
    for tip in [
        ("Preview the PPT before class; redo every worked example with the solution covered.",
         "课前预读课件;做例题时先遮住详解,完成后再逐行对照。"),
        ("Do homework the same day; answers are in a separate booklet — check only after finishing.",
         "当天完成课后练习;答案独立成册,务必做完再对答案。"),
        ("Timed unit tests build exam pace: Section A 12 min, Section B 24 min, Challenge 8 min.",
         "综合测试按考试节奏限时:A 部分 12 分钟、B 部分 24 分钟、挑战题 8 分钟。"),
        ("All formulas follow the same notation as the slides; superscripts/subscripts are typeset, never garbled.",
         "全套材料公式记号与课件一致,上下标均为排版生成,不会出现乱码。"),
    ]:
        p = doc.add_paragraph(style="List Bullet")
        docx_rich(p, tip[0] + "  ", size=10.5)
        docx_rich(p, tip[1], size=10.5, color=DRGB(0x40, 0x48, 0x50))
    doc.save(doc_path)

