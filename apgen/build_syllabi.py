# -*- coding: utf-8 -*-
"""生成两份 Word 教学大纲。python build_syllabi.py"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from framework import build_syllabus  # noqa: E402
from validate_unit import load  # noqa: E402


def collect(prefix, folder):
    metas = []
    for no in range(1, 12):
        path = os.path.join(HERE, "content", f"{prefix}_u{no:02d}.py")
        if not os.path.exists(path):
            continue
        U = load(path).UNIT
        metas.append((U["unit_no"], U["unit_title_en"], U["unit_title_zh"],
                      [(l["no"], l["title_en"], l["title_zh"]) for l in U["lessons"]],
                      3 * 2 + 2))
    return metas


calc = collect("calc", None)
phys = collect("phys", None)
em = collect("em", None)
phys_em = phys + [(7 + i, u[1], u[2], u[3], u[4]) for i, u in enumerate(em, 1)]

build_syllabus(
    os.path.join(ROOT, "AP微积分BC", "教学大纲_AP_Calculus_BC.docx"),
    "AP Calculus BC 教学大纲",
    "A full-year, 10-unit bilingual (EN/ZH) course covering the complete AP Calculus BC curriculum: limits, differentiation, integration, differential equations, parametric/polar/vector functions and infinite series. Every unit ships with a 16-17 slide bilingual deck, graded homework with rubric-style answer keys, and a timed unit test with released-exam excerpts.",
    "本课程为 AP 微积分 BC 全程双语课程(10 个单元):极限、求导、积分、微分方程、参数/极坐标/向量函数与无穷级数。每个单元配套双语课件(16–17 页)、按官方评分规范书写解析的课后练习(题目与答案分册)、含真题节选的限时章节测试。全套材料公式均采用排版上下标,杜绝乱码。",
    "calc", calc,
    ["课件 PPTX:每单元一份(概念 + 例题详解 + 真题节选 + 课堂练习 + 考情)",
     "课后练习 DOCX:题目卷与答案卷分离,答案按评分规范逐步给分",
     "章节综合测试 DOCX:Section A 选择 + Section B 大题 + Challenge,含真题原题节选,限时 40 分钟",
     "官方真题 PDF:已存档于课程文件夹,仅教学使用并注明出处"])

build_syllabus(
    os.path.join(ROOT, "AP物理C", "教学大纲_AP_Physics_C_力学与电磁学.docx"),
    "AP Physics C 教学大纲(力学 + 电磁学)",
    "A calculus-based bilingual course covering AP Physics C: Mechanics (7 units) and Electricity & Magnetism (5 units). Mechanics: kinematics, Newton's laws with variable forces, energy, momentum, rotation, oscillations, gravitation. E&M: electrostatics, conductors/capacitors, circuits (incl. RC transients), magnetic fields, induction. Every unit ships with a bilingual deck, rubric-graded homework, and a timed unit test.",
    "本课程为基于微积分的 AP 物理 C 双语课程:力学 7 单元(运动学、牛顿定律与变力、能量、动量、转动、振动、引力)+ 电磁学 5 单元(静电、导体与电容、电路含 RC 暂态、磁场、电磁感应)。每单元配套双语课件、评分规范式课后练习(题目答案分册)、含真题节选的限时章节测试。",
    "phys", phys_em,
    ["课件 PPTX:每单元一份(概念 + 例题详解 + 真题评分标准节选 + 课堂练习 + 考情)",
     "课后练习 DOCX:题目卷与答案卷分离,解析按官方评分规范逐步给分并标注得分点",
     "章节综合测试 DOCX:Section A 选择 + Section B 大题 + Challenge,含 2024 官方真题节选,限时 40 分钟",
     "官方真题 PDF 与 CED:已存档于课程文件夹(不入 git),仅教学使用并注明出处"])

print("syllabi built:")
print("  AP微积分BC/教学大纲_AP_Calculus_BC.docx")
print("  AP物理C/教学大纲_AP_Physics_C_力学与电磁学.docx")
