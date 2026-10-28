# -*- coding: utf-8 -*-
"""校验单元内容文件:python validate_unit.py apgen/content/calc_u01.py"""
import importlib.util
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

ALLOWED_SPECIAL = set("∫√πθαβμωΣΔ∞≈≠≤≥×·→′°⇒∂ρλτσΩεϕτ½¼¾±∥⊥ᵢⱼ⊗∈")
ALLOWED_SPECIAL |= set("()[]{},.;:+-*/=<>|'\" !?%&@#`~_0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")


def check_text(path_ctx, s, errs):
    """拦截真正会导致乱码的字符:私用区、控制符、代理区。常用数学/标点符号全部放行。"""
    if not isinstance(s, str):
        errs.append(f"{path_ctx}: 不是字符串 {type(s)}")
        return
    for ch in s:
        o = ord(ch)
        bad = (
            0xE000 <= o <= 0xF8FF            # private use area
            or 0xFFF0 <= o <= 0xFFFF         # specials / replacements (含乱码方块)
            or 0xD800 <= o <= 0xDFFF         # surrogates
            or (o < 32 and ch not in "\n\t")
        )
        if bad:
            errs.append(f"{path_ctx}: 非法字符 {ch!r} (U+{o:04X}) — 私用区/替换符会乱码")
    if "\\frac" in s or "\\sqrt" in s or "$" in s or "\\" in s:
        errs.append(f"{path_ctx}: 含 LaTeX 命令或反斜杠,禁止使用;分数写 a/b,根号写 √(...)")
    if "⁰" in s or "¹" in s or "²" in s or "³" in s or "ⁿ" in s:
        errs.append(f"{path_ctx}: 请用 ^{{}} 标记表示上标(如 x^{{2}}),不要用 Unicode 上标字符")


def need(d, key, ctx, errs, typ=None):
    if key not in d:
        errs.append(f"{ctx}: 缺少字段 {key}")
        return None
    if typ is not None and not isinstance(d[key], typ):
        errs.append(f"{ctx}.{key}: 类型应为 {typ.__name__}")
    return d[key]


def validate(mod):
    errs = []
    U = getattr(mod, "UNIT", None)
    if U is None:
        return ["文件必须定义 UNIT 字典"]
    for k in ("course", "unit_no", "unit_title_en", "unit_title_zh", "objectives", "lessons",
              "mistakes", "exam_link", "deck_practice", "deck_answers", "summary", "hw", "review"):
        need(U, k, "UNIT", errs)
    if U.get("course") not in ("calc", "phys", "em"):
        errs.append("UNIT.course 必须是 calc/phys/em")
    if len(U.get("lessons", [])) != 3:
        errs.append("UNIT.lessons 必须恰好 3 个小节")
    if len(U.get("objectives", [])) < 4:
        errs.append("UNIT.objectives 至少 4 条")
    for i, o in enumerate(U.get("objectives", [])):
        if len(o) != 2:
            errs.append(f"objectives[{i}] 必须是 (en, zh)")
        else:
            check_text(f"objectives[{i}].en", o[0], errs); check_text(f"objectives[{i}].zh", o[1], errs)

    lesson_nos = []
    for li, L in enumerate(U.get("lessons", [])):
        ctx = f"lessons[{li}]"
        for k in ("no", "title_en", "title_zh", "concepts", "formulas", "examples"):
            need(L, k, ctx, errs)
        lesson_nos.append(L.get("no"))
        if len(L.get("concepts", [])) != 4:
            errs.append(f"{ctx}.concepts 必须恰好 4 条(版式约束)")
        if not (2 <= len(L.get("formulas", [])) <= 4):
            errs.append(f"{ctx}.formulas 应 2-4 条")
        if len(L.get("tables", [])) > 1:
            errs.append(f"{ctx}.tables 最多 1 个")
        for t in L.get("tables", []):
            if len(t.get("rows", [])) > 3:
                errs.append(f"{ctx}.tables 行数 ≤ 3(版式约束)")
        if len(L.get("examples", [])) < 2:
            errs.append(f"{ctx}.examples 至少 2 道例题")
        for c in L.get("concepts", []):
            check_text(ctx + ".concept.en", c["en"], errs); check_text(ctx + ".concept.zh", c["zh"], errs)
        for fi, f in enumerate(L.get("formulas", [])):
            check_text(f"{ctx}.formulas[{fi}].f", f.get("f", ""), errs)
            if len(f.get("f", "")) > 95:
                errs.append(f"{ctx}.formulas[{fi}].f 过长({len(f['f'])}字符),拆成两条或精简")
            if f.get("note"):
                check_text(f"{ctx}.formulas[{fi}].note", f["note"], errs)
        for t in L.get("tables", []):
            need(t, "header", ctx + ".tables", errs)
            need(t, "rows", ctx + ".tables", errs)
        for ei, ex in enumerate(L.get("examples", [])):
            ec = f"{ctx}.examples[{ei}]"
            for k in ("en", "zh", "steps", "answer"):
                need(ex, k, ec, errs)
            check_text(f"{ec}.en", ex.get("en", ""), errs)
            check_text(f"{ec}.zh", ex.get("zh", ""), errs)
            if len(ex.get("steps", [])) < 3:
                errs.append(f"{ec}.steps 至少 3 步")
            for lab, line in ex.get("steps", []):
                check_text(f"{ec}.step", line, errs)

    e = U.get("excerpt")
    if e:
        import os
        if not os.path.exists(e.get("img", "")):
            errs.append(f"excerpt.img 文件不存在: {e.get('img')}")

    for name, lst in (("mistakes", U.get("mistakes")), ("exam_link", U.get("exam_link")),
                      ("summary", U.get("summary"))):
        if len(lst) < 3:
            errs.append(f"{name} 至少 3 条")
        for i, o in enumerate(lst):
            if len(o) != 2:
                errs.append(f"{name}[{i}] 必须是 (en, zh)")
            else:
                check_text(f"{name}[{i}].en", o[0], errs); check_text(f"{name}[{i}].zh", o[1], errs)
    if len(U.get("deck_practice", [])) != 3:
        errs.append("deck_practice 必须恰好 3 题")
    check_text("deck_answers", U.get("deck_answers", ""), errs)

    if len(U.get("hw", [])) != 3:
        errs.append("hw 必须恰好 3 个小节,与 lessons 对应")
    hw_nos = [h.get("lesson") for h in U.get("hw", [])]
    if lesson_nos and hw_nos != lesson_nos:
        errs.append(f"hw.lesson {hw_nos} 与 lessons 编号 {lesson_nos} 不一致")
    for hi, h in enumerate(U.get("hw", [])):
        if len(h.get("problems", [])) != 5:
            errs.append(f"hw[{hi}].problems 必须恰好 5 题")
        for pi, pr in enumerate(h.get("problems", [])):
            pc = f"hw[{hi}].problems[{pi}]"
            for k in ("q", "ans"):
                need(pr, k, pc, errs)
            check_text(f"{pc}.q", pr.get("q", ""), errs)
            check_text(f"{pc}.ans", pr.get("ans", ""), errs)
            sol = pr.get("sol")
            if sol is None:
                errs.append(f"{pc}.sol 缺失 — 评分步骤式解析为必填")
            else:
                lines = sol if isinstance(sol, list) else [sol]
                if len(lines) < 2:
                    errs.append(f"{pc}.sol 至少 2 个评分步骤(按官方评分规范)")
                for line in lines:
                    check_text(f"{pc}.sol", line, errs)

    rv = need(U, "review", "UNIT", errs, dict) or {}
    if rv.get("img"):
        import os
        if not os.path.exists(rv["img"]):
            errs.append(f"review.img 文件不存在: {rv['img']}")
    if len(rv.get("mcq", [])) != 5:
        errs.append("review.mcq 必须恰好 5 题")
    for i, q in enumerate(rv.get("mcq", [])):
        qc = f"review.mcq[{i}]"
        for k in ("q", "choices", "ans", "sol"):
            need(q, k, qc, errs)
        if len(q.get("choices", [])) != 4:
            errs.append(f"{qc}.choices 必须 4 个选项 (A)-(D)")
        check_text(f"{qc}.ans", q.get("ans", ""), errs)
        sol = q.get("sol")
        lines = sol if isinstance(sol, list) else ([sol] if sol else [])
        if len(lines) < 1:
            errs.append(f"{qc}.sol 缺失")
        for line in lines:
            check_text(f"{qc}.sol", line, errs)
    if len(rv.get("frq", [])) != 2:
        errs.append("review.frq 必须恰好 2 题")
    for i, q in enumerate(rv.get("frq", [])):
        qc = f"review.frq[{i}]"
        for k in ("q", "parts", "ans", "sol"):
            need(q, k, qc, errs)
        if len(q.get("parts", [])) < 2:
            errs.append(f"{qc}.parts 至少 2 问")
        for line in ([q["sol"]] if isinstance(q.get("sol"), str) else q.get("sol", [])):
            check_text(f"{qc}.sol", line, errs)
    b = need(rv, "bonus", "review", errs, dict)
    if b:
        for k in ("q", "ans", "sol"):
            need(b, k, "review.bonus", errs)
        for line in ([b["sol"]] if isinstance(b.get("sol"), str) else b.get("sol", [])):
            check_text("review.bonus.sol", line, errs)
    return errs


def load(path):
    spec = importlib.util.spec_from_file_location("unit_content", path)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: validate_unit.py <content_file.py>")
        sys.exit(2)
    m = load(sys.argv[1])
    errs = validate(m)
    if errs:
        print("FAIL")
        for e in errs:
            print("  -", e)
        sys.exit(1)
    print(f"OK  {m.UNIT['course']} U{m.UNIT['unit_no']:02d} {m.UNIT['unit_title_en']}")
