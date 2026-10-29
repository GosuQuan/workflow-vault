# -*- coding: utf-8 -*-
"""构建全部单元材料:python build_all.py [course ...]
用法: python build_all.py            # 全部
      python build_all.py calc      # 只构建 calc
"""
import importlib.util
import os
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from framework import build_deck, build_workbooks, build_syllabus  # noqa: E402
from validate_unit import validate, load  # noqa: E402

COURSE_FOLDER = {"calc": "AP微积分BC", "phys": "AP物理C", "em": "AP物理C电磁学"}


def content_files(courses):
    out = []
    cdir = os.path.join(HERE, "content")
    for fn in sorted(os.listdir(cdir)):
        if not fn.endswith(".py") or fn.startswith("_"):
            continue
        key = fn.split("_")[0]
        if key in courses:
            out.append(os.path.join(cdir, fn))
    return out


def main():
    courses = set(sys.argv[1:]) or {"calc", "phys", "em"}
    files = content_files(courses)
    if not files:
        print("no content files found")
        return
    ok, bad = [], []
    for path in files:
        m = load(path)
        errs = validate(m)
        name = os.path.basename(path)
        if errs:
            bad.append((name, errs))
            continue
        U = m.UNIT
        folder = COURSE_FOLDER[U["course"]]
        base = f"Unit{U['unit_no']:02d}_{U['unit_title_zh']}"
        deck_path = os.path.join(ROOT, folder, "课件", f"{base}.pptx")
        os.makedirs(os.path.dirname(deck_path), exist_ok=True)
        slides = build_deck(U, deck_path)
        hw_dir = os.path.join(ROOT, folder, "课后习题")
        test_dir = os.path.join(ROOT, folder, "章节测试")
        build_workbooks(U, hw_dir, test_dir, base)
        ok.append((name, folder, base, slides))
        print(f"OK   {name} -> {folder}/{base} ({slides} slides)")
    for name, errs in bad:
        print(f"FAIL {name}")
        for e in errs:
            print("   -", e)
    print(f"\n{len(ok)} built, {len(bad)} failed")
    if bad:
        sys.exit(1)


if __name__ == "__main__":
    main()
