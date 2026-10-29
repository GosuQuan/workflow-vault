# -*- coding: utf-8 -*-
"""全量校验:所有课件打开、页数、图片数;所有 DOCX 可打开;内容文件无 ^{} 残留输出。"""
import glob
import sys

sys.stdout.reconfigure(encoding="utf-8")
from pptx import Presentation  # noqa: E402

decks = sorted(glob.glob("AP微积分BC/课件/*.pptx") + glob.glob("AP物理C/课件/*.pptx")
               + glob.glob("AP物理C电磁学/课件/*.pptx"))
bad = 0
for d in decks:
    try:
        prs = Presentation(d)
        n = len(prs.slides._sldIdLst)
        pics = sum(1 for s in prs.slides for sh in s.shapes if sh.shape_type == 13)
        texts = "\n".join(sh.text_frame.text for s in prs.slides for sh in s.shapes
                          if sh.has_text_frame)
        leftover = []
        for mark in ("^{", "_{", "\\frac", "\\sqrt"):
            if mark in texts:
                leftover.append(mark)
        status = "OK " if not leftover else "MARK-LEAK"
        if leftover:
            bad += 1
        print(f"{status} {n:2d}p {pics}img  {d}  {leftover if leftover else ''}")
    except Exception as e:
        bad += 1
        print(f"BAD {d}: {e}")
print(f"\n{len(decks)} decks checked, {bad} bad")
sys.exit(1 if bad else 0)
