# -*- coding: utf-8 -*-
"""Crop real AP exam question excerpts from rendered pages."""
from PIL import Image
import os

os.makedirs("assets/crops", exist_ok=True)

# coords estimated on 1224x1584 (144 dpi) renders, scaled by 1870/1224
S = 1870 / 1224.0


def crop(src, box144, dst):
    img = Image.open(src)
    l, t, r, b = [int(v * S) for v in box144]
    l = max(0, l); t = max(0, t)
    r = min(img.width, r); b = min(img.height, b)
    img.crop((l, t, r, b)).save(dst)
    print(dst, r - l, "x", b - t)


# Physics C 2024 Q2 scoring guideline: part (a) rubric + example solution
crop("assets/render/q2hi/page-0001.png", (80, 148, 1150, 800), "assets/crops/phys_q2a.png")
# Physics C 2024 Q2: (b)(i) rubric row + v-t graph (terminal velocity)
crop("assets/render/q2hi/page-0001.png", (80, 818, 1150, 1370), "assets/crops/phys_q2b.png")
# Calculus BC 2024 FRQ Q5 full question block (with page header)
crop("assets/render/calchi/page-0001.png", (55, 15, 1165, 670), "assets/crops/calc_q5.png")
