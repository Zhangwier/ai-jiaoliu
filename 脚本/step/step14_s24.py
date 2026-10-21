# -*- coding: utf-8 -*-
import os
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

P = r"D:\MG\program\ai交流\成品\_work29\base.pptx"
DIRS = [r"D:\MG\program\ai交流\成品\_work29\logos3", r"D:\MG\program\ai交流\成品\_work29\logos2", r"D:\MG\program\ai交流\成品\_work29\logos"]
def pick(n):
    for d in DIRS:
        p = os.path.join(d, n+".png")
        if os.path.exists(p): return p
    return None

prs = Presentation(P)
s = prs.slides[23]
S = {sh.shape_id: sh for sh in s.shapes}

# (cell id, logo x start, logo keys, ring color)
ROWS = [
    (9,  ["openai","anthropic"],        RGBColor(0xB9,0xC6,0xD8)),
    (13, ["google","meta"],             RGBColor(0xB9,0xC6,0xD8)),
    (17, ["grok","alibaba","zhipu","moonshot"], RGBColor(0x8F,0xC9,0xC3)),
    (21, ["tencent","minimax","doubao"], RGBColor(0xC9,0xC6,0xBC)),
]
SIZE, GAP = 0.62, 0.11

for cid, keys, ring in ROWS:
    cell = S[cid]
    top = (cell.top + Cm(1.55)) / 360000
    left = (cell.left + Cm(0.39)) / 360000
    x = left
    n = 0
    for k in keys:
        pth = pick(k)
        if not pth:
            print("miss", cid, k); continue
        ov = s.shapes.add_shape(MSO_SHAPE.OVAL, Cm(x), Cm(top), Cm(SIZE), Cm(SIZE))
        ov.fill.solid(); ov.fill.fore_color.rgb = RGBColor(0xFF,0xFF,0xFF)
        ov.line.color.rgb = ring; ov.line.width = Pt(0.75)
        ov.shadow.inherit = False; ov.name = "LogoBadge"
        pad = SIZE*0.17
        s.shapes.add_picture(pth, Cm(x+pad), Cm(top+pad), Cm(SIZE-2*pad), Cm(SIZE-2*pad)).name = "Logo"
        x += SIZE + GAP
        n += 1
    print("cell", cid, keys, "->", n, "badges, right edge", round(x-GAP,2))
prs.save(P)
