# -*- coding: utf-8 -*-
"""s40：标题改单行，重排节奏"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
EMU=360000
prs = openprs(); s = prs.slides[39]
def at(t):
    lst=[sh for sh in s.shapes if sh.has_text_frame and abs(sh.top/EMU-t)<0.03]
    lst.sort(key=lambda x: x.left); return lst

heads = ["你不知道有它", "一堆低级活", "要判断、要取舍"]
for sh, h in zip(at(5.7), heads):
    settext(sh, h, size=16)
    sh.height = Cm(0.90)
    for p in sh.text_frame.paragraphs: p.line_spacing = 1.15

# 正文 → 6.70
for sh in at(7.02):
    sh.top = Cm(6.70); sh.height = Cm(1.75)
# 典型场景标签 → 8.55；内容 8.98
for sh in at(8.58): sh.top = Cm(8.52)
for sh in at(9.05): sh.top = Cm(8.95); sh.height = Cm(1.75)
for sh in at(10.78): sh.top = Cm(10.72)
for sh in at(11.25): sh.top = Cm(11.15); sh.height = Cm(1.75)
for sh in at(13.05):
    sh.top = Cm(13.05)
save(prs); print("ok")
