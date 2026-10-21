# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm
EMU=360000
prs = openprs(); s = prs.slides[39]
def at(t, idx=0):
    lst=[]
    for sh in s.shapes:
        if sh.has_text_frame and abs(sh.top/EMU-t)<0.03: lst.append(sh)
    lst.sort(key=lambda x: x.left); return lst

# 卡片白底缩到 15.30
for sh in s.shapes:
    if not sh.has_text_frame and abs(sh.top/EMU-4.50)<0.03 and sh.height/EMU>10.8:
        sh.height = Cm(10.85)
# 标题：高度 1.35
for sh in at(5.7):
    sh.height = Cm(1.32)
# 正文下移到 7.02
for sh in at(6.9):
    sh.top = Cm(7.02); sh.height = Cm(1.55)
# 典型场景 8.65 → 8.58
for sh in at(8.65):
    sh.top = Cm(8.58)
for sh in at(9.15):
    sh.top = Cm(9.05); sh.height = Cm(1.70)
for sh in at(10.95):
    sh.top = Cm(10.78)
for sh in at(11.45):
    sh.top = Cm(11.25); sh.height = Cm(1.70)
for sh in at(13.15):
    if sh.has_text_frame and sh.text_frame.text.startswith("怎么做"):
        sh.top = Cm(13.05); sh.height = Cm(1.30)
    elif not sh.has_text_frame:
        sh.top = Cm(13.05); sh.height = Cm(1.30)
# 顺序行
for sh in at(15.6):
    sh.top = Cm(15.42)
save(prs); print("ok")
