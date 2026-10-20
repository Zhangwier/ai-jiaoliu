# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s=prs.slides[108]
for sh in s.shapes:
    if sh.name=="Rectangle 3":
        print("before:", repr(sh.text_frame.text[:40]) if sh.has_text_frame else "no tf")
        if sh.has_text_frame:
            for p in list(sh.text_frame.paragraphs)[1:]:
                p._p.getparent().remove(p._p)
            p0=sh.text_frame.paragraphs[0]
            for r in list(p0.runs): r._r.getparent().remove(r._r)
        print("cleared")
# 副标位置微调（不要压装饰条下的第一行）
for sh in s.shapes:
    if sh.name=="TextBox 5":
        sh.top=Cm(3.62); print("sub top ->", sh.top/360000)
save(prs)
