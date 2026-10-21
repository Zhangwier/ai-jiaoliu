# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
for i, s in enumerate(prs.slides, 1):
    num = ""
    for sh in s.shapes:
        if sh.name == "TextBox 2" and sh.has_text_frame:
            num = sh.text_frame.text.strip()
            break
    print("%3d  id=%-3s num=%-4s  %s" % (i, s.slide_id, num, title_of(s).replace("\n"," / ")[:50]))
