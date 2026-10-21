# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs()
s=prs.slides[120]
for sh in s.shapes:
    if sh.has_text_frame and sh.left/EMU<22 and 12.0<sh.top/EMU<14.6:
        settext(sh,"Z1 按编号归档；Z2、Z3 为可直接点开的原链接；正文页脚都指到这里", size=14)
        print("ok")
save(prs)
