import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s=prs.slides[62]
for sh in s.shapes:
    if sh.has_text_frame and "见 C5" in sh.text_frame.text:
        print(repr(sh.text_frame.text))
        print(sh.name, sh.left/360000, sh.top/360000, sh.width/360000, sh.height/360000)
