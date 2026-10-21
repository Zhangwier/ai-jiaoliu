import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[4]
for sh in s.shapes:
    if sh.has_text_frame and "图纸审查" in sh.text_frame.text:
        print(sh.name, sh.left/EMU, sh.top/EMU, sh.width/EMU, sh.height/EMU)
        for p in sh.text_frame.paragraphs:
            for r in p.runs: print("  sz", r.font.size)
        settext(sh, "图纸审查（看图与 OCR）、报告原文对比", size=14.5)
        print("ok")
save(prs)
