import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s=prs.slides[4]
for sh in s.shapes:
    if sh.has_text_frame and "图纸审查" in sh.text_frame.text:
        settext(sh, "图纸审查（看图与 OCR）、报告原文对比", size=15)
        print([r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs])
    elif sh.has_text_frame and ("鹈鹕盲测" in sh.text_frame.text or "五个短板" in sh.text_frame.text):
        print(sh.text_frame.text[:20], [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs][:3])
save(prs)
