import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[108]
for sh in s.shapes:
    if sh.has_text_frame and 4.7<sh.top/EMU<15.5:
        szs=[r.font.size.pt if r.font.size else None for p in sh.text_frame.paragraphs for r in p.runs]
        bolds=[r.font.bold for p in sh.text_frame.paragraphs for r in p.runs]
        print(f"t{sh.top/EMU:5.2f} {sh.name:12s} sz={szs} b={bolds} | {sh.text_frame.text.strip()[:40]}")
print("---- 行底/斑马 ----")
for sh in s.shapes:
    if sh.shape_type==1:
        print(sh.name, sh.left/EMU, sh.top/EMU, sh.width/EMU, sh.height/EMU, getattr(sh.fill.fore_color,'rgb',None) if sh.fill.type is not None else None)
