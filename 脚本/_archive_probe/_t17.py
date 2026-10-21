import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for n in (54,59,93,110):
    s=prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/EMU<22 and 6<sh.top/EMU<16:
            print(n, sh.name, round(sh.top/EMU,2), round(sh.width/EMU,2), round(sh.height/EMU,2), "|", repr(sh.text_frame.text[:90]))
            print("   runs:", [(r.text[:10], r.font.size.pt if r.font.size else None, r.font.color.rgb if r.font.color and r.font.color.type is not None else None) for p in sh.text_frame.paragraphs for r in p.runs][:4])
