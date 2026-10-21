import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for n in (53,59,80,88,93,110,121):
    s=prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame and sh.top/EMU<14 and sh.left/EMU<22 and sh.top/EMU>7:
            print(n, repr(sh.text_frame.text))
    # side list
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/EMU>22 and sh.top/EMU>4.9:
            print("   ", sh.name, sh.top/EMU, [p.text for p in sh.text_frame.paragraphs])
    print()
