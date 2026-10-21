import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for n in (54,59,80,88,93,110,121):
    s=prs.slides[n-1]
    rows=[]
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/EMU>22 and sh.top/EMU>4.9:
            rows.append((round(sh.top/EMU,2), sh.text_frame.text.strip()))
    print(n, sorted(rows))
