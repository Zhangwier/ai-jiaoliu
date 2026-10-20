import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm
EMU=360000
prs = openprs(); s = prs.slides[80]
bands = [4.55 + 1.78*i for i in range(7)]   # 原 7 行起始
newY = [4.42 + 1.50*i for i in range(7)]
for sh in s.shapes:
    t = sh.top/EMU
    if t < 4.3 or t > 15.9: continue
    idx = None
    for i,b in enumerate(bands):
        if b - 0.05 <= t < b + 1.85:
            idx = i; break
    if idx is None: continue
    delta = newY[idx] - bands[idx]
    sh.top = Cm(t + delta)
    h = sh.height/EMU
    if h > 1.6: sh.height = Cm(1.50)
    elif abs(h-0.66) < 0.06: sh.height = Cm(0.60)
    else: sh.height = Cm(1.42)
save(prs); print("ok")
