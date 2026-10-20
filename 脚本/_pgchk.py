import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
EMU=360000
bad=[]
for i, s in enumerate(prs.slides, 1):
    found=None
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text.strip()
        if re.fullmatch(r"\d{1,3}", t) and sh.left/EMU>28 and sh.top/EMU<1.6:
            found=int(t); break
    if found is None:
        if i<=124: bad.append((i,"none"))
    elif found != i:
        bad.append((i,found))
print("mismatch:", bad)
