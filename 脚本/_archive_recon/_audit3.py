import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
bad=[]
for i,s in enumerate(prs.slides,1):
    if i>124: break
    found=None
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        tx=sh.text_frame.text.strip()
        if sh.top/EMU<1.6 and sh.left/EMU>28 and re.fullmatch(r"\d+", tx):
            found=int(tx)
    if found!=i: bad.append((i,found))
print("mismatch:", bad)
