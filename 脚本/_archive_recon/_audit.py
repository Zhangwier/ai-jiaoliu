import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
EMU=360000
for i,s in enumerate(prs.slides,1):
    t=None; pg=None
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        tx=sh.text_frame.text.strip()
        if not tx: continue
        if sh.top/EMU<1.6 and sh.left/EMU<16 and len(tx)>4 and t is None: t=tx
        if sh.top/EMU<1.6 and sh.left/EMU>28: pg=tx
    print(f"{i:3d} pg={pg!s:>4} | {(t or '')[:60].replace(chr(10),'/')}")
