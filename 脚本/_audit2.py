import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for i,s in enumerate(prs.slides,1):
    if i<50: continue
    best=None
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        tx=sh.text_frame.text.strip().replace("\n","/")
        if not tx: continue
        t=sh.top/EMU
        if 1.6<=t<=4.8 and sh.left/EMU<22:
            sz=0
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size: sz=max(sz,r.font.size.pt)
            if sz>=20 and (best is None or len(tx)>len(best[1]) or sz>best[2]):
                if best is None or sz>best[2] or (sz>=best[2] and len(tx)>4):
                    best=(round(t,1),tx,sz)
    print(f"{i:3d} | {(best[1] if best else '')[:64]}")
