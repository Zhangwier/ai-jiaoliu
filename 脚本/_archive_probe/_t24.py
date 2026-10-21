import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs(); EMU=360000
for idx in (53,58,79,87,92,109,120):
    s=prs.slides[idx]
    print("="*16, idx+1)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.top/EMU<15:
            v=sh.text_frame.text.replace("\n","\\n")
            print(f"   t{sh.top/EMU:5.2f} l{sh.left/EMU:5.2f} {sh.name:12s} {v[:100]!r}")
