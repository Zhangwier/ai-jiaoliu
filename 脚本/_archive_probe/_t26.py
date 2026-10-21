import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs(); EMU=360000
for idx in (5,17,27,38):
    s=prs.slides[idx]
    print("="*16, idx+1)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.left/EMU>20:
            print(f"   t{sh.top/EMU:5.2f} {sh.text_frame.text.strip()!r}")
