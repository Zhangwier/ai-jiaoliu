import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs(); EMU=360000
for idx in (92,52,108):
    s=prs.slides[idx]
    print("="*20, idx+1)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"  {sh.name:12s} t{sh.top/EMU:5.2f} | {sh.text_frame.text.strip()[:160]!r}")
