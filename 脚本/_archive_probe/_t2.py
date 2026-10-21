import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for n in (53,59,80,88,93,110,121,52,5):
    s=prs.slides[n-1]
    print("="*10,n)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.top/EMU>2:
            print(f"  t{sh.top/EMU:5.2f} | {sh.text_frame.text.strip()[:150].replace(chr(10),' / ')}")
