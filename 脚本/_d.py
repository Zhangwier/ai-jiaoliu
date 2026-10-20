import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
for n in (6,18,28,39):
    s=prs.slides[n-1]
    print("="*8,n)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.left/360000>20:
            print(f"  t{sh.top/360000:6.2f} | {sh.text_frame.text[:70].replace(chr(10),'/')}")
