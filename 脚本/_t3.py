import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
for n in (85,86,87):
    s=prs.slides[n-1]
    print("="*8,n)
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            print(f"  t{sh.top/EMU:5.2f} | {sh.text_frame.text.strip()[:90].replace(chr(10),' / ')}")
