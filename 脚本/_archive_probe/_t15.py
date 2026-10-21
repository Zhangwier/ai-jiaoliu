import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[73]
for sh in s.shapes:
    if sh.has_text_frame and ("内网单机" in sh.text_frame.text or "0.6B" in sh.text_frame.text):
        print(f"t{sh.top/EMU:5.2f} l{sh.left/EMU:5.2f} | {sh.text_frame.text.strip()[:120]}")
