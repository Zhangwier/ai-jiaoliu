import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[108]
for sh in s.shapes:
    if sh.has_text_frame and sh.top/EMU<4.3:
        print(f"t{sh.top/EMU:5.2f} l{sh.left/EMU:5.2f} w{sh.width/EMU:5.2f} h{sh.height/EMU:5.2f} {sh.name:12s} | {sh.text_frame.text.strip()[:60]}")
print("---bottom---")
for sh in s.shapes:
    if sh.has_text_frame and sh.top/EMU>17.5:
        print(sh.name, round(sh.left/EMU,2), round(sh.top/EMU,2), round(sh.width/EMU,2), "|", sh.text_frame.text[:80])
