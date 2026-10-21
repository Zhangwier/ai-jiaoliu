import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s5=prs.slides[4]
for sh in s5.shapes:
    if sh.has_text_frame and sh.left/EMU>22 and sh.top/EMU>6:
        print("### s5", sh.name, sh.top/EMU)
        for p in sh.text_frame.paragraphs:
            print("   |", repr(p.text), [r.font.size.pt if r.font.size else None for r in p.runs])
s=prs.slides[108]
print("="*10,"s109 E4")
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip():
        print(f"t{sh.top/EMU:5.2f} l{sh.left/EMU:5.2f} w{sh.width/EMU:5.2f} h{sh.height/EMU:5.2f} | {sh.text_frame.text.strip()[:300].replace(chr(10),' / ')}")
