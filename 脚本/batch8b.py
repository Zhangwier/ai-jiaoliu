import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[66]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("Grok 已并入 X 订阅"):
        print("BEFORE", sh.name, sh.top/EMU, repr(sh.text_frame.text))
        settext(sh, "中转与代理池（E2–E2f）；支持地区（C2）；Grok 并入 X 订阅（C4）；省流版（C5）。")
        print("AFTER", repr(sh.text_frame.text), sh.text_frame.paragraphs[0].runs[0].font.size)
save(prs)
