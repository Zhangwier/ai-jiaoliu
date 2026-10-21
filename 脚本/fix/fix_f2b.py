import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s = prs.slides[112]
pn = [sh for sh in s.shapes if sh.has_text_frame and sh.left/360000>29 and sh.top/360000<1.6]
# 去重：保留一个
for extra in pn[1:]:
    extra._element.getparent().remove(extra._element)
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("01  认识 AI"):
        settext(sh, "附录  速查", size=12)
        break
save(prs); print("ok", len(pn))
