import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); EMU=360000
s=prs.slides[79]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("价格与地区政策以官方页面为准"):
        settext(sh, "价格与地区政策以官方页面为准。C1–C5（含 C1b、C4b）共 7 页：四件事自查、照着做七步、地区与支付、注册、价格、省钱的路子、省流版。")
        print("ok")
save(prs)
