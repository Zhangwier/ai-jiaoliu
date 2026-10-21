import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
# 1) C7 -> C5 标题
s = prs.slides[86]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("C7"):
        old = sh.text_frame.text
        settext(sh, old.replace("C7", "C5", 1), size=26)
        print("C7->C5:", sh.text_frame.text[:60])
# 2) s52 引用
s = prs.slides[51]
for sh in s.shapes:
    if sh.has_text_frame and "C1–C7" in sh.text_frame.text:
        settext(sh, sh.text_frame.text.replace("C1–C7", "C1–C5"), size=17)
        print("s52 ok")
# 3) s80 分隔页副标
s = prs.slides[79]
for sh in s.shapes:
    if sh.has_text_frame and "C1–C4b、C7" in sh.text_frame.text:
        settext(sh, "价格与地区政策以官方页面为准。C1–C5 共 7 页：四件事自查、照着做七步、地区与支付、注册、价格、省钱的路子、省流版。", size=14)
        print("s80 ok")
save(prs); print("saved")
