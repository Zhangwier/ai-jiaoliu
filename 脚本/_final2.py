import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
# 检查各章分隔页的"这一部分"列条与页数一致性
checks = {59:("B1–B13",20), 80:("C1–C5",7), 93:("E1",16), 110:("F1",10), 54:("A1–A4",4), 121:("Z1",None)}
for n in checks:
    s=prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame and ("页" in sh.text_frame.text and ("共" in sh.text_frame.text or "（" in sh.text_frame.text)):
            t=sh.text_frame.text.replace("\n"," / ")
            print(n, "|", t[:200])
