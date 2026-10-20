import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
# 已重编号后，反推 old->new：用之前 save 前的 changed 不可得，改由人工推算表
# 已知新增：7(总览)、20(三来源)、69(Muse)、62→(B1c 挪到 62)、65(B2c)、82(C1b)、86(C4b)、113(神经网络移入F)
for n in (20,60,63,66,68,70,71,73,75,76,77,78,79,58,87,92,120):
    s = prs.slides[n-1]
    for sh in s.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text
            if re.search(r"[P第]\s*\d", t):
                print(f"s{n} [{sh.name}] {t[:120]}")
