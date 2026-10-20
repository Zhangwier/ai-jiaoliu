import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
hits=[]
for i, s in enumerate(prs.slides, 1):
    for sh in s.shapes:
        if sh.has_text_frame and "会后" in sh.text_frame.text:
            hits.append((i, sh.text_frame.text[:60]))
print("会后:", hits)
# 标签一致性
bad=[]
for i, s in enumerate(prs.slides, 1):
    if i>124: continue
    tb=None
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/360000<3 and sh.top/360000<1.5 and len(sh.text_frame.text)<20:
            tb=sh.text_frame.text.strip()
    if tb is None: bad.append((i,"none"))
print("无左上标签页:", bad)
