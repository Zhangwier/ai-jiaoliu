import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
cur=None
for i, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame and "C4b" in sh.text_frame.text:
            cur=i; break
    if cur is not None: break
move_to(prs, prs.slides[cur], 85)   # 0-based 85 -> 第 86 页
save(prs)
for i, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame and (sh.text_frame.text.startswith(("C1 ","C1b","C2 ","C3 ","C4 ","C4b","C7 ","D1"))):
            print(i+1, sh.text_frame.text[:44])
