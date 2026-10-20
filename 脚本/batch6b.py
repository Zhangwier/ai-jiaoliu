import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
# 找到 C4b 页
cur = None
for i, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame and "C4b" in sh.text_frame.text:
            cur = i; break
    if cur is not None: break
print("C4b now at", cur+1)
move_to(prs, prs.slides[cur], 86)
save(prs)
for i, sl in enumerate(prs.slides):
    for sh in sl.shapes:
        if sh.has_text_frame and ("C4b" in sh.text_frame.text or "C4  " in sh.text_frame.text or "C7" in sh.text_frame.text):
            print(i+1, sh.text_frame.text[:40])
