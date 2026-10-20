import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
mx=0
for i,s in enumerate(prs.slides,1):
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        for m in re.finditer(r"P(\d+)", sh.text_frame.text):
            mx=max(mx,int(m.group(1)))
        for m in re.finditer(r"第\s*(\d+)\s*页", sh.text_frame.text):
            if int(m.group(1))>mx: mx=int(m.group(1)); print("第N页 ref", i, m.group(0), sh.text_frame.text[:50])
print("max ref", mx)
s=prs.slides[108]
print("pics:", [sh.name for sh in s.shapes if sh.shape_type==13])
print("layout:", s.slide_layout.name)
