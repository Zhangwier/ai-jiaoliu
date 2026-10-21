import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
def has_pn(s):
    for sh in s.shapes:
        if sh.has_text_frame and sh.left/360000>29 and sh.top/360000<1.6:
            t = sh.text_frame.text.strip()
            if re.fullmatch(r"\d+", t): return t
    return None
missing=[]
for i, sl in enumerate(prs.slides, start=1):
    pn = has_pn(sl)
    if pn is None:
        ttl=""
        for sh in sl.shapes:
            if sh.has_text_frame and sh.name=="TextBox 4": ttl=sh.text_frame.text.replace("\n","/")[:40]
        missing.append((i,ttl))
for m in missing: print(m)
