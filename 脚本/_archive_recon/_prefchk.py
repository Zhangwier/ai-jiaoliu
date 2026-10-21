import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
pat = re.compile(r"(P\d{1,3}|第\s?\d{1,3}\s?页)")
for i, s in enumerate(prs.slides, 1):
    hits=[]
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        for m in pat.finditer(sh.text_frame.text):
            hits.append(m.group(0))
    if hits:
        print(i, "|", title_of(s)[:40], "|", sorted(set(hits)))
