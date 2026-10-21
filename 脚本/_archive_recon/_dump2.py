import sys, re, io
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
out=[]
for i,s in enumerate(prs.slides,1):
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(f"{i}\t{sh.text_frame.text.strip()}")
open(r"D:\MG\program\ai交流\成品\_work29\all2.txt","w",encoding="utf-8").write("\n".join(out))
print(len(out))
