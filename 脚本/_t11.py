import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
for i,s in enumerate(prs.slides,1):
    for sh in s.shapes:
        if sh.has_text_frame:
            t=sh.text_frame.text
            if "E1–E4" in t or "开源平台对照" in t or "个分页" in t or "共 1" in t or "共 4" in t or "共 7" in t or "共 10" in t:
                print(i, "|", t[:120].replace("\n"," / "))
