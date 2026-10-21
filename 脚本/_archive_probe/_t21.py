import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs(); EMU=360000
kw=["开源平台对照","B1–B13","C1–C7","F1–F9","十七页","九页","十六页"]
for i,s in enumerate(prs.slides):
    if i>=124: break
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        for k in kw:
            if k in t:
                print(f"s{i+1:3d} {sh.name:14s} [{k}] {t.strip()[:130]!r}")
