import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs()
for i,s in enumerate(prs.slides):
    if i>=124: break
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        if "E4" in t or "开源平台" in t:
            print(f"s{i+1:3d} {sh.name:14s} {t.strip()[:170]!r}")
