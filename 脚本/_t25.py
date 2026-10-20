import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs=openprs()
bad=["会后","会后可查","一页一家公司","案例、风险与破限","降智测试与破限","开源平台对照","十七页","九页","十六页","十五页","七页：","四页："]
for i,s in enumerate(prs.slides):
    if i>=124: break
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        for b in bad:
            if b in t:
                print(f"s{i+1:3d} {sh.name:14s} [{b}] {t.strip()[:90]!r}")
print("scan done")
