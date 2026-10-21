# -*- coding: utf-8 -*-
import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
# 1) 修正 6 个新加页的页码为实际序号
FIX = {17:17, 20:20, 69:69, 82:82, 86:86, 113:113}
def pagenum_shape(s):
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        if sh.left/360000 > 29 and sh.top/360000 < 1.5:
            t = sh.text_frame.text.strip()
            if re.fullmatch(r"\d+", t): return sh
    return None
# 先采集映射（修正前）
mapping = {}
dup = {}
for i, sl in enumerate(prs.slides, start=1):
    sh = pagenum_shape(sl)
    if sh is None: continue
    old = int(sh.text_frame.text.strip())
    if i in FIX:
        continue
    if old in mapping and mapping[old] != i:
        dup.setdefault(old, [mapping[old]]).append(i)
    mapping[old] = i
print("dup", dup)
print("mapping sample", {k:mapping[k] for k in sorted(mapping)[:20]})
print("count", len(mapping), "max old", max(mapping))
import json
json.dump(mapping, open("_pmap.json","w"))
