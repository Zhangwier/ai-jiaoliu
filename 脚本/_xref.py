# -*- coding: utf-8 -*-
import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
pat_page = re.compile(r"第\s*\d+\s*页")
pat_P = re.compile(r"P\d+(?:\s*[–\-—]\s*P?\d+)?")
for i, sl in enumerate(prs.slides, start=1):
    for sh in sl.shapes:
        if not sh.has_text_frame: continue
        t = sh.text_frame.text
        if pat_page.search(t) or pat_P.search(t):
            hits = pat_page.findall(t) + pat_P.findall(t)
            print(f"s{i:3d} [{sh.name}] {hits} :: {t[:70]}")
