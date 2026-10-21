# -*- coding: utf-8 -*-
import sys, importlib
sys.path.insert(0, r"D:\MG\program\ai交流\_v35")
import ds
from ds import *
import p_open, p_ch1, p_ch2, p_ch3, p_ch4

OUT = r"D:\MG\program\ai交流\_v35\build.pptx"
prs = new_deck()
p_open.build(prs)
p_ch1.build(prs, 6)
p_ch2.build(prs, 18)
p_ch3.build(prs, 28)
p_ch4.build(prs, 39)
prs.save(OUT)
print("slides:", len(prs.slides))
for w in ds.WARN: print(w)
