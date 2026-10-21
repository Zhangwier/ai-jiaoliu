# -*- coding: utf-8 -*-
import io, re, sys
P = r"D:\MG\program\ai交流\_v35\p_ch3.py"
NEW = r"D:\MG\program\ai交流\_v35\new3031.py"
src = open(P, encoding="utf-8").read()
new = open(NEW, encoding="utf-8").read()

a = src.find("def p30")
b = src.find("# ---------------------------------------------------------------- 32")
if a < 0 or b < 0 or a > b:
    raise SystemExit("ANCHOR FAIL")
out = src[:a] + new + src[b:]
open(P, "w", encoding="utf-8").write(out)
print("patched", len(src), "->", len(out))
