# -*- coding: utf-8 -*-
import io
f="app.py"; s=io.open(f,encoding="utf-8").read()
s=s.replace("\u7701\u6d41\u7248\u4e09\u6863\", 66)", "\u7701\u6d41\u7248\u4e09\u6863\", 64)",1)
io.open(f,"w",encoding="utf-8").write(s)
print("66->64")
