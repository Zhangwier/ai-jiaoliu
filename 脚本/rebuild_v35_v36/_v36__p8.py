# -*- coding: utf-8 -*-
import io
f="app.py"; s=io.open(f,encoding="utf-8").read()
reps=[
 ("\u9644\u5f55\uff1a\u516d\u5757\u901f\u67e5\uff0c\u4e00\u6761\u6761\u67e5\u5f97\u5230", "\u9644\u5f55\uff1a\u4e03\u5757\u901f\u67e5\uff0c\u4e00\u6761\u6761\u67e5\u5f97\u5230"),
 ("\u672c\u6587\u4e0e\u9644\u5f55\u91cc\u7684\u6bcf\u4e2a\u6570\u5b57", "\u672c\u6587\u91cc\u6bcf\u4e2a\u6570\u5b57"),
 ("    yy = y + 0.18\n    for tag, nm, desc, pg in IT:\n        h = 0.62", "    yy = y + 0.02\n    for tag, nm, desc, pg in IT:\n        h = 0.56"),
 ('text(s, ML, yy + 0.09, 0.60, 0.46, tag, 30,', 'text(s, ML, yy + 0.07, 0.60, 0.44, tag, 27,'),
 ('text(s, ML + 0.72, yy + 0.13, 2.20, 0.32, nm, 17,', 'text(s, ML + 0.72, yy + 0.12, 2.20, 0.30, nm, 16,'),
 ('text(s, ML + 3.02, yy + 0.17, 7.30, 0.30, desc, 12.5,', 'text(s, ML + 3.02, yy + 0.15, 7.30, 0.28, desc, 12,'),
 ('text(s, X2 - 0.80, yy + 0.17, 0.80, 0.28, f"P{pg}", 11.5,', 'text(s, X2 - 0.80, yy + 0.15, 0.80, 0.26, f"P{pg}", 11,'),
 ('text(s, ML, yy + 0.22, CW, 0.30,\n         "\u8bfb\u6cd5\uff1a', 'text(s, ML, yy + 0.12, CW, 0.28,\n         "\u8bfb\u6cd5\uff1a'),
]
for a,b in reps:
    if a not in s: print("MISS", a[:44]); continue
    s=s.replace(a,b,1)
io.open(f,"w",encoding="utf-8").write(s)
print("ok")
