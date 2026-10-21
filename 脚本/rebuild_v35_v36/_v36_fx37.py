# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch3.py"
s = io.open(p, encoding="utf-8").read()

a = '''    for i, (v, cap, col) in enumerate(N):
        x = ML + i * 2.62
        text(s, x, top, 2.50, 0.44, v, 26, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 0.50, 2.50, 0.24, cap, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    # 参考函
    text(s, X2 - 6.90, top - 0.02, 6.90, 0.24, "参考函原文（设计后服务函 第 1 页第 8 项）", 10, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    frame(s, X2 - 6.90, top + 0.24, 6.90, 0.78, LINE, fill=CARD)
    pic(s, "s037_1.png", X2 - 6.90, top + 0.24, 6.90, 0.78, mode="fit")'''
b = '''    for i, (v, cap, col) in enumerate(N):
        x = ML + i * 2.04
        text(s, x, top, 1.96, 0.44, v, 25, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 0.50, 1.96, 0.24, cap, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    # 参考函
    RX = 6.98
    text(s, RX, top - 0.02, X2 - RX, 0.24, "参考函原文（设计后服务函 第 1 页第 8 项）", 10, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    pic(s, "s037_1.png", RX, top + 0.26, X2 - RX, 0.66, mode="fit")'''
assert a in s
s = s.replace(a, b)
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
