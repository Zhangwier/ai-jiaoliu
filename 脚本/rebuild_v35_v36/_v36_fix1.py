# -*- coding: utf-8 -*-
import io, sys
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

# p08 两段扩高
rep('text(s, ML, 5.76, 5.50, 0.96,', 'text(s, ML, 5.72, 5.50, 1.18,')
rep('text(s, 6.60, 5.76, 6.10, 0.96,', 'text(s, 6.60, 5.72, 6.10, 1.18,')
# p09 结论块
rep('text(s, RX + 0.22, 5.82, 5.46, 0.52,', 'text(s, RX + 0.20, 5.76, 5.50, 0.68,')
# p11 说明区
rep('text(s, ML, 5.72, 5.60, 0.86,', 'text(s, ML, 5.68, 5.62, 1.12,')
rep('text(s, 6.92, 5.72, 5.52, 0.86,', 'text(s, 6.92, 5.68, 5.52, 1.12,')
# p15 三机制
rep('text(s, ML + 1.70, yy, 4.60, 0.42, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.4)',
    'text(s, ML + 1.66, yy, 4.64, 0.52, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.36, check=False)')
# p15 右侧
rep('text(s, RX, 3.16, 5.00, 0.52,', 'text(s, RX, 3.14, 5.00, 0.72,')
rep('text(s, RX, 4.28, 5.00, 0.90,', 'text(s, RX, 4.24, 5.00, 1.06,')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
