# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()
a = '''        text(s, fx(shi) + 0.22, yy - 0.10, 1.20, 0.26, f"+{shi-slo} 分", 12.5, ORANGE, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, fx(shi) + 0.22, yy + 0.16, 1.60, 0.24, hi, 10.5, NAVY, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)'''
b = '''        text(s, fx(shi) + 0.24, yy - 0.14, 1.20, 0.26, f"+{shi-slo} 分", 13, ORANGE, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)'''
assert a in s
s = s.replace(a, b)
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
