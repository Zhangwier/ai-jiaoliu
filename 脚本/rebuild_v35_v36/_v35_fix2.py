# -*- coding: utf-8 -*-
p = "p_ch1.py"
s = open(p, encoding="utf-8").read()
s = s.replace('table(s, ML, top + 0.28, 4.55,', 'table(s, ML, top + 0.28, 5.15,')
s = s.replace('    text(s, ML, 4.06, 4.55, 0.28,', '    text(s, ML, 4.06, 5.15, 0.28,')
open(p, "w", encoding="utf-8").write(s)
print("ok")
