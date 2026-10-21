# -*- coding: utf-8 -*-
p = "mk.py"
s = open(p, encoding="utf-8").read()
s = s.replace("import p_open, p_ch1", "import p_open, p_ch1, p_ch2")
s = s.replace("p_ch1.build(prs, 6)", "p_ch1.build(prs, 6)\np_ch2.build(prs, 18)")
open(p, "w", encoding="utf-8").write(s)
print(s)
