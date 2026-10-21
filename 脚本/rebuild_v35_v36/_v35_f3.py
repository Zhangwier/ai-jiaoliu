# -*- coding: utf-8 -*-
p = "p_ch2.py"
s = open(p, encoding="utf-8").read()
s = s.replace('tbl_rows.append([r[0], r[1], (r[2], GRN, 10.5, True), r[3]])',
              'tbl_rows.append([r[0], r[1], ((r[2], GRN, 10.5, True),), r[3]])')
open(p, "w", encoding="utf-8").write(s)
print("ok")
