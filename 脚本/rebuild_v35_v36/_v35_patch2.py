# -*- coding: utf-8 -*-
import sys
p = "ds.py"
s = open(p, encoding="utf-8").read()

# 1) table: strip tableStyleId so no theme fill/band
s = s.replace("""    tblPr = tbl._tbl.tblPr
    for a in ('firstRow','bandRow','bandCol','firstCol','lastRow','lastCol'):
        tblPr.set(a, '0')""",
"""    tblPr = tbl._tbl.tblPr
    for a in ('firstRow','bandRow','bandCol','firstCol','lastRow','lastCol'):
        tblPr.set(a, '0')
    for e in tblPr.findall(qn('a:tableStyleId')):
        tblPr.remove(e)""")

# 2) explicit background fill for every cell (inherit = paper)
s = s.replace("""            if head and ri == 0 and headfill is not None:
                cell.fill.solid(); cell.fill.fore_color.rgb = headfill
            elif zebra is not None and ri % 2 == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = zebra""",
"""            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD
            if head and ri == 0 and headfill is not None:
                cell.fill.fore_color.rgb = headfill
            elif zebra is not None and ri % 2 == 0:
                cell.fill.fore_color.rgb = zebra""")

# 3) pic: optional hairline frame
s = s.replace("""def pic(s, fn, x, y, w=None, h=None, mode="fit", path=None, shadow=False):""",
"""def frame(s, x, y, w, h, color=RULE, lw=0.75, fill=None):
    r = box(s, x, y, w, h, fill=fill, line=color, lw=lw)
    return r


def pic(s, fn, x, y, w=None, h=None, mode="fit", path=None, shadow=False, border=None):""")
s = s.replace("""        else:
            h2 = w / ar
            sh = s.shapes.add_picture(p, I(x), I(y + (h - h2)/2), I(w), I(h2))""",
"""        else:
            h2 = w / ar
            sh = s.shapes.add_picture(p, I(x), I(y + (h - h2)/2), I(w), I(h2))""")
open(p, "w", encoding="utf-8").write(s)
print("ok")
