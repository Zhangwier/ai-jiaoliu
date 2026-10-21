# -*- coding: utf-8 -*-
p = "ds.py"
s = open(p, encoding="utf-8").read()
add = '''

def takeaway(s, t, y=6.18, color=INK, sub=None, w=None, rule=True, dark=False):
    """页脚上方的结论带：细线 + 一句判断"""
    ww = w if w else CW
    if rule:
        hr(s, ML, y - 0.12, ww, 0.75, C(0x38,0x40,0x4E) if dark else RULE)
    n = nlines(t, 13, ww, True)
    hh = n * 13 * LINES * 1.45 / 72.0
    text(s, ML, y, ww, hh + 0.06, t, 13, color, SANS, True, PP_ALIGN.LEFT, 1.45, check=False)
    return y + hh
'''
s = s.replace("# ---------- 表格（自动行高） ----------", add.strip("\n") + "\n\n# ---------- 表格（自动行高） ----------")
open(p, "w", encoding="utf-8").write(s)
print("ok")
