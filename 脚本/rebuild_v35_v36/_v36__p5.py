# -*- coding: utf-8 -*-
import io
f = "app2.py"
s = io.open(f, encoding="utf-8").read()
reps = [
 ("yy = y + 0.04 + (i if left else i - 6) * 0.735", "yy = y + 0.04 + (i if left else i - 6) * 0.700"),
 ('text(s, x + 0.02, yy + 0.36, cw - 0.04, 0.22, good, 10,', 'text(s, x + 0.02, yy + 0.35, cw - 0.04, 0.21, good, 9.8,'),
 ('text(s, x + 0.02, yy + 0.58, cw - 0.04, 0.22, "\u8fb9\u754c\uff1a" + bad, 10, ORANGE, SANS, False,', 'text(s, x + 0.02, yy + 0.55, cw - 0.04, 0.21, "\u8fb9\u754c\uff1a" + bad, 9.6, ORANGE, SANS, False,'),
 ('hr(s, LX, y + 0.04 + 6 * 0.735, cw, 1.0, LINE)\n    hr(s, RX, y + 0.04 + 5 * 0.735, cw, 1.0, LINE)', 'hr(s, LX, y + 0.04 + 6 * 0.700, cw, 1.0, LINE)\n    hr(s, RX, y + 0.04 + 5 * 0.700, cw, 1.0, LINE)'),
 ('text(s, ML, 6.42, CW, 0.28,', 'text(s, ML, 6.50, CW, 0.26,'),
 ('11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)\n    footer(s, "\u6765\u6e90\uff1a\u5404\u516c\u53f8\u6a21\u578b\u9875\u4e0e\u516c\u544a\u3001artificialanalysis.ai', '10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)\n    footer(s, "\u6765\u6e90\uff1a\u5404\u516c\u53f8\u6a21\u578b\u9875\u4e0e\u516c\u544a\u3001artificialanalysis.ai'),
]
for a,b in reps:
    if a not in s:
        print("MISS:", a[:50]); continue
    s = s.replace(a,b,1)
io.open(f,"w",encoding="utf-8").write(s)
print("done")
