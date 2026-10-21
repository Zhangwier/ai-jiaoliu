# -*- coding: utf-8 -*-
p = "p_ch1.py"
s = open(p, encoding="utf-8").read()

# P8 竖条缩短、间距加大
s = s.replace('        box(s, x0, yy, 0.04, 1.18, fill=col)', '        box(s, x0, yy, 0.035, 1.02, fill=col)')
s = s.replace('            line(s, x0 + 0.02, yy + 1.18, x0 + 0.02, yy + 1.44, RULE2, 1.0, dash=None)',
              '            line(s, x0 + 0.017, yy + 1.02, x0 + 0.017, yy + 1.52, RULE2, 0.9)')
s = s.replace('        yy += 1.44\n', '        yy += 1.52\n')

# P15 表格与注记
s = s.replace('    table(s, ML, top + 0.30, 4.55, rows, [1.0, 1.15, 1.05], size=10.5, head_size=10,\n'
              '          first_bold=True, headfill=None, aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT])\n'
              '    text(s, ML, 3.94, 4.55, 0.28,',
              '    table(s, ML, top + 0.28, 4.55, rows, [1.0, 1.15, 1.05], size=10.5, head_size=10,\n'
              '          first_bold=True, headfill=None, aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT],\n'
              '          minrow=0.28, ls=1.18)\n    text(s, ML, 4.06, 4.55, 0.28,')
s = s.replace('    yy = 4.36\n    for ttl, desc in mech:', '    yy = 4.52\n    for ttl, desc in mech:')
s = s.replace('        yy += 0.62\n', '        yy += 0.60\n')

# P16 图加白底卡片
s = s.replace('    pic(s, "s016_1.png", ML, 2.28, 6.10, 4.06, mode="fit")',
              '    frame(s, ML, 2.28, 6.30, 4.14, RULE, 0.75, fill=CARD)\n'
              '    pic(s, "s016_1.png", ML + 0.10, 2.38, 6.10, 3.94, mode="fit")')
open(p, "w", encoding="utf-8").write(s)
print("ok")
