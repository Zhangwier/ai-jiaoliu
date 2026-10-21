# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:50].replace("\n", "\\n"))
    s = s.replace(a, b)

# ---- P15 表格压字：缩表、下移说明
rep('''    table(s, ML, 2.02, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=13.5, head_size=13.5,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], minrow=0.44,
          first_bold=False)
    text(s, ML, 3.94, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, 4.30, 6.30, 0.26, "三个机制叠在一起", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)''',
    '''    table(s, ML, 2.04, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=13.5, head_size=13.5,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], minrow=0.42,
          head_h=0.44, first_bold=False)
    text(s, ML, 4.06, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, 4.44, 6.30, 0.26, "三个机制叠在一起", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)''')
rep('''    yy = 4.62
    for i, (t, b) in enumerate(M):''', '''    yy = 4.76
    for i, (t, b) in enumerate(M):''')

# ---- P10 卡片过高：收短，takeaway 上移
rep('''        box(s, x, 2.10, 3.86, 3.34, fill=CARD, line=LINE)
        box(s, x, 2.10, 3.86, 0.055, fill=col)''',
    '''        box(s, x, 2.08, 3.86, 3.14, fill=CARD, line=LINE)
        box(s, x, 2.08, 3.86, 0.055, fill=col)''')
rep('takeaway(s, "所以“让它更听话”和“让它更诚实”不是同一件事：调法一样，目标不一样。", y=5.72,',
    'takeaway(s, "所以“让它更听话”和“让它更诚实”不是同一件事：调法一样，目标不一样。", y=5.56,')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
