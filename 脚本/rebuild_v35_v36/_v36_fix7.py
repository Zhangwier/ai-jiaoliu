# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:50].replace("\n", "\\n"))
    s = s.replace(a, b)

# p15 用表格返回值定位下方文字
rep('''    table(s, ML, 2.02, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=13, head_size=13,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], minrow=0.38,
          head_h=0.42, first_bold=False)
    text(s, ML, 3.86, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 4.26, 6.30, 1.0, LINE)
    text(s, ML, 4.38, 6.30, 0.26, "三个机制叠在一起，错就被盖住了", 12, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)''',
    '''    ty = table(s, ML, 2.02, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=12.5, head_size=12.5,
               aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT],
               minrow=0.34, head_h=0.38, first_bold=False, pad=0.10, ls=1.1)
    text(s, ML, ty + 0.10, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, ty + 0.50, 6.30, 1.0, LINE)
    text(s, ML, ty + 0.62, 6.30, 0.26, "三个机制叠在一起，错就被盖住了", 12, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)''')
rep('''    yy = 4.74
    for i, (t, b) in enumerate(M):''', '''    yy = ty + 0.98
    for i, (t, b) in enumerate(M):''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
