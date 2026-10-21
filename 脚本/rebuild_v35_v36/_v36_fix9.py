# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

rep('''    y = head(s, CH, pno, "编造不是故障，是“挑最像的下一个字”的副作用",
             "它没有“查证”这个动作。材料里没有，它就按最像的样子造一个 —— 语感越好，越像真的。")
    HDR = [("时间", ML, 1.30), ("谁", ML + 1.36, 3.60), ("编了什么", ML + 5.02, 3.10),
           ("代价", ML + 8.24, 2.90), ("出处", X2 - 0.90, 0.90)]
    for t, x, w in HDR:
        al = PP_ALIGN.RIGHT if t == "出处" else PP_ALIGN.LEFT
        text(s, x, 1.92, w, 0.24, t, 10.5, FAINT, SANS, True, al, 1.0, cs=1.0, check=False)
    hr(s, ML, 2.20, CW, 1.0, LINE2)''',
    '''    y = head(s, CH, pno, "编造不是故障，是“挑最像的下一个字”的副作用",
             "它没有“查证”这个动作。材料里没有，它就按最像的样子造一个 —— 语感越好，越像真的。")
    ty = y + 0.20
    HDR = [("时间", ML, 1.30), ("谁", ML + 1.36, 3.60), ("编了什么", ML + 5.02, 3.10),
           ("代价", ML + 8.24, 2.90), ("出处", X2 - 0.90, 0.90)]
    for t, x, w in HDR:
        al = PP_ALIGN.RIGHT if t == "出处" else PP_ALIGN.LEFT
        text(s, x, ty, w, 0.24, t, 10.5, FAINT, SANS, True, al, 1.0, cs=1.0, check=False)
    hr(s, ML, ty + 0.28, CW, 1.0, LINE2)''')
rep('''        yy = 2.30 + i * 0.86''', '''        yy = ty + 0.40 + i * 0.86''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
