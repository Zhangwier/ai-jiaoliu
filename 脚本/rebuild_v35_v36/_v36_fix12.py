# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

rep('''    KEY = [("Opus 5.5", 58, 0.72, "L"), ("Sonnet 5.5", 56, 0.36, "L"),
           ("GPT-6 Astra", 53, 1.05, "R"), ("GPT-6.1 Sol", 52, 0.72, "R"),''',
    '''    KEY = [("Opus 5.5", 58, 0.72, "L"), ("Sonnet 5.5", 56, 0.36, "L"),
           ("GPT-6 Astra", 53, 1.05, "R"), ("GPT-6.1 Sol", 52, 0.72, "L"),''')
rep('''    text(s, x0 - 0.74, y1 - 0.32, 1.40, 0.22, "智能指数", 9.5, GREY2, SANS, True, PP_ALIGN.LEFT,
         1.0, check=False)''',
    '''    text(s, x0 + 0.06, y1 - 0.02, 1.40, 0.22, "智能指数", 9.5, GREY2, SANS, True, PP_ALIGN.LEFT,
         1.0, check=False)''')
rep('''    yy = 2.52
    for i, t in enumerate(RD):
        text(s, RX, yy, 0.34, 0.22, f"0{i+1}", 9.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        nl = nlines(t, 11.5, 2.62)
        text(s, RX + 0.36, yy - 0.02, 2.64, nl * 11.5 * 1.32 * 1.42 / 72 + 0.05, t, 11.5, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.42, check=False)
        yy += 0.18 + nl * 11.5 * 1.32 * 1.42 / 72 + 0.14
    box(s, RX, 4.66, 3.18, 0.82, fill=ORANGEL)
    text(s, RX + 0.18, 4.76, 2.84, 0.64,
         "Sol 与 Opus 5.5 同价，分数差 6 分 —— 同价位也要挑。", 11.5, C(0x7A, 0x3A, 0x22), SANS,
         True, PP_ALIGN.LEFT, 1.35)
    text(s, RX, 5.62, 3.18, 0.72,
         "灰点是同一批里的其他模型：\\n多数集中在 40–46 分、成本不到 $0.2。", 10.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.35)''',
    '''    yy = 2.50
    for i, t in enumerate(RD):
        text(s, RX, yy, 0.34, 0.22, f"0{i+1}", 9.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        nl = nlines(t, 11, 2.62)
        text(s, RX + 0.36, yy - 0.02, 2.64, nl * 11 * 1.32 * 1.42 / 72 + 0.05, t, 11, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.42, check=False)
        yy += 0.16 + nl * 11 * 1.32 * 1.42 / 72 + 0.10
    box(s, RX, 4.72, 3.18, 0.72, fill=ORANGEL)
    text(s, RX + 0.18, 4.82, 2.84, 0.56,
         "Sol 与 Opus 5.5 同价，分数差 6 分 —— 同价位也要挑。", 11, C(0x7A, 0x3A, 0x22), SANS,
         True, PP_ALIGN.LEFT, 1.35)
    text(s, RX, 5.56, 3.18, 0.72,
         "灰点是同一批里的其他模型：\\n多数集中在 40–46 分、成本不到 $0.2。", 10.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.35)''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
