# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()

# ---------------- 重写 p23 ----------------
i0 = s.index('def p23(prs, pno):')
i1 = s.index('def p24(prs, pno):')

new = '''def p23(prs, pno):
    """AA 智能指数散点：横轴每题成本（对数），纵轴指数"""
    import math
    s = slide(prs)
    y = head(s, CH, pno, "分数和价格不成正比：贵的不一定更强",
             "横轴是每道题的成本（对数刻度），纵轴是 AA 智能指数。越靠左上，性价比越高。")
    # 只标 8 个关键点，其余作背景群
    KEY = [("Opus 5.5", 58, 0.72, "L"), ("Sonnet 5.5", 56, 0.36, "L"),
           ("GPT-6 Astra", 53, 1.05, "R"), ("GPT-6.1 Sol", 52, 0.72, "R"),
           ("GLM-5.3", 45, 0.09, "L"), ("DeepSeek V4.1", 39, 0.02, "R"),
           ("MiniMax-M3", 29, 0.015, "R"), ("GPT-6 Luna", 38, 0.012, "L")]
    BG = [(53, 0.42), (53, 0.58), (52, 0.55), (46, 0.38), (46, 0.12), (45, 0.16),
          (44, 0.08), (44, 0.11), (43, 0.14), (42, 0.03), (41, 0.05), (40, 0.04)]
    x0, x1 = ML + 0.72, 8.94
    y0, y1 = 5.42, 2.26

    def fx(v):
        return x0 + (math.log10(v) + 2.05) / 3.05 * (x1 - x0)

    def fy(v):
        return y0 - (v - 24) / (62 - 24) * (y0 - y1)

    for g in [26, 34, 42, 50, 58]:
        hr(s, x0 - 0.08, fy(g), x1 - x0 + 0.12, 0.6, C(0xE9, 0xE4, 0xD8))
        text(s, x0 - 0.74, fy(g) - 0.11, 0.62, 0.22, str(g), 9.5, FAINT, SANS, False,
             PP_ALIGN.RIGHT, 1.0, check=False)
    for g in [0.01, 0.1, 1.0]:
        vr(s, fx(g), y1 - 0.08, y0 - y1 + 0.12, 0.6, C(0xE9, 0xE4, 0xD8))
        text(s, fx(g) - 0.34, y0 + 0.08, 0.68, 0.22, f"${g:g}", 9.5, FAINT, SANS, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    vr(s, x0 - 0.08, y1 - 0.12, y0 - y1 + 0.16, 0.8, LINE2)
    hr(s, x0 - 0.08, y0, x1 - x0 + 0.12, 0.8, LINE2)
    text(s, x0 - 0.74, y1 - 0.32, 1.40, 0.22, "智能指数", 9.5, GREY2, SANS, True, PP_ALIGN.LEFT,
         1.0, check=False)
    text(s, x1 - 1.30, y0 + 0.36, 1.40, 0.22, "每题成本（美元）", 9.5, GREY2, SANS, True,
         PP_ALIGN.RIGHT, 1.0, check=False)
    for sc, cost in BG:
        circ(s, fx(cost), fy(sc), 0.105, C(0xB9, 0xC8, 0xC4))
    for nm, sc, cost, side in KEY:
        cx, cy = fx(cost), fy(sc)
        hi = nm in ("Opus 5.5", "GPT-6.1 Sol")
        circ(s, cx, cy, 0.14 if not hi else 0.17, ORANGE if hi else TEAL)
        w = est_w(nm, 9.5, hi) / 72.0 + 0.08
        lx = cx + 0.13 if side == "R" else cx - 0.13 - w
        text(s, lx, cy - 0.095, w + 0.04, 0.19, nm, 9.5, ORANGE if hi else C(0x33, 0x3D, 0x4D),
             SANS, hi, PP_ALIGN.LEFT, 1.0, check=False)

    RX = 9.50
    vr(s, RX - 0.32, 2.20, 3.28, 1.0, LINE)
    text(s, RX, 2.20, 3.20, 0.24, "怎么读这张图", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    RD = ["左上角：分高、便宜 —— 性价比最好",
          "右下角：分低、还贵 —— 基本可排除",
          "同一条竖线：同等价位谁分更高",
          "价格差 10 倍，分数可能只差 3 分"]
    yy = 2.52
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
         PP_ALIGN.LEFT, 1.35)
    footer(s, "来源：AA 智能指数 v4.3.2（每道题成本为公开定价换算，美元）")
    return s


'''
s = s[:i0] + new + s[i1:]

# ---------------- p22：去掉重叠的 ×，压表格 ----------------
s = s.replace('''    text(s, MID - 0.30, ty + 2.30, 0.90, 0.90, "×", 46, LINE2, SANS, True, PP_ALIGN.CENTER, 1.0,
         check=False)
    RX = 6.92''', '''    vr(s, 6.62, ty + 0.10, 3.30, 1.0, LINE)
    text(s, 6.40, ty - 0.02, 0.46, 0.30, "×", 22, LINE2, SANS, True, PP_ALIGN.CENTER, 1.0,
         check=False)
    RX = 6.92''')
s = s.replace('''    table(s, RX, ty + 0.32, ML + CW - RX, rows, [1.50, 2.40, 2.30], size=10.5, head_size=10.5,
          minrow=0.50, head_h=0.30, first_bold=True, pad=0.10, ls=1.15,''',
              '''    table(s, RX, ty + 0.32, ML + CW - RX, rows, [1.44, 2.44, 2.32], size=10, head_size=10,
          minrow=0.42, head_h=0.28, first_bold=True, pad=0.09, ls=1.1,''')

io.open(p, "w", encoding="utf-8").write(s)
print("ok")
