# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

def rep(a, b):
    global s
    if a not in s:
        print("MISS:", a[:60].replace("\n", "\\n"))
    s = s.replace(a, b)

# ---------- P16：图放大，右栏左移 ----------
rep('''    pic(s, "s016_1.png", ML, 2.02, 7.20, 3.06, mode="fit")
    text(s, ML, 5.18, 7.20, 0.26,
         "横轴：输入长度（token，对数）；纵轴：归一化编辑距离得分 —— 越低说明越用不上前面的内容", 10.5,
         FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    RX = 8.10
    vr(s, RX - 0.34, 2.02, 3.10, 1.0, LINE)''',
    '''    pic(s, "s016_1.png", ML, 1.98, 6.62, 3.34, mode="fit")
    text(s, ML, 5.44, 6.70, 0.26,
         "横轴：输入长度（token，对数）；纵轴：归一化编辑距离得分 —— 越低说明越用不上前面给的词", 10.5,
         FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    RX = 7.52
    vr(s, RX - 0.36, 1.98, 3.30, 1.0, LINE)''')
rep('''    text(s, RX, 4.40, 4.58, 0.72,
         "《Lost in the Middle》：关键信息放在中段，比放在开头或结尾更容易被漏掉。", 12.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45)''',
    '''    text(s, RX, 4.40, 5.12, 0.86,
         "《Lost in the Middle》：关键信息放在中段，比放在开头或结尾更容易被漏掉。", 12.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45)''')

# ---------- P14：下半部补足 ----------
rep('''    box(s, RX, 3.92, 5.70, 2.06, fill=MINT)
    text(s, RX + 0.24, 4.04, 5.20, 0.22, "开放的问法（本稿后来改用的）", 10.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, RX + 0.24, 4.30, 5.22, 1.50,
         "“请判断这份图纸内部是否存在不一致，没有也可以；\\n每一项写明位置、依据原文，以及你的把握程度；\\n找不到依据的写“待核”，不要补。”",
         13.5, TEAL, SANS, False, PP_ALIGN.LEFT, 1.55)''',
    '''    box(s, RX, 3.92, 5.70, 2.02, fill=MINT)
    text(s, RX + 0.24, 4.02, 5.20, 0.22, "开放的问法（本稿后来改用的）", 10.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, RX + 0.24, 4.26, 5.22, 1.54,
         "“请判断这份图纸内部是否存在不一致，没有也可以；\\n每一项写明位置、依据原文，以及你的把握程度；\\n找不到依据的写“待核”，不要补。”",
         13.5, TEAL, SANS, False, PP_ALIGN.LEFT, 1.55)
    # 底部三条对照
    hr(s, ML, 5.96, CW, 1.0, LINE)
    OUT = [("提问里带答案", "它给的是你想听的", ORANGE), ("提问里不带倾向", "它才可能说“没有”", TEAL),
           ("要求写把握程度", "它会把不确定的标出来", C(0x4A, 0x6C, 0x9E))]
    for i, (t, b, col) in enumerate(OUT):
        x = ML + i * 4.05
        vr(s, x, 6.10, 0.46, 2.4, col)
        text(s, x + 0.18, 6.06, 3.70, 0.26, t, 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x + 0.18, 6.32, 3.70, 0.26, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)''')

io.open(p, "w", encoding="utf-8").write(s)
print("done")
