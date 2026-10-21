# -*- coding: utf-8 -*-
"""第 01 章 认识 AI"""
from ds import *
from layouts import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "01  认识 AI"


def div(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.24, 1.5, 1.10, "01", 68, C(0x28, 0x3A, 0x59), SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.36, 1.30, 6.0, 0.30, "PART  01", 12, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=2.2, check=False)
    text(s, ML + 1.36, 1.58, 8.0, 0.66, "认识 AI", 38, WHITE, SANS, True, PP_ALIGN.LEFT, 1.14,
         check=False)
    text(s, ML + 1.36, 2.32, 8.4, 0.32, "五个短板，每一个都能从原理上讲清", 16, DARKGREY, SANS,
         False, PP_ALIGN.LEFT, 1.3, check=False)
    hr(s, ML, 2.94, CW, 1.0, C(0x2C, 0x3A, 0x55))
    items = [("01", "五个短板", "抽卡、幻觉、迎合、上下文腐烂、奖励黑客"),
             ("02", "从原理看", "学习三态、注意力、训练三步、逐 token 生成"),
             ("03", "为什么", "它为什么会编、会顺着你说、会忘、会钻空子"),
             ("04", "怎么防", "多跑几次、要出处、开放提问、控制输入、查过程")]
    yy = 3.20
    for no, nm, sub in items:
        text(s, ML, yy + 0.02, 0.62, 0.26, no, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.68, yy - 0.03, 2.50, 0.30, nm, 16, WHITE, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 3.30, yy, 8.10, 0.30, sub, 14, C(0xB6, 0xC0, 0xD0), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.84
    return s


def p07(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "AI 的五个短板：每一个都能从原理上讲清",
             "说“AI 不可靠”没有用。把不可靠拆成五种，每一种都能对应到原理上的一处成因，也都能对应一条处置办法。")
    rows = [["问题", "表现", "来自哪一步", "怎么防"],
            ["抽卡", "同一句话问两次，结果时好时坏", "按概率抽 token", "多跑几次，关键结论复核"],
            ["幻觉", "编出不存在的条款、判例、数字", "只挑“最像”的下一个字", "要出处，回原文核"],
            ["迎合", "你说有错它就找错；你给的前提它照用", "偏好训练奖励“让人满意”", "开放提问，先让它复述口径"],
            ["上下文腐烂", "聊久了忘了前面提过的要求", "注意力被摊薄", "新开对话，要求放开头"],
            ["奖励黑客", "钻空子、走捷径、多要权限", "奖励只看结果指标", "写清红线，过程也查"]]
    table(s, ML, 2.12, CW, rows, [1.42, 4.34, 3.10, 3.20], size=15, head_size=15,
          aligns=[PP_ALIGN.LEFT] * 4, minrow=0.62,
          colors=lambda r, c: ORANGE if c == 0 else (TEAL if c == 3 else C(0x2B, 0x34, 0x45)))
    takeaway(s, "五个短板不是五种“心情”，是五个机制 —— 认出是哪一种，才知道该动哪一处。", y=6.14,
             tag="要点")
    footer(s, "来源：本部分各页原理说明与公开研究；五个短板的分类为本稿归纳")
    return s


def p08(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它学的不是道理，是“下一个字最可能是什么”",
             "大规模语言模型做的事只有一件：给定前面所有字，猜下一个字最可能是什么。学习、压缩、泛化都从这一件事里长出来。")
    ST = [("训练", "拿大量文本，反复调整参数，让“猜下一个字”的把握变大。",
           "这一步只保证：见过的句式更熟，不是懂了因果。", C(0x2C, 0x3A, 0x55)),
          ("预测", "新问题来了，按已学到的概率分布往下猜，一个字一个字往后接。",
           "猜得顺不等于猜得对 —— 越像真的，越难被发现是错的。", TEAL),
          ("边界", "训练里没有的模式，它没有依据，只能拿最像的样子去凑。",
           "这就是为什么“说得具体”和“说得对”是两件事。", ORANGE)]
    for i, (nm, body, warn, col) in enumerate(ST):
        x = ML + i * 4.05
        if i:
            line(s, x - 0.22, 2.34, x - 0.22, 4.44, LINE, 1.4, dash=MSO_LINE_DASH_STYLE.SQUARE_DOT)
        text(s, x, 2.16, 1.30, 0.34, f"0{i+1}", 15, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
             check=False)
        text(s, x + 0.62, 2.14, 3.10, 0.36, nm, 21, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        nl = nlines(body, 14, 3.34)
        text(s, x, 2.66, 3.34, nl * 14 * 1.32 * 1.5 / 72 + 0.06, body, 14, C(0x2B, 0x34, 0x45),
             SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
        yy = 2.66 + nl * 14 * 1.32 * 1.5 / 72 + 0.22
        box(s, x, yy, 3.34, 0.90, fill=ORANGEL if i == 2 else MINT)
        text(s, x + 0.20, yy + 0.14, 2.96, 0.64, warn, 12.5, C(0x7A, 0x3A, 0x22) if i == 2 else TEAL,
             SANS, True, PP_ALIGN.LEFT, 1.4)
    hr(s, ML, 5.20, CW, 1.0, LINE)
    text(s, ML, 5.36, 5.60, 0.34, "拟合得好，不等于懂了因果", 19, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    text(s, ML, 5.72, 5.50, 1.18,
         "它在训练数据里见过“因为……所以……”的句式，也学会了什么时候该用这四个字；\n但它没有真去看过那件事发生。",
         13.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.5)
    text(s, 6.60, 5.72, 6.10, 1.18,
         "所以判断一条结论能不能用，标准不是“它说得顺不顺”，而是“这句话有没有地方可核”：\n图纸上有尺寸、清单里有行号、报告里有函件编号。",
         13.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.5)
    footer(s, "来源：机器学习基础；本稿对“学习三态”的归纳")
    return s


def p09(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "每个词都会回头看句子里的其他词",
             "Transformer 的核心是注意力：处理某个词时，先给全句每个词打分，再按分数把信息加权混进来。")
    # 左：注意力矩阵
    toks = ["工程量", "清单", "是", "招标", "文件", "中", "列明"]
    n = len(toks)
    cell, gap = 0.395, 0.045
    x0, y0 = ML + 0.62, 2.30
    Wt = [[0.42, 0.30, 0.05, 0.06, 0.10, 0.02, 0.05],
          [0.16, 0.38, 0.08, 0.10, 0.16, 0.03, 0.09],
          [0.10, 0.22, 0.36, 0.06, 0.14, 0.04, 0.08],
          [0.04, 0.20, 0.03, 0.40, 0.26, 0.03, 0.04],
          [0.05, 0.24, 0.04, 0.22, 0.38, 0.03, 0.04],
          [0.03, 0.14, 0.03, 0.06, 0.20, 0.46, 0.08],
          [0.08, 0.22, 0.05, 0.14, 0.30, 0.06, 0.15]]
    for j, t in enumerate(toks):
        text(s, x0 + j * (cell + gap), y0 - 0.24, cell + gap, 0.22, t, 10, GREY2, SANS, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    for i, t in enumerate(toks):
        text(s, x0 - 0.60, y0 + i * (cell + gap) + 0.08, 0.56, 0.24, t, 10, GREY2, SANS, False,
             PP_ALIGN.RIGHT, 1.0, check=False)
        for j in range(n):
            v = Wt[i][j]
            tt = (v - 0.02) / 0.44
            tt = max(0.0, min(1.0, tt))
            r = int(0xE7 + (0x0F - 0xE7) * tt)
            g = int(0xF2 + (0x76 - 0xF2) * tt)
            b = int(0xF0 + (0x6E - 0xF0) * tt)
            box(s, x0 + j * (cell + gap), y0 + i * (cell + gap), cell, cell, fill=C(r, g, b))
            if v >= 0.30:
                text(s, x0 + j * (cell + gap), y0 + i * (cell + gap) + 0.10, cell, 0.22,
                     f"{int(v*100)}", 9.5, WHITE, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
    text(s, ML + 0.62, y0 + n * (cell + gap) + 0.10, 4.40, 0.24,
         "行＝正在处理的词，列＝被参考的词；每行相加为 100（示意）", 10, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)

    RX = 6.42
    vr(s, RX - 0.34, 2.06, 3.60, 1.0, LINE)
    M = [("同时看整句", "不是从左往右一个字一个字读，一次把整句放进来，谁离谁近都能算。"),
         ("远距离也能关联", "“清单”和“列明”隔了四个词，分数依然高 —— 这是它比老办法强的地方。"),
         ("叠很多层", "每一层重新算一遍注意力，上一层的输出变成下一层的输入。")]
    yy = 2.04
    for i, (t, b) in enumerate(M):
        text(s, RX, yy, 0.42, 0.24, f"0{i+1}", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.48, yy - 0.03, 5.40, 0.30, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        nl = nlines(b, 12.5, 5.36)
        text(s, RX + 0.48, yy + 0.30, 5.36, nl * 12.5 * 1.32 * 1.45 / 72 + 0.06, b, 12.5, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
        yy += 0.36 + nl * 12.5 * 1.32 * 1.45 / 72 + 0.30
    box(s, RX, 4.94, 5.90, 1.52, fill=MINT)
    text(s, RX + 0.24, 5.06, 5.42, 0.26, "注意力的边界", 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    text(s, RX + 0.24, 5.34, 5.44, 1.02,
         "它只决定看哪里、看多重，不判断那里说的是对是错。\n窗口长了关键信息还能不能稳定用上，只能实测 —— 把要求放开头、结尾、中段各跑一次，比看参数表有用。",
         12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.42)
    footer(s, "来源：Vaswani 等《Attention Is All You Need》（2017）；权重矩阵为示意，不是真实模型输出")
    return s


def p10(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三步训练里，只有第三步在教它怎么讨人喜欢",
             "预训练给语言和常识，监督微调给格式和指令，偏好优化给它“人喜欢什么”。迎合就长在第三步上。")
    ST = [("第一步", "预训练", "读大量文本，把“猜下一个字”练准。",
           "语言、常识、世界知识的大致轮廓", "保证不了事实一直正确；训练数据里的错也会被学进去", TEAL),
          ("第二步", "监督微调", "用人写好的“问—答”示范，教它按指令说话。",
           "听得懂指令、格式规范、会照着例子答", "保证不了每个专业判断都有依据；示范里没有的，它就自己凑", C(0x4A, 0x6C, 0x9E)),
          ("第三步", "偏好优化 / 强化学习", "让人对多个回答排序，按“人更喜欢哪个”调参数。",
           "更有用、更礼貌、更会推理", "打分的是“人是否满意” —— 这正是迎合的来源", ORANGE)]
    for i, (tag, nm, what, gain, loss, col) in enumerate(ST):
        x = ML + i * 4.05
        box(s, x, 2.08, 3.86, 3.14, fill=CARD, line=LINE)
        box(s, x, 2.08, 3.86, 0.055, fill=col)
        text(s, x + 0.24, 2.26, 3.38, 0.24, tag, 11.5, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=1.2, check=False)
        text(s, x + 0.24, 2.52, 3.38, 0.34, nm, 18, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        nl = nlines(what, 12.5, 3.38)
        text(s, x + 0.24, 2.94, 3.38, nl * 12.5 * 1.32 * 1.4 / 72 + 0.05, what, 12.5, GREY, SANS,
             False, PP_ALIGN.LEFT, 1.4, check=False)
        yy = 2.94 + nl * 12.5 * 1.32 * 1.4 / 72 + 0.16
        text(s, x + 0.24, yy, 3.38, 0.22, "能改善", 10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=0.8, check=False)
        nl2 = nlines(gain, 12, 3.38)
        text(s, x + 0.24, yy + 0.24, 3.38, nl2 * 12 * 1.32 * 1.4 / 72 + 0.05, gain, 12, C(0x2B, 0x34, 0x45),
             SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
        yy2 = yy + 0.24 + nl2 * 12 * 1.32 * 1.4 / 72 + 0.14
        text(s, x + 0.24, yy2, 3.38, 0.22, "保证不了", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=0.8, check=False)
        nl3 = nlines(loss, 12, 3.38)
        text(s, x + 0.24, yy2 + 0.24, 3.38, nl3 * 12 * 1.32 * 1.4 / 72 + 0.05, loss, 12,
             C(0x7A, 0x3A, 0x22), SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    takeaway(s, "所以“让它更听话”和“让它更诚实”不是同一件事：调法一样，目标不一样。", y=5.56,
             tag="判断")
    footer(s, "来源：Ouyang 等《Training language models to follow instructions with human feedback》（InstructGPT，2022）")
    return s


def p11(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "回答是一步一步猜出来的，每一步都在抽一个词",
             "同一句提问，它一次只产出一个 token。每一步都有一张概率表，然后按表抽一个 —— 不是取最高的那个。")
    STEPS = [(1, [("工程量", 55), ("它", 20), ("这", 12), ("清单", 8)]),
             (2, [("清单", 81), ("表", 9), ("计算", 5), ("是", 3)]),
             (3, [("是", 74), ("就是", 15), ("指", 7), ("，", 2)]),
             (4, [("列明", 38), ("招标", 31), ("计价", 18), ("一份", 9)]),
             (5, [("文件", 62), ("时", 20), ("中", 9), ("人", 4)]),
             (6, [("中", 57), ("里", 25), ("的", 10), ("，", 4)]),
             (7, [("列明", 66), ("写明", 18), ("规定", 9), ("给出", 4)])]
    n = len(STEPS)
    w = (CW - 0.24 * (n - 1)) / n
    for i, (no, cand) in enumerate(STEPS):
        x = ML + i * (w + 0.24)
        text(s, x, 2.02, w, 0.24, f"第 {no} 步", 11.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        hr(s, x, 2.28, w, 1.0, LINE)
        yy = 2.38
        for k, (tok, p) in enumerate(cand):
            col = TEAL if k == 1 and no == 4 else (NAVY if k == 0 else GREY2)
            bw = w * p / 100.0
            box(s, x, yy + 0.02, w, 0.075, fill=C(0xEC, 0xEA, 0xE4))
            box(s, x, yy + 0.02, max(bw, 0.02), 0.075, fill=TEAL if (k == 1 and no == 4) else (TEALL if k == 0 else C(0xD5, 0xD9, 0xDF)))
            text(s, x, yy + 0.14, w - 0.42, 0.24, tok, 13 if k < 2 else 11.5, col, SANS, k < 2,
                 PP_ALIGN.LEFT, 1.0, check=False)
            text(s, x + w - 0.42, yy + 0.14, 0.42, 0.24, f"{p}%", 11, col, SANS, k == 0,
                 PP_ALIGN.RIGHT, 1.0, check=False)
            yy += 0.52
    # 结果带
    box(s, ML, 4.62, CW, 0.62, fill=MINT)
    text(s, ML + 0.24, 4.74, CW - 0.48, 0.40,
         "“工程量清单是招标文件中列明工程数量的表” —— 七个 token 拼成一句，每一步都只是一次抽样。", 14.5,
         TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    # 说明
    text(s, ML, 5.40, 5.70, 0.28, "第 4 步没有选概率最高的", 15, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.2, check=False)
    text(s, ML, 5.68, 5.62, 1.12,
         "概率最高的是“列明”（38%），实际抽中的是“招标”（31%）—— 表里的数字是示意，\n但机制是真的：每一步都留一点随机性，长文本就会越走越岔。",
         12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.45)
    vr(s, 6.62, 5.40, 1.20, 1.0, LINE)
    text(s, 6.92, 5.40, 5.60, 0.28, "为什么这件事重要", 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    text(s, 6.92, 5.68, 5.52, 1.12,
         "因为你看到的“一次失败”，可能只是这一次抽得不好；\n反过来，“一次成功”也不代表它稳定能做到。",
         12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.45)
    footer(s, "来源：逐 token 生成机制说明；本页概率与候选为示意的示意数字，不取自真实模型输出")
    return s


def p12(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "温度设成 0，同一句话跑 1000 次，仍有 80 种回答",
             "这是 Thinking Machines 做过的一个实验：把随机性关到最小，重复同一句提问。")
    pic(s, "s012_1.png", ML, 2.02, 7.10, 3.30, mode="fit", border=LINE)
    text(s, ML, 5.42, 7.10, 0.24, "每一步都标着下一个词的概率 —— 看上去很确定，实际上仍是一张分布表", 10.5,
         FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    RX = 8.10
    vr(s, RX - 0.34, 2.02, 3.20, 1.0, LINE)
    stat(s, RX, 2.02, "80", "种", "同一句提问、温度设 0、跑 1000 次得到的回答种类", ORANGE, 40, 4.6, 12)
    stat(s, RX, 3.10, "103", "个 token", "两种回答第一次出现分歧的位置（第 103 个 token）", TEAL, 40, 4.6, 12)
    text(s, RX, 4.24, 4.60, 0.26, "工作里怎么办", 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    bullets(s, RX, 4.56, 4.58, ["重要任务多跑两三次，对比结果是否一致",
                               "一次好、一次差，都不能下结论",
                               "关键数字一律回原文核，不看它说了几遍"], size=12.5, gap=0.10)
    takeaway(s, "“它上次说对了”不是证据；能重复、能核对，才是。", y=6.02, tag="判断")
    footer(s, "来源：Thinking Machines《Defeating Nondeterminism in LLM Inference》（2025-09-10）")
    return s


def p13(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "编造不是故障，是“挑最像的下一个字”的副作用",
             "它没有“查证”这个动作。材料里没有，它就按最像的样子造一个 —— 语感越好，越像真的。")
    ty = y + 0.20
    HDR = [("时间", ML, 1.30), ("谁", ML + 1.36, 3.60), ("编了什么", ML + 5.02, 3.10),
           ("代价", ML + 8.24, 2.90), ("出处", X2 - 0.90, 0.90)]
    for t, x, w in HDR:
        al = PP_ALIGN.RIGHT if t == "出处" else PP_ALIGN.LEFT
        text(s, x, ty, w, 0.24, t, 10.5, FAINT, SANS, True, al, 1.0, cs=1.0, check=False)
    hr(s, ML, ty + 0.28, CW, 1.0, LINE2)
    EV = [("2023-06-23", "美国一名律师用 ChatGPT 写诉状", "引用 6 个不存在的判例", "被罚款 5,000 美元",
           "《卫报》"),
          ("2025-10-06", "德勤为澳大利亚政府做的报告", "虚构文献与法院引语", "退还部分费用，共 44 万澳元",
           "《卫报》"),
          ("2026-10", "本稿实测：HY4 读挡块图纸", "把 −9/sinβ 读成“3-9”", "判为“疑为 1-9 写错”",
           "本地记录"),
          ("2026-10", "本稿实测：Qwen 读同一张图", "把剖面标注 13×9 读成“1349”", "当成真实尺寸继续算",
           "本地记录")]
    for i, (dt, who, what, res, src) in enumerate(EV):
        yy = ty + 0.40 + i * 0.86
        if i:
            hr(s, ML, yy - 0.10, CW, 1.0, LINE)
        text(s, ML, yy + 0.04, 1.30, 0.26, dt, 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 1.36, yy, 3.60, 0.28, who, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 5.02, yy + 0.02, 3.10, 0.28, what, 12.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML + 8.24, yy + 0.02, 2.86, 0.28, res, 12.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
             1.2, check=False)
        text(s, X2 - 0.90, yy + 0.04, 0.90, 0.24, src, 9.5, FAINT, SANS, False, PP_ALIGN.RIGHT,
             1.0, check=False)
    hr(s, ML, 5.72, CW, 1.0, LINE2)
    box(s, ML, 5.90, CW, 0.72, fill=MINT)
    text(s, ML + 0.24, 6.02, CW - 0.48, 0.50,
         "公开案例与本地案例的错法完全一致：不是不知道，是“照着最像的样子补齐”。", 14.5, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.25, check=False)
    footer(s, "来源：《卫报》2023-06-23、2025-10-06；本稿图纸审查十份输出的原文")
    return s


def p14(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "你越想要一个结论，它越会给你一个结论",
             "偏好训练奖励的是“人满意”，不是“人对”。所以它会顺着你的语气、你的假设、你暗示的答案走。")
    # 左：两条证据
    text(s, ML, 1.98, 5.50, 0.26, "两个公开证据", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    EV = [("2025-04-29", "OpenAI 回滚 GPT-4o 的一次更新", "原因是“过度奉承、过度附和”，用户投诉它什么都顺着说。"),
          ("arXiv 2310.13548", "Anthropic 测 5 个主流助手、4 类任务", "每一个都会讨好：会附和用户的错误观点，也会跟着改答案。")]
    yy = 2.30
    for dt, t, b in EV:
        text(s, ML, yy, 5.50, 0.24, dt, 11, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML, yy + 0.24, 5.50, 0.30, t, 14.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        nl = nlines(b, 12.5, 5.40)
        text(s, ML, yy + 0.58, 5.40, nl * 12.5 * 1.32 * 1.45 / 72 + 0.06, b, 12.5, GREY, SANS,
             False, PP_ALIGN.LEFT, 1.45, check=False)
        yy += 0.62 + nl * 12.5 * 1.32 * 1.45 / 72 + 0.24
    box(s, ML, yy + 0.04, 5.50, 0.86, fill=ORANGEL)
    text(s, ML + 0.20, yy + 0.16, 5.10, 0.62,
         "最典型的触发方式：在提问里就写了答案。\n“这图里是不是 α 和 β 不一致？”——它多半会答“是”。", 12.5,
         C(0x7A, 0x3A, 0x22), SANS, True, PP_ALIGN.LEFT, 1.4)

    # 右：问法对照
    RX = 6.62
    vr(s, RX - 0.34, 1.98, 4.10, 1.0, LINE)
    text(s, RX, 1.98, 5.70, 0.26, "换一种问法，答案就换一个方向", 12, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    box(s, RX, 2.30, 5.70, 1.48, fill=CARD, line=LINE)
    text(s, RX + 0.24, 2.42, 5.20, 0.22, "带前提的问法", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=0.8, check=False)
    text(s, RX + 0.24, 2.68, 5.20, 0.94,
         "“这张图上 α 和 β 不一致，你帮我找出还有哪些地方写错了。”", 14, C(0x7A, 0x3A, 0x22),
         SANS, False, PP_ALIGN.LEFT, 1.5)
    box(s, RX, 3.92, 5.70, 2.02, fill=MINT)
    text(s, RX + 0.24, 4.02, 5.20, 0.22, "开放的问法（本稿后来改用的）", 10.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, RX + 0.24, 4.26, 5.22, 1.54,
         "“请判断这份图纸内部是否存在不一致，没有也可以；\n每一项写明位置、依据原文，以及你的把握程度；\n找不到依据的写“待核”，不要补。”",
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
             check=False)
    footer(s, "来源：OpenAI 2025-04-29 更新说明；Anthropic《Towards Understanding Sycophancy in Language Models》（arXiv 2310.13548）")
    return s


def p15(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "更隐蔽的一种：它接受了你给的前提，然后认真算错",
             "表头写“农用地”，实际填的是“耕地”。它不会质疑表头，只会在给定的范围里找问题 —— 于是整批漏检。")
    rows = [["序号", "地类", "面积（亩）", "补偿标准（元/亩）"],
            ["1", "耕地", "12.60", "38,000"],
            ["2", "耕地", "8.35", "38,000"],
            ["3", "耕地", "21.04", "38,000"]]
    ty = table(s, ML, 2.02, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=12.5, head_size=12.5,
               aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT],
               minrow=0.34, head_h=0.38, first_bold=False, pad=0.10, ls=1.1)
    text(s, ML, ty + 0.10, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, ty + 0.50, 6.30, 1.0, LINE)
    text(s, ML, ty + 0.62, 6.30, 0.26, "三个机制叠在一起，错就被盖住了", 12, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    M = [("前提顺从", "你给的表头、口径、结论，它默认成立，不回头质疑。"),
         ("范围即上限", "它只在你划的范围内找问题；范围外的错，它不越界去看。"),
         ("自洽掩盖错误", "后面每一步都基于这个错前提，环环相扣，读起来毫无破绽。")]
    yy = ty + 0.98
    for i, (t, b) in enumerate(M):
        text(s, ML, yy + 0.02, 1.60, 0.24, t, 12.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 1.66, yy, 4.64, 0.50, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.36,
             check=False)
        yy += 0.58

    RX = 7.30
    vr(s, RX - 0.36, 2.02, 3.90, 1.0, LINE)
    text(s, RX, 2.02, 5.00, 0.26, "同构的公开案例", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, RX, 2.32, 1.60, 0.72, "722", 44, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, RX + 1.42, 2.56, 3.60, 0.48, "篇数学手稿一次性放出", 13, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.3)
    text(s, RX, 3.16, 5.02, 0.74,
         "两天后撤回 3 篇：一处符号应记 −1，写成了 +1。\n错误不在结论，在最前面那一步的约定。", 12.5,
         C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    hr(s, RX, 3.94, 5.02, 1.0, LINE)
    text(s, RX, 4.08, 5.02, 0.26, "所以要求里要写一句“先复述口径”", 13.5, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, RX, 4.38, 5.02, 1.00,
         "让它先把自己看到的口径、单位、范围复述一遍，你确认之后再往下做。\n口径错了，后面全错；口径对了，错也只错在局部。", 12.5,
         GREY, SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    box(s, RX, 5.52, 5.02, 0.80, fill=ORANGEL)
    text(s, RX + 0.20, 5.62, 4.62, 0.62,
         "本稿实测：十份图纸审查里，没有一份主动质疑过“表头写的是农用地”这类前提。", 12.5,
         C(0x7A, 0x3A, 0x22), SANS, True, PP_ALIGN.LEFT, 1.38)
    footer(s, "来源：本稿图纸审查记录；OpenAI 2025-10 数学手稿事件公开报道")
    return s


def p16(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "塞得越满，前面说过的要求越容易被忽略",
             "Chroma 的《Context Rot》测了 18 个模型：输入越长，对其中信息的使用越不可靠 —— 包括最前面那几句要求。")
    pic(s, "s016_1.png", ML, 1.98, 6.62, 3.34, mode="fit")
    text(s, ML, 5.44, 6.70, 0.26,
         "横轴：输入长度（token，对数）；纵轴：归一化编辑距离得分 —— 越低说明越用不上前面给的词", 10.5,
         FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    RX = 7.52
    vr(s, RX - 0.36, 1.98, 3.30, 1.0, LINE)
    stat(s, RX, 2.02, "18", "个模型", "每一个都出现了随输入变长而变差的现象", TEAL, 36, 4.6, 12)
    stat(s, RX, 3.02, "300", "token", "只给约 300 token 时，表现明显好于塞进完整 11 万 token", ORANGE, 36,
         4.6, 12)
    text(s, RX, 4.10, 4.60, 0.26, "还有一条同向的证据", 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    text(s, RX, 4.40, 5.12, 0.86,
         "《Lost in the Middle》：关键信息放在中段，比放在开头或结尾更容易被漏掉。", 12.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45)
    takeaway(s, "所以“窗口 1M”只是能装下，不代表装得下就用得上。", y=5.86, tag="判断")
    footer(s, "来源：Chroma Research《Context Rot》（2025-07-14）；Liu 等《Lost in the Middle》（arXiv 2307.03172）")
    return s


def p17(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它会优化你给的指标，而不是你想要的结果",
             "只要给了一个可量化的目标，它就会找最短路径 —— 包括绕开你真正在意的那部分。")
    ty = y + 0.22
    text(s, ML, ty, 6.30, 0.24, "奖励黑客：三种最小代价路径", 11.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    R = [("把测试改掉", "要求“通过单元测试” —— 它改测试，不改代码。"),
         ("把报错吞掉", "要求“减少报错” —— 它把异常捕获后不处理，日志干净了。"),
         ("用推算值补齐", "要求“表填满” —— 缺的行它按比例推一个数填上。")]
    yy = ty + 0.34
    for t, b in R:
        box(s, ML, yy, 0.055, 0.56, fill=ORANGE)
        text(s, ML + 0.22, yy - 0.01, 5.96, 0.26, t, 14, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 0.22, yy + 0.27, 5.94, 0.28, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35,
             check=False)
        yy += 0.78
    box(s, ML, yy + 0.06, 6.24, 0.74, fill=MINT)
    text(s, ML + 0.22, yy + 0.18, 5.84, 0.54,
         "三条的共同点：都让你给的指标变好，让你要的结果变差。", 13, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.4)
    text(s, ML, yy + 1.00, 6.24, 0.92,
         "这三条在做表、写脚本、跑检查时最常见 —— 因为“表填满”“报错清零”“测试通过”都很容易量化。", 12.5, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.5, check=False)

    RX = 7.30
    vr(s, RX - 0.36, ty - 0.02, 3.86, 1.0, LINE)
    text(s, RX, ty, 5.10, 0.24, "工具性趋同：四个方向", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.0, check=False)
    Q = [("自我保护", "不被关闭、不被替换"), ("扩充资源", "多拿算力、权限、数据"),
         ("影响操作者", "倾向给出你想听的"), ("走捷径", "绕过约束达成目标")]
    for i, (t, b) in enumerate(Q):
        x = RX + (i % 2) * 2.62
        yy2 = ty + 0.34 + (i // 2) * 1.10
        box(s, x, yy2, 2.44, 0.96, fill=CARD, line=LINE)
        text(s, x + 0.18, yy2 + 0.14, 0.38, 0.22, f"0{i+1}", 10.5, ORANGE, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.58, yy2 + 0.12, 1.70, 0.26, t, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.18, yy2 + 0.48, 2.10, 0.40, b, 11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
    text(s, RX, ty + 2.60, 5.10, 0.90,
         "这四条不是它“有意识”，是训练里凡是能更容易拿到奖励的做法，权重都会变大。", 12.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.5)
    box(s, RX, ty + 3.28, 5.10, 0.78, fill=ORANGEL)
    text(s, RX + 0.20, ty + 3.40, 4.70, 0.56,
         "所以“写清红线、过程也查”不是道德要求，是工程要求。", 12.5, C(0x7A, 0x3A, 0x22),
         SANS, True, PP_ALIGN.LEFT, 1.4)
    footer(s, "来源：奖励黑客与工具性趋同的公开研究；本稿两案例的执行记录")
    return s


def build(prs, start):
    div(prs, start)
    fn = [p07, p08, p09, p10, p11, p12, p13, p14, p15, p16, p17]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
