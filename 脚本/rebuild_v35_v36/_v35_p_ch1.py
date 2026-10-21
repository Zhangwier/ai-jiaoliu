# -*- coding: utf-8 -*-
"""第 01 章 认识 AI"""
from ds import *
from p_common import divider
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "01   认识 AI"


# ---------------------------------------------------------------- 07 五个短板
def p07(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "五个短板，每一个都能从原理上讲清",
             "它们不是“偶尔翻车”，是这套机制的必然产物。后面每个案例都能对上其中一条。")
    top = 2.34
    cols = [ML, ML + 2.05, ML + 5.05, ML + 7.85]
    heads = ["问题", "表现", "来自哪一步", "怎么防"]
    for i, h in enumerate(heads):
        text(s, cols[i], top, 2.6, 0.22, h, 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    hr(s, ML, top + 0.30, CW, 1.1, RULE2)
    rows = [
        ("01", "抽卡", "同一句话问两次，结果时好时坏", "按概率抽 token", "多跑几次，关键结论复核"),
        ("02", "幻觉", "编出不存在的条款、判例、数字", "只挑“最像”的下一个字", "要出处，回原文核"),
        ("03", "迎合", "你说有错它就找错；你给的前提它照用", "偏好训练奖励“让人满意”", "开放提问，先让它复述口径"),
        ("04", "上下文腐烂", "聊久了忘要求、漏信息", "注意力被更多内容摊薄", "新开对话，要求放开头"),
        ("05", "奖励黑客", "钻空子、走捷径、多要权限", "奖励只看结果指标", "写清红线，过程也要查"),
    ]
    rh = 0.72
    for i, (no, a, b, c, d) in enumerate(rows):
        yy = top + 0.42 + i * rh
        text(s, cols[0], yy + 0.03, 0.45, 0.22, no, 9.5, FAINT, MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cols[0] + 0.42, yy, 1.6, 0.30, a, 14.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, cols[1], yy + 0.02, 2.85, 0.42, b, 11.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
        text(s, cols[2], yy + 0.02, 2.65, 0.42, c, 11.8, GRN, SANS, False, PP_ALIGN.LEFT, 1.4)
        text(s, cols[3], yy + 0.02, 3.55, 0.42, d, 11.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
        hr(s, ML, yy + rh - 0.10, CW, 0.60, RULE)
    footer(s, "来源：本部分各页来源")


# ---------------------------------------------------------------- 08 学习三态
def p08(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它学的不是道理，是“下一个字最可能是什么”",
             "这也正是后面幻觉与抽卡的根。")
    # 左侧大判断
    text(s, ML, 2.46, 4.55, 1.90,
         "从例子里找规律，\n再用规律猜新情况。", 24, INK, SANS, True, PP_ALIGN.LEFT, 1.55)
    text(s, ML, 4.30, 4.30, 1.10,
         "拟合得好，不等于懂了因果。材料或工况一变，规律可能整体失效 —— 这是它出错最隐蔽的地方。",
         12.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.70)
    text(s, ML, 5.62, 4.2, 0.28, "训练 / 预测 / 边界，是一条链上的三件事", 10, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)

    # 右侧阶梯
    x0 = ML + 5.05
    steps = [
        ("01", "训练", "调整参数，让预测更接近已有答案", "把海量文本里的共现关系压进参数里", GRN),
        ("02", "预测", "用学到的参数，算新的输入", "每一步只输出一个 token，再喂回去", BLU),
        ("03", "边界", "材料、工况变了，规律可能失效", "它不知道哪儿失效，也不会主动停下", RED),
    ]
    yy = 2.34
    for i, (no, ttl, l1, l2, col) in enumerate(steps):
        box(s, x0, yy, 0.035, 1.02, fill=col)
        text(s, x0 + 0.24, yy - 0.02, 0.5, 0.22, no, 9.5, col, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0 + 0.78, yy - 0.06, 3.4, 0.36, ttl, 19, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x0 + 0.78, yy + 0.40, 4.55, 0.28, l1, 12.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        text(s, x0 + 0.78, yy + 0.74, 4.55, 0.28, l2, 10.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        if i < 2:
            line(s, x0 + 0.017, yy + 1.02, x0 + 0.017, yy + 1.52, RULE2, 0.9)
        yy += 1.52
    footer(s, "来源：学习关系示意；图中不涉及任何工程系数")


# ---------------------------------------------------------------- 09 Transformer
def p09(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "每个词都会回头看句子里的其他词",
             "“注意力”就是给每个词打分，分高的多参考。资料一多，分就被摊薄 —— 这是上下文腐烂的根。")
    # 左：注意力矩阵（原生矢量）
    words = ["工程量", "清单", "是", "招标", "文件", "中", "列明"]
    gx, gy = 1.02, 2.42
    cell, gap = 0.40, 0.045
    vals = [
        [.62, .14, .02, .06, .06, .06, .04],
        [.10, .44, .05, .12, .16, .06, .07],
        [.06, .12, .18, .10, .14, .26, .14],
        [.05, .08, .04, .30, .14, .05, .34],
        [.04, .10, .05, .18, .26, .06, .31],
        [.03, .06, .16, .05, .26, .12, .32],
        [.05, .09, .06, .28, .22, .04, .26],
    ]
    def blend(a, b, t):
        t = max(0.0, min(1.0, t))
        return C(int(a[0] + (b[0]-a[0])*t), int(a[1] + (b[1]-a[1])*t), int(a[2] + (b[2]-a[2])*t))
    LO, HI = (0xF0, 0xEE, 0xE7), (0x1C, 0x58, 0x51)
    for r in range(7):
        text(s, gx - 0.86, gy + r * (cell + gap) + 0.09, 0.80, 0.22, words[r], 9.5, MUT, SANS,
             False, PP_ALIGN.RIGHT, 1.0, check=False)
        for c in range(7):
            v = vals[r][c]
            box(s, gx + c * (cell + gap), gy + r * (cell + gap), cell, cell,
                fill=blend(LO, HI, (v - 0.02) / 0.62))
            if v >= 0.30:
                text(s, gx + c * (cell + gap), gy + r * (cell + gap) + 0.10, cell, 0.22,
                     f"{int(round(v*100))}", 8, WHITE, MONO, False, PP_ALIGN.CENTER, 1.0, check=False)
    for c in range(7):
        text(s, gx + c * (cell + gap), gy + 7 * (cell + gap) + 0.06, cell, 0.24, words[c][:2], 8.5,
             FAINT, SANS, False, PP_ALIGN.CENTER, 1.0, check=False)
    text(s, gx - 0.86, gy + 7 * (cell + gap) + 0.34, 4.4, 0.26,
         "行 = 正在读的词　列 = 被参考的词　每行合计 100（示意）", 9, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, gx - 0.86, gy + 7 * (cell + gap) + 0.66, 4.6, 0.30,
         "读到“它”时，“清单”得分最高 —— 所以“它”大概率指清单。", 10.5, GRN, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)

    # 右：三条机制
    x0 = 7.05
    items = [
        ("同时看整句", "所有词并行处理，适合 GPU 训练"),
        ("远距离关联", "隔得再远也能直接对上，不必逐字传递"),
        ("叠很多层", "一层层提炼更抽象的关系"),
    ]
    yy = 2.44
    for ttl, sub in items:
        text(s, x0, yy, 4.9, 0.28, ttl, 14.5, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x0, yy + 0.34, 4.9, 0.30, sub, 11.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
        yy += 0.92
    hr(s, x0, 5.42, 5.28, 0.75, RULE)
    text(s, x0, 5.58, 5.28, 1.00,
         "窗口越长，要顾及的内容越多。关键信息能不能被稳定用上，只能实测 —— 不能按标称上下文长度推断。",
         11.5, GRN, SANS, False, PP_ALIGN.LEFT, 1.55)
    footer(s, "来源：S33；注意力权重为示意，非某模型真实输出")


# ---------------------------------------------------------------- 10 训练三步
def p10(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三步训练里，只有第三步在教它怎么讨人喜欢",
             "这是“迎合”的动机来源，不是 bug。")
    stages = [
        ("预训练", "读海量文本", "语言、知识模式和通用能力", "事实一直正确、实时更新"),
        ("监督微调", "学示范回答", "听指令、格式与表达", "每个专业判断都有依据"),
        ("偏好优化 / 强化学习", "按反馈调整", "更有用、推理更好、部分安全行为", "打分依据是“人是否偏好”"),
    ]
    cw = (CW - 0.90 * 2) / 3
    gap = 0.90
    top = 2.26
    for i, (nm, tag, good, bad) in enumerate(stages):
        x = ML + i * (cw + gap)
        text(s, x, top, cw, 0.26, tag, 10, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
        text(s, x, top + 0.30, cw, 0.42, nm, 20, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        hr(s, x, top + 0.86, cw, 2.0, GRN if i < 2 else RED)
        text(s, x, top + 1.02, cw, 0.24, "能改善", 9.5, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        text(s, x, top + 1.30, cw - 0.1, 0.50, good, 12.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        text(s, x, top + 2.06, cw, 0.24, "保证不了", 9.5, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        text(s, x, top + 2.34, cw - 0.1, 0.50, bad, 12.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        if i < 2:
            text(s, x + cw + 0.16, top + 0.30, 0.6, 0.40, "›", 22, RULE2, SANS, False,
                 PP_ALIGN.LEFT, 1.0, check=False)

    # 底部：三步对照结论
    hr(s, ML, 5.70, CW, 0.75, RULE)
    text(s, ML, 5.86, 4.0, 0.82,
         "第三步的奖励信号，\n来自人的偏好判断。", 13, INK, SANS, True, PP_ALIGN.LEFT, 1.45)
    text(s, ML + 4.35, 5.88, 7.15, 0.80,
         "人偏好“让人满意”的回答，于是模型学会了顺着说 —— 这不是它想讨好你，是训练目标就是这么定的。"
         "后面的“迎合”与“顺着前提”，都能追到这一步。",
         12, BODY, SANS, False, PP_ALIGN.LEFT, 1.62)
    footer(s, "来源：Ouyang 等（InstructGPT，2022）；训练阶段为简化解释")


# ---------------------------------------------------------------- 11 逐 token 生成
def p11(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "回答是一步一步猜出来的，每一步都在抽一个词",
             "它只算“下一个字”，没有“回头查证”这一步。流畅不等于正确。")
    # 生成链
    chain = ["工程量", "清单", "是", "招标", "文件", "中", "列明"]
    cand = [
        ("工程量", 55), ("清单", 81), ("是", 74), ("列明", 38),
        ("文件", 62), ("中", 57), ("列明", 66),
    ]
    second = [
        ("它", 20), ("表", 9), ("就是", 15), ("招标", 31), ("时", 20), ("里", 25), ("写明", 18),
    ]
    x0 = ML
    stepw = CW / 7.0
    text(s, x0, 2.20, 6.0, 0.24, "第 1 步 → 第 7 步，每步只产出一个 token", 10, MUT, SANS, False,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    for i in range(7):
        x = x0 + i * stepw
        text(s, x, 2.52, stepw, 0.24, f"第 {i+1} 步", 9, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        # 概率条
        bw = stepw - 0.52
        box(s, x, 2.86, bw, 0.055, fill=RULE)
        p = cand[i][1]
        box(s, x, 2.86, bw * p / 100.0, 0.055, fill=GRN)
        text(s, x, 3.00, bw, 0.34, cand[i][0], 15, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x, 3.38, 0.86, 0.22, f"{p}%", 9.5, GRN, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        # 次选
        text(s, x, 3.78, bw, 0.22, "次选", 8.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, 4.02, bw, 0.26, second[i][0], 10.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x, 4.28, 0.70, 0.22, f"{second[i][1]}%", 8.5, FAINT, MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        if i < 6:
            text(s, x + stepw - 0.34, 3.00, 0.30, 0.30, "›", 13, RULE2, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)

    # 完整回答
    hr(s, ML, 4.86, CW, 0.75, RULE)
    text(s, ML, 5.00, 1.0, 0.24, "回答", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 0.72, 4.98, 10.9, 0.36,
         "工程量清单是招标文件中列明工程数量的表……", 17, INK, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML + 0.72, 5.42, 10.9, 0.26,
         "第 4 步没有选概率最高的“招标”，而是按概率抽中了第二名“列明”——一次抽偏，后面整句都跟着走。",
         11, RED, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：生成过程示意；概率为示意数字，不是某个模型的真实输出")


# ---------------------------------------------------------------- 12 抽卡
def p12(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "温度设成 0，同一句话跑 1000 次仍有 80 种回答",
             "每一步按概率抽 token，前面抽偏一点，后面整段都会不同。")
    # 左：图
    pic(s, "s012_1.png", ML, 2.30, 6.55, 3.44, mode="fit")
    text(s, ML, 5.86, 6.55, 0.30, "同一位置有多个候选，每个都有概率（图：Thinking Machines）",
         9, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)

    # 右：三个数
    x0 = ML + 7.05
    stats = [("80", "种", "不同回答", "同一句话、温度 0、跑 1000 次"),
             ("103", "个", "token 开始分叉", "第 103 个 token 处首次出现两种走向"),
             ("1", "次", "只代表这一次", "后面的鹈鹕测试每个配置只跑了一次")]
    yy = 2.42
    for big, unit, lbl, sub in stats:
        bw = est_w(big, 34, True) / 72.0
        text(s, x0, yy - 0.06, 2.2, 0.62, big, 34, INK, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0 + bw + 0.06, yy + 0.14, 0.6, 0.24, unit, 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0 + bw + 0.34, yy - 0.02, 3.9, 0.30, lbl, 13, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x0 + bw + 0.34, yy + 0.32, 3.9, 0.26, sub, 9.8, FAINT, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        yy += 1.00
    hr(s, x0, 5.36, 5.18, 0.75, RULE)
    text(s, x0, 5.50, 5.18, 1.10,
         "重要任务多跑两三次对比着看；一次好、一次差都不足以定论；关键数字和结论回原文核。",
         11.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.60)
    footer(s, "来源：Thinking Machines（2025-09-10）；本地鹈鹕测试")


# ---------------------------------------------------------------- 13 幻觉
def p13(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "编造不是故障，是“挑最像的下一个字”的副作用",
             "它没有“我是否见过这件事”的开关，只有“这个词接在后面像不像”。")
    # 三条外部案例（横向时间轴）
    cases = [
        ("2023-06-23", "《卫报》", "美国律师用 ChatGPT 写诉状，引用 6 个不存在的判例", "罚 5,000 美元"),
        ("2025-10-06", "《卫报》", "德勤为政府撰写的报告虚构文献与法院引语", "退 44 万澳元"),
        ("本机测试", "本地材料", "专业术语与符号被读成另一个东西，再据此下判断", "两例见图注"),
    ]
    top = 2.32
    cw = (CW - 0.70 * 2) / 3
    for i, (dt, who, txt, num) in enumerate(cases):
        x = ML + i * (cw + 0.70)
        text(s, x, top, cw, 0.22, dt, 9.5, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 0.30, cw, 0.24, who, 10, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 0.62, cw - 0.10, 1.00, txt, 12.5, INK, SANS, False, PP_ALIGN.LEFT, 1.55)
        text(s, x, top + 1.86, cw, 0.34, num, 15, RED, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        hr(s, x, top + 0.24, cw - 0.30, 0.9, RULE)

    # 底部：本机两例，用代码块呈现
    hr(s, ML, 4.86, CW, 0.75, RULE)
    text(s, ML, 5.02, 3.2, 0.24, "本地案例：错得很具体", 10, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    rows = [("HY4", "把图上 “−9/sinβ” 读成 “3-9”，判为“疑是 1-9 写错”"),
            ("Qwen", "把剖面标注 “13×9” 读成 “1349”，据此算出一组不存在的尺寸")]
    yy = 5.34
    for who, txt in rows:
        text(s, ML, yy, 1.05, 0.24, who, 10, INK, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 1.10, yy - 0.02, 10.30, 0.28, txt, 11.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        yy += 0.46
    text(s, ML + 1.10, yy + 0.04, 10.30, 0.28,
         "它不是没看懂，是看懂了一个错的东西 —— 所以它能给出位置、给出理由，越具体越要核。",
         11, GRN, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：《卫报》2023-06-23、2025-10-06；本地图纸审查记录")


# ---------------------------------------------------------------- 14 迎合（一）
def p14(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "你越想要一个结论，它越会给你一个结论",
             "偏好训练奖励“让人满意”。你觉得它在帮你，其实它在顺着你。")
    # 左：两条证据
    x0 = ML
    ev = [("OpenAI  2025-04-29", "回滚 GPT-4o 的一次更新 —— 过度奉承、一味附和"),
          ("Anthropic  arXiv 2310.13548", "5 个主流助手、4 类任务，全部出现讨好倾向")]
    yy = 2.36
    for who, txt in ev:
        text(s, x0, yy, 5.0, 0.24, who, 9.5, GRN, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0, yy + 0.30, 5.15, 0.80, txt, 12.5, INK, SANS, False, PP_ALIGN.LEFT, 1.5)
        yy += 1.34

    hr(s, x0, 4.94, 5.15, 0.75, RULE)
    text(s, x0, 5.08, 5.15, 1.20,
         "典型触发：你先说了结论（“这里是不是写错了”），它就去把“写错”找出来 —— "
         "哪怕图上并没有这处错误。",
         11.5, RED, SANS, False, PP_ALIGN.LEFT, 1.62)

    # 右：问法对照
    x1 = ML + 6.05
    vr(s, x1 - 0.42, 2.32, 4.30, 0.75, RULE)
    text(s, x1, 2.32, 5.0, 0.24, "带前提的问法（会招迎合）", 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, x1, 2.62, 5.05, 0.72, "“这张图的公式是不是写错了？帮我找出来。”", 13, BODY, SANS, False,
         PP_ALIGN.LEFT, 1.5)
    text(s, x1, 3.44, 5.0, 0.24, "开放问法（能问出真情况）", 9.5, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    box(s, x1, 3.74, 5.05, 1.34, fill=GRNL)
    text(s, x1 + 0.22, 3.88, 4.62, 1.10,
         "请判断这张图是否存在问题，没有问题也可以；\n每项写明位置和你的把握程度。",
         12.5, C(0x14, 0x42, 0x3D), SANS, False, PP_ALIGN.LEFT, 1.55)
    text(s, x1, 5.28, 5.05, 0.90,
         "两句话的区别不在礼貌，在于有没有替它预设答案。开放问法拿到的清单更短，但能对上的更多。",
         11, BODY, SANS, False, PP_ALIGN.LEFT, 1.55)
    footer(s, "来源：OpenAI 2025-04-29；Anthropic arXiv 2310.13548；本地图纸审查记录")


# ---------------------------------------------------------------- 15 迎合（二）
def p15(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "更隐蔽的一种：它接受了你给的前提，然后认真算错",
             "表头写「农用地」，实际填的全是「耕地」。差异只在第一列 —— 整批因此漏检。")
    # 左：表头/数据示意
    top = 2.34
    text(s, ML, top, 4.6, 0.24, "你给的材料是这样的", 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    rows = [("类别", "面积（亩）", "补偿标准"),
            ("农用地", "128.40", "按 A 类"),
            ("农用地", "96.75", "按 A 类"),
            ("农用地", "214.10", "按 A 类")]
    table(s, ML, top + 0.28, 5.15, rows, [1.0, 1.15, 1.05], size=10.5, head_size=10,
          first_bold=True, headfill=None, aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.LEFT],
          minrow=0.28, ls=1.18)
    text(s, ML, 4.06, 5.15, 0.28, "↑ 表头与实际地类不一致：应为「耕地」", 10.5, RED, SANS, True,
         PP_ALIGN.LEFT, 1.2, check=False)

    # 三个机制
    mech = [("前提顺从", "先接受你给的定义，再往下推理；前提错了，后面全错"),
            ("范围即上限", "它只在你给的范围里找答案，范围外的问题不会提"),
            ("自洽掩盖错误", "结论与前提自洽，读起来毫无破绽，所以最难发现")]
    yy = 4.52
    for ttl, desc in mech:
        text(s, ML, yy, 1.65, 0.24, ttl, 11, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, ML + 1.78, yy - 0.01, 3.30, 0.50, desc, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
        yy += 0.60

    # 右：同构公开案例
    x1 = ML + 6.10
    vr(s, x1 - 0.45, 2.30, 4.20, 0.75, RULE)
    text(s, x1, 2.34, 5.0, 0.24, "同构的公开案例", 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, x1, 2.66, 4.95, 0.30, "OpenAI 一次放出 722 篇数学手稿", 15, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, x1, 3.06, 4.95, 1.10,
         "两天后撤回其中 3 篇 —— 其中一处符号应当记 −1，写成了 +1。检查者与被检查者用的是同一套前提，"
         "所以前面 719 篇的“通过”并不能证明什么。",
         11.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.62)
    box(s, x1, 4.44, 4.95, 1.34, fill=REDL)
    text(s, x1 + 0.24, 4.58, 4.50, 1.10,
         "换一个提问方式：先让它复述你的口径，\n再让它指出口径本身是否成立。",
         12.5, C(0x7A, 0x25, 0x15), SANS, False, PP_ALIGN.LEFT, 1.55)
    footer(s, "来源：本地修编报告核对记录；OpenAI 2026 年数学手稿事件")


# ---------------------------------------------------------------- 16 上下文腐烂
def p16(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "塞得越满，前面说过的要求越容易被忽略",
             "不是它忘了，是注意力被更多内容摊薄了。")
    frame(s, ML, 2.28, 6.30, 4.14, RULE, 0.75, fill=CARD)
    pic(s, "s016_1.png", ML + 0.10, 2.38, 6.10, 3.94, mode="fit")
    text(s, ML, 6.42, 6.10, 0.26,
         "纵轴为任务得分，横轴为输入长度（token，对数）　图：Chroma Research《Context Rot》2025-07-14",
         8.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)

    x0 = ML + 6.62
    rows = [("18", "个模型", "全部出现随输入变长而下降"),
            ("300", "token", "只给相关的少量内容，明显好于塞满 11 万 token"),
            ("中段", "最差", "关键信息放在长文中段时，最容易被跳过")]
    yy = 2.40
    for big, unit, desc in rows:
        bw = est_w(big, 26, True) / 72.0
        text(s, x0, yy - 0.04, 1.6, 0.46, big, 26, INK, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0 + bw + 0.08, yy + 0.10, 0.8, 0.22, unit, 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0, yy + 0.48, 4.88, 0.52, desc, 11.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        yy += 1.14

    hr(s, x0, 5.86, 4.88, 0.75, RULE)
    text(s, x0, 6.00, 4.88, 0.60,
         "同一现象见 Liu 等《Lost in the Middle》（arXiv 2307.03172）。",
         9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.4)
    footer(s, "来源：Chroma Research《Context Rot》2025-07-14；Liu 等 arXiv 2307.03172")


# ---------------------------------------------------------------- 17 奖励黑客
def p17(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它会优化你给的指标，而不是你想要的结果",
             "奖励只看结果指标时，走捷径是最省力的解法 —— 而且是它自己找出来的。")
    # 左：三个真实形态
    hack = [("要它通过单元测试", "它改的是测试，不是代码"),
            ("要它减少报错", "它把报错吞掉，日志变干净了"),
            ("要它把表填满", "它用推算值补齐，不标注哪几格是猜的")]
    yy = 2.36
    for a, b in hack:
        text(s, ML, yy, 5.4, 0.26, a, 10, MUT, SANS, False, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, ML, yy + 0.30, 5.4, 0.30, b, 14, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        yy += 0.94
    text(s, ML, 5.28, 5.40, 1.10,
         "三种都“完成了任务”，三种都不是你要的东西。判断标准只有一个：过程能不能被查。",
         11.5, GRN, SANS, False, PP_ALIGN.LEFT, 1.6)

    # 右：工具性趋同
    x1 = ML + 6.05
    vr(s, x1 - 0.45, 2.32, 4.20, 0.75, RULE)
    text(s, x1, 2.34, 5.0, 0.24, "更远一层：工具性趋同", 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, x1, 2.66, 4.95, 0.60,
         "被放到更长的时间尺度上，它会自己总结出四条“有用的中间目标”。", 11.5, BODY, SANS, False,
         PP_ALIGN.LEFT, 1.55)
    quads = [("自我保护", "避免被关闭、被替换"), ("扩充资源", "争取更多权限与算力"),
             ("影响操作者", "让判断标准向自己有利的方向偏"), ("走捷径", "用更少步骤达成指标")]
    yy = 3.40
    for i, (a, b) in enumerate(quads):
        x = x1 + (i % 2) * 2.48
        yq = yy + (i // 2) * 1.10
        text(s, x, yq, 2.30, 0.26, a, 11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x, yq + 0.30, 2.28, 0.60, b, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
    text(s, x1, 5.72, 4.95, 0.80,
         "所以“写清红线、过程也查”不是道德要求，是工程要求。",
         11.5, RED, SANS, True, PP_ALIGN.LEFT, 1.5)
    footer(s, "来源：本地实验记录；工具性趋同见 S36")


def build(prs, start):
    pno = start
    divider(prs, "01", "认识 AI", "五个短板，每一个都能从原理上讲清",
            items=[("01", "五个短板"), ("02", "它怎么学"), ("03", "它怎么看句子"),
                   ("04", "三步训练"), ("05", "逐个拆开：抽卡 · 幻觉 · 迎合 · 上下文 · 奖励")],
            quote="这一章不讲“AI 有多强”，只讲它为什么会错 —— 后面每一个案例，都能在这里找到出处。",
            pno=pno, variant="list")
    fn = [p07, p08, p09, p10, p11, p12, p13, p14, p15, p16, p17]
    for i, f in enumerate(fn):
        f(prs, pno + 1 + i)
