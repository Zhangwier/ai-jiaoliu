# -*- coding: utf-8 -*-
"""第 02 章 模型与工具"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "02  模型与工具"

CFG = [("A", "GLM 5.3 Flash", "Trae"), ("B", "Seed 2.1 Turbo", "Trae"),
       ("C", "DeepSeek V4.1 Flash 1M", "WorkBuddy"), ("D", "Opus 5.5", "Hark"),
       ("E", "HY4 Preview 1M high", "WorkBuddy"), ("F", "Step 5 Preview", "WorkBuddy"),
       ("G", "GPT 6.1 Sol high", "Codex"), ("H", "Auto fast", "WorkBuddy"),
       ("I", "Qwen 3.8 Flash", "Qoder"), ("J", "MiniMax M3", "WorkBuddy")]


def div(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.22, 1.5, 1.10, "02", 68, C(0x28, 0x3A, 0x59), SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.36, 1.28, 8.0, 0.30, "PART  02", 12, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=2.2, check=False)
    text(s, ML + 1.36, 1.56, 9.4, 1.32,
         "同一句提示词，\n十个配置给出十种结果。", 34, WHITE, SANS, True, PP_ALIGN.LEFT, 1.28,
         check=False)
    text(s, ML + 1.36, 3.04, 9.2, 0.32,
         "差别不在“AI 行不行”，在档次、专精、外壳这三件事上。", 15, DARKGREY, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    hr(s, ML, 3.62, CW, 1.0, C(0x2C, 0x3A, 0x55))
    IT = [("01", "鹈鹕盲测", "同一句提示词，十个配置"), ("02", "三个差别", "档次、专精、外壳"),
          ("03", "同品牌不同档", "网页版、会员、API 常不是同一个"), ("04", "模型 × 外壳", "六项能力与六类入口"),
          ("05", "榜单怎么读", "换一种考法，第一名就换人"), ("06", "格局与脾气", "谁在哪一档、各自什么性格"),
          ("07", "免费额度", "先用积分跑通，再决定付费")]
    for i, (no, nm, sub) in enumerate(IT):
        c = i // 4
        r = i % 4
        x = ML + c * 4.10
        yy = 3.84 + r * 0.72
        text(s, x, yy + 0.02, 0.46, 0.24, no, 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.46, yy - 0.02, 3.50, 0.28, nm, 15, WHITE, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.46, yy + 0.26, 3.54, 0.24, sub, 11.5, C(0xA8, 0xB2, 0xC4), SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
    return s


def p19(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "一句提示词，十个配置，十份结果",
             "提示词：创建一个 HTML，内容是 SVG 绘制一个鹈鹕骑自行车的 2D 动画，你不需要任何测试。一次生成、无历史记忆。")
    top = y + 0.16
    gx, gy = 0.22, 0.20
    cw = (CW - gx * 4) / 5
    chh = 1.32
    for i, (no, md, cl) in enumerate(CFG):
        c = i % 5
        r = i // 5
        x = ML + c * (cw + gx)
        yy = top + r * (chh + gy + 0.44)
        frame(s, x, yy, cw, chh, LINE)
        pic(s, f"s019_{i+1}.gif", x, yy, cw, chh, mode="fill")
        text(s, x, yy + chh + 0.06, 0.30, 0.22, no, 11.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.26, yy + chh + 0.04, cw - 0.26, 0.22, md, 10.5, NAVY, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.26, yy + chh + 0.26, cw - 0.26, 0.20, cl, 9.5, FAINT, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
    footer(s, "来源：10 份原始 HTML 录屏（统一视口、去黑边）；一次生成、无历史记忆")
    return s


def p20(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "十种结果的差别，来自三件不同的事",
             "同一个提示词，为什么结果差这么远？把差别按来源拆开，选型才有依据。")
    ty = y + 0.20
    C3 = [("01", "档次", "同一个壳里，模型本身有多强",
           [("例", "GPT-6.1 Sol 与 GPT-6 Luna 在 AA 智能指数上差 20 分以上"),
            ("所以", "预算够就上高一档，别在低档模型上堆工作流"),
            ("自测", "同一个任务，两个档各跑一次，看差多少")], C(0x4A, 0x6C, 0x9E)),
          ("02", "专精", "这个模型被练得最熟的是哪一类活",
           [("例", "修编报告最快 4 分钟、最慢 34 分钟，还有一份套错了条目"),
            ("所以", "写报告、读图纸、写代码，最优选择不是同一个"),
            ("自测", "拿自己最常做的一类活做小样本对比")], TEAL),
          ("03", "外壳", "包在模型外面的那层工具怎么喂材料",
           [("例", "只靠 OCR 的两个配置，两个检查点零命中"),
            ("所以", "能不能直接看图，比分数更影响图纸类任务"),
            ("自测", "问它一句“你看到的是原图还是识别文本”")], ORANGE)]
    cw = (CW - 0.56) / 3
    for i, (no, nm, lead, rows, col) in enumerate(C3):
        x = ML + i * (cw + 0.28)
        box(s, x, ty, cw, 0.055, fill=col)
        text(s, x, ty + 0.16, 0.46, 0.26, no, 11.5, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.46, ty + 0.12, cw - 0.46, 0.34, nm, 21, NAVY, SANS, True, PP_ALIGN.LEFT,
             1.15, check=False)
        text(s, x, ty + 0.56, cw, 0.28, lead, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        hr(s, x, ty + 0.92, cw, 1.0, LINE)
        yy = ty + 1.04
        for tag, body in rows:
            text(s, x, yy, 0.44, 0.22, tag, 10.5, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
                 check=False)
            nl = nlines(body, 12.5, cw - 0.52)
            text(s, x + 0.48, yy - 0.02, cw - 0.48, nl * 12.5 * 1.32 * 1.42 / 72 + 0.05, body,
                 12.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.42, check=False)
            yy += 0.26 + nl * 12.5 * 1.32 * 1.42 / 72 + 0.22
    box(s, ML, 5.86, CW, 0.70, fill=MINT)
    text(s, ML + 0.24, 5.97, CW - 0.48, 0.48,
         "三件事里，只有“外壳”是你今天就能换的 —— 换壳比换模型便宜得多。", 14.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.25, check=False)
    footer(s, "来源：本地鹈鹕盲测与图纸审查记录；AA 智能指数 v4.3.2")
    return s


def p21(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一品牌里，免费版和付费版常常不是同一个模型",
             "同名产品，网页免费版、会员版、API 可能指向三个不同的模型 —— 分数差 12 到 15 分。")
    ty = y + 0.24
    G = [("ChatGPT", "免费 GPT-6 Luna  ·  上下文 27K", "付费 GPT-6.1 Sol", 38, 52),
         ("Claude", "Haiku 5.5", "Opus 5.5", 43, 58),
         ("Gemini", "3.8 Flash", "4 Argon", 41, 53)]
    for i, (brand, lo, hi, slo, shi) in enumerate(G):
        yy = ty + i * 1.26
        if i:
            hr(s, ML, yy - 0.10, 6.30, 1.0, LINE)
        text(s, ML, yy, 2.10, 0.30, brand, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML, yy + 0.34, 3.10, 0.26, lo, 12, GREY2, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML, yy + 0.62, 3.10, 0.26, "→ " + hi, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
    # 右：分数轴
    AX = 7.10
    vr(s, AX - 0.32, ty - 0.06, 3.90, 1.0, LINE)
    text(s, AX, ty - 0.06, 5.00, 0.24, "AA 智能指数（同一口径下对比）", 11, GREY2, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    x0, x1 = AX + 0.30, X2 - 0.30
    v0, v1 = 32.0, 62.0

    def fx(v):
        return x0 + (v - v0) / (v1 - v0) * (x1 - x0)

    for g in [36, 42, 48, 54, 60]:
        vr(s, fx(g), ty + 0.26, 3.02, 0.6, C(0xE4, 0xDF, 0xD3))
    for i, (brand, lo, hi, slo, shi) in enumerate(G):
        yy = ty + 0.72 + i * 1.06
        line(s, fx(slo), yy, fx(shi), yy, LINE2, 1.6, head=True)
        circ(s, fx(slo), yy, 0.15, C(0xD5, 0xD9, 0xDF), CARD, 1.2)
        circ(s, fx(shi), yy, 0.17, TEAL)
        text(s, fx(slo) - 0.36, yy + 0.14, 0.86, 0.24, str(slo), 11, GREY2, SANS, True,
             PP_ALIGN.CENTER, 1.0, check=False)
        text(s, fx(shi) - 0.34, yy + 0.14, 0.80, 0.24, str(shi), 11, TEAL, SANS, True,
             PP_ALIGN.CENTER, 1.0, check=False)
        text(s, fx(shi) + 0.24, yy - 0.14, 1.20, 0.26, f"+{shi-slo} 分", 13, ORANGE, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
    for g in [36, 42, 48, 54, 60]:
        text(s, fx(g) - 0.24, ty + 3.34, 0.48, 0.22, str(g), 10, FAINT, SANS, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    text(s, AX, ty + 3.62, 5.00, 0.24, "横轴：AA 智能指数（越高越好）", 10, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    takeaway(s, "所以问“AI 行不行”没用，得先问“是哪一个模型的哪一个入口”。", y=5.90, tag="判断")
    footer(s, "来源：AA 智能指数 v4.3.2；各产品公开定价页与模型说明")
    return s


def p22(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "实际效果 = 模型 × 外壳",
             "模型决定“能想多深”，外壳决定“喂得进什么、交得出什么”。两项都看，才解释得了那十份结果。")
    ty = y + 0.18
    text(s, ML, ty, 5.40, 0.24, "模型：六项能力", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    MX = [("上下文窗口", "一次能装多少字"), ("最大输出", "一次能写多长"),
          ("输入模态", "能不能读图、读表、读 PDF"), ("输出模态", "只出文字还是也出文件"),
          ("真能看见图吗", "原图还是 OCR 文本"), ("记忆与工具", "会不会记、能不能跑程序")]
    for i, (t, b) in enumerate(MX):
        x = ML + (i % 2) * 2.64
        yy = ty + 0.34 + (i // 2) * 0.92
        box(s, x, yy + 0.02, 0.055, 0.58, fill=TEAL)
        text(s, x + 0.20, yy, 2.40, 0.26, t, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.20, yy + 0.28, 2.38, 0.34, b, 11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
    vr(s, 6.62, ty + 0.10, 3.30, 1.0, LINE)
    text(s, 6.40, ty - 0.02, 0.46, 0.30, "×", 22, LINE2, SANS, True, PP_ALIGN.CENTER, 1.0,
         check=False)
    RX = 6.92
    text(s, RX, ty, 5.80, 0.24, "外壳：六类入口", 11.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    SH = [("聊天网页 / App", "DeepSeek、千问、ChatGPT", "最省事，看不见文件系统"),
          ("编辑器插件", "CodeBuddy、Trae、Cursor、Cline", "能读整个工程，配置复杂"),
          ("终端 Agent", "Codex、Claude Code、OpenCode", "权限最大，最需要边界"),
          ("桌面办公 Agent", "WorkBuddy、Claude 桌面版", "直接改 Excel、PPT"),
          ("常驻 Agent", "Hermes、OpenClaw", "长期跑、有记忆、要人盯"),
          ("云端代理", "Devin、Jules、Hark、Muse", "交出去就不看了，最难查")]
    rows = [["类型", "代表", "特点"]] + [[a, b, c] for a, b, c in SH]
    table(s, RX, ty + 0.32, ML + CW - RX, rows, [1.44, 2.44, 2.32], size=10, head_size=10,
          minrow=0.42, head_h=0.28, first_bold=True, pad=0.09, ls=1.1,
          colors=lambda r, c: TEAL if c == 0 else (GREY if c == 2 else C(0x2B, 0x34, 0x45)))
    footer(s, "来源：各家产品文档与本地实测；外壳分类为本稿归纳")
    return s


def p23(prs, pno):
    """AA 智能指数散点：横轴每题成本（对数），纵轴指数"""
    import math
    s = slide(prs)
    y = head(s, CH, pno, "分数和价格不成正比：贵的不一定更强",
             "横轴是每道题的成本（对数刻度），纵轴是 AA 智能指数。越靠左上，性价比越高。")
    # 只标 8 个关键点，其余作背景群
    KEY = [("Opus 5.5", 58, 0.72, "L"), ("Sonnet 5.5", 56, 0.36, "L"),
           ("GPT-6 Astra", 53, 1.05, "R"), ("GPT-6.1 Sol", 52, 0.72, "L"),
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
    text(s, x0 + 0.06, y1 - 0.02, 1.40, 0.22, "智能指数", 9.5, GREY2, SANS, True, PP_ALIGN.LEFT,
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
    yy = 2.50
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
         "灰点是同一批里的其他模型：\n多数集中在 40–46 分、成本不到 $0.2。", 10.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.35)
    footer(s, "来源：AA 智能指数 v4.3.2（每道题成本为公开定价换算，美元）")
    return s


def p24(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "换一种考法，第一名就换人",
             "没有“最强模型”，只有“在某一种考法上最强”。看榜单先看它考的是什么。")
    rows = [["榜单", "主要测什么", "第一名", "紧随其后"],
            ["Epoch 综合", "跨多个基准加权", "Claude Opus 5.5", "GPT-6 Astra 几乎并列"],
            ["FrontierMath", "研究级数学", "GPT-6 Astra  97.6%", "Claude Opus 5.5"],
            ["ARC-AGI-3", "抽象推理、少样本", "GPT-6 Astra  99.9%", "Gemini 4 Argon  97.5%"],
            ["LiveBench 推理", "逻辑与推理", "GPT-6 Astra  92.7%", "Claude Opus 5.5"],
            ["AA 智能指数", "综合智能 + 成本", "Claude Opus 5.5  58", "Sonnet 5.5  56"],
            ["LMArena 文字", "人类盲评偏好", "Gemini 4 Argon  1524", "GPT-6 Astra"],
            ["LMArena 智能体", "多步工具使用", "Claude Opus 5.5", "GPT-6 Astra"]]
    ty = y + 0.16
    table(s, ML, ty, CW, rows, [2.30, 3.20, 3.30, 3.20], size=13, head_size=13,
          minrow=0.50, head_h=0.42, rowh=0.50, first_bold=True, pad=0.14,
          colors=lambda r, c: (ORANGE if c == 2 else (GREY if c == 1 else C(0x2B, 0x34, 0x45))))
    takeaway(s, "所以榜单的正确用法是：先看你这类活被哪个榜考到了，再看那个榜的第一名是谁。", y=6.10,
             tag="判断")
    footer(s, "来源：Epoch AI、FrontierMath、ARC-AGI-3、LiveBench、AA、LMArena 公开榜单")
    return s


def p25(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "美国两强领跑，国产头部落后约十分",
             "按 AA 智能指数分层：第一梯队 52 分以上，第二梯队 44–46 分，其后 29 分上下。")
    ty = y + 0.16
    LAY = [("第一梯队", "OpenAI · Anthropic", "Astra 53、Sol 52 ／ Opus 58、Sonnet 56", 1.00,
            C(0x4A, 0x6C, 0x9E)),
           ("紧随其后", "Google · Meta", "Argon 53 ／ Muse Wonder 52", 0.84, TEAL),
           ("第二梯队", "xAI · 国产头部",
            "Grok 4.7 46、MiMo 46、Qwen3.8 Max 45、GLM-5.3 45、Kimi K3 44", 0.66,
            C(0x63, 0x71, 0x88)),
           ("其后", "腾讯 · MiniMax · 字节", "Hy4 未进同口径榜；MiniMax-M3 29；豆包无同口径数据",
            0.40, C(0x9A, 0xA1, 0xAC))]
    BX, BW = 2.46, 10.24
    for i, (nm, who, models, frac, col) in enumerate(LAY):
        yy = ty + i * 0.94
        text(s, ML, yy - 0.02, 2.00, 0.28, nm, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML, yy + 0.28, 2.02, 0.24, who, 9.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        box(s, BX, yy + 0.02, BW, 0.28, fill=C(0xF0, 0xEE, 0xE7))
        bw = BW * frac
        box(s, BX, yy + 0.02, bw, 0.28, fill=col)
        room = bw - 0.28
        if i == 3:
            text(s, BX, yy + 0.34, BW, 0.24, models, 10, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
        else:
            text(s, BX + 0.14, yy + 0.075, room, 0.24, models, 10.5, WHITE, SANS, True,
                 PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 6.04, CW, 1.0, LINE)
    text(s, ML, 6.16, 6.10, 0.28, "另一条值得注意：AA 开源前六名，全部是中国模型", 12.5, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 6.44, 6.10, 0.26, "闭源看两强，开源看国产 —— 这是当前的格局", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML + 6.84, 6.16, 5.28, 0.26,
         "分层比的是“最好那一档”，不是平均水位。", 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
         check=False)
    text(s, ML + 6.84, 6.44, 5.28, 0.26,
         "日常任务里，外壳选错造成的差距常常大于档位差距。", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：AA 智能指数 v4.3.2；各厂商公开模型说明")
    return s


def p26(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "脾气是真的：按任务挑，比按总分挑靠谱",
             "同一档位的模型，做事风格差别很大。用熟了，选型判断比看分数快。")
    rows = [["模型", "使用体感", "能对上的证据"],
            ["GPT 系", "防御式写代码，先加校验再干活；推理最硬，但话少、不爱解释",
             "FrontierMath 97.6%、ARC-AGI-3 99.9% 均为第一"],
            ["Claude 系", "最听得懂人话，会主动查证、主动说“这里我不确定”", "综合与智能体榜单第一；写动画会自己截图验证"],
            ["Gemini 系", "知识面最广、中文最顺，但很少反对你", "LMArena 文字第一 1524；上下文可到 1M"],
            ["DeepSeek 系", "出字最快、最便宜，但先大段思考再答", "V4.1 Flash 约 217 token/s，输出 8 元/百万 token"],
            ["豆包 / 元宝", "彻底顺着你，适合改稿润色，不适合找错", "同口径榜单无数据；实测中不主动反驳"]]
    ty = y + 0.16
    table(s, ML, ty, CW, rows, [1.90, 6.10, 4.10], size=12.5, head_size=12.5, minrow=0.60,
          head_h=0.42, rowh=0.62, first_bold=True, pad=0.16,
          colors=lambda r, c: TEAL if c == 0 else (GREY if c == 2 else C(0x2B, 0x34, 0x45)))
    takeaway(s, "“脾气”不是性格，是训练取舍的结果 —— 也就意味着它稳定、可预期、可利用。", y=6.10,
             tag="判断")
    footer(s, "来源：本地实测记录；AA 智能指数与 FrontierMath / ARC-AGI-3 / LMArena 公开数据")
    return s


def p27(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "先用免费积分跑通一个任务，再决定付不付",
             "八种常见入口的入门价与免费额度。价格随时会变，看的是量级和计费方式。")
    rows = [["入口", "免费额度", "入门价", "适合"],
            ["WorkBuddy", "—", "¥99/月（连续包月 ¥70），500 积分", "桌面办公、改 Excel/PPT"],
            ["Trae 国内版", "—", "¥49/月起，500 积分", "写代码、读工程"],
            ["Qoder", "2 周试用 + 300 credits", "$20/月 或 ¥59/月", "代码补全与重构"],
            ["Cline（插件）", "开源免费", "$9.99/月", "自选模型、可控成本"],
            ["Codex", "免费号可用小模型", "$20/月", "终端里跑多步任务"],
            ["Kiro", "50 credits", "$20/月", "规格驱动的开发"],
            ["Claude Code", "无免费额度", "$20/月", "长任务、大工程"],
            ["ChatGPT / Claude / Gemini", "有限次免费对话", "$20/月（Pro 档 $100）", "通用问答与写作"]]
    ty = y + 0.14
    table(s, ML, ty, CW, rows, [2.10, 2.70, 3.40, 3.90], size=12, head_size=12,
          minrow=0.40, head_h=0.38, rowh=0.40, first_bold=True, pad=0.14,
          colors=lambda r, c: (TEAL if (c == 2 and r > 0) else
                               (GREY if c in (1, 3) else C(0x2B, 0x34, 0x45))))
    text(s, ML, 6.06, CW, 0.52,
         "顺序建议：先用免费积分把一个小任务跑通 → 核结果 → 再决定换模型或付费。反过来先买再找场景，多半是用两次就放着。",
         13, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.4)
    footer(s, "来源：各产品公开定价页（2026-10）；价格与额度以官方页面为准")
    return s


def build(prs, start):
    div(prs, start)
    fn = [p19, p20, p21, p22, p23, p24, p25, p26, p27]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
