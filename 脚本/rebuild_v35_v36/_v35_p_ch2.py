# -*- coding: utf-8 -*-
"""第 02 章 模型与工具"""
import math
from ds import *
from p_common import divider
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "02   模型与工具"


# ---------------------------------------------------------------- 章节页（右构图）
def div02(prs, pno):
    s = slide(prs, DARK)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.30, 6.0, 0.30, "PART  02", 9.5, C(0x7E, 0x88, 0x96), SANS, False,
         PP_ALIGN.LEFT, 1.0, cs=2.6, check=False)
    text(s, ML - 0.05, 1.62, 7.6, 1.90,
         "同一句提示词，\n十个配置给出十种结果。", 42, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.42,
         check=False)
    text(s, ML, 4.30, 6.55, 1.10,
         "十份结果不是“谁画得好看”，是三件事叠在一起：用的哪一档、专不专这一项、壳撑不撑得住。",
         12.5, C(0xB2, 0xBB, 0xC7), SANS, False, PP_ALIGN.LEFT, 1.62)

    # 右侧纵向目录
    x0 = 8.62
    items = [("01", "十份鹈鹕，一次生成"), ("02", "三个差别：档次 / 专精 / 外壳"),
             ("03", "同品牌不同档，差多少"), ("04", "模型 × 外壳 = 实际效果"),
             ("05", "榜单怎么读"), ("06", "格局与脾气"), ("07", "免费额度")]
    yy = 1.62
    for k, v in items:
        text(s, x0, yy, 0.5, 0.22, k, 9, C(0x6E, 0x78, 0x86), MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x0 + 0.52, yy - 0.04, 3.3, 0.28, v, 11.5, C(0xD6, 0xDC, 0xE4), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.52
    vr(s, x0 - 0.35, 1.60, 3.60, 0.75, C(0x38, 0x40, 0x4E))
    footer(s, "来源：本地十份原始 HTML 录屏；各产品官网", dark=True)
    return s


# ---------------------------------------------------------------- 19 十宫格
def p19(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "一句提示词，十个配置，十份结果",
             "提示词：“创建一个 HTML，内容是 SVG 绘制一个鹈鹕骑自行车的 2D 动画，你不需要任何测试”")
    names = [("A", "GLM 5.3 Flash", "Trae"), ("B", "Seed 2.1 Turbo", "Trae"),
             ("C", "DeepSeek V4.1 Flash 1M", "WorkBuddy"), ("D", "Opus 5.5", "Hark"),
             ("E", "HY4 Preview 1M high", "WorkBuddy"), ("F", "Step 5 Preview", "WorkBuddy"),
             ("G", "GPT 6.1 Sol high", "Codex"), ("H", "Auto fast", "WorkBuddy"),
             ("I", "Qwen 3.8 Flash", "Qoder"), ("J", "MiniMax M3", "WorkBuddy")]
    gap = 0.24
    cw = (CW - gap * 4) / 5
    ih = cw * 0.75
    top = 2.16
    for i, (lt, mdl, cli) in enumerate(names):
        cx = ML + (i % 5) * (cw + gap)
        cy = top + (i // 5) * (ih + 0.66)
        frame(s, cx, cy, cw, ih, RULE, 0.75, fill=C(0xEC, 0xE9, 0xE1))
        pic(s, f"s019_{i+1}.gif", cx, cy, cw, ih, mode="fill")
        text(s, cx, cy + ih + 0.10, 0.34, 0.22, lt, 10, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cx + 0.30, cy + ih + 0.09, cw, 0.24, mdl, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, cx + 0.30, cy + ih + 0.32, cw, 0.22, cli, 9, MUT, SANS, False, PP_ALIGN.LEFT, 1.1, check=False)
    footer(s, "来源：10 份原始 HTML 录屏（统一视口，去黑边）；一次生成、无历史记忆")


# ---------------------------------------------------------------- 20 三个差别
def p20(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "十种结果的差别，来自三件不同的事",
             "先认清是哪一种差别，再决定要不要换 —— 混在一起看，只会得出“AI 不行”。")
    lab_w = 0.66
    cw = 3.20
    gap = 0.44
    cols = [ML + lab_w + i * (cw + gap) for i in range(3)]
    heads = [("01", "档次差别", "同一产品线，分数差很多"),
             ("02", "专精差别", "同一档位，各有所长"),
             ("03", "外壳差别", "同一模型，换壳就换能力")]
    for i, (no, ttl, sub) in enumerate(heads):
        text(s, cols[i], 2.10, 0.5, 0.22, no, 10, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cols[i], 2.36, cw, 0.34, ttl, 19, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, cols[i], 2.74, cw - 0.1, 0.24, sub, 10, MUT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        hr(s, cols[i], 3.06, cw - 0.18, 1.6, RED if i == 0 else (GRN if i == 1 else BLU))

    rows = [
        ("例", 3.20, [
            "GPT-6.1 Sol 与 GPT-6 Luna，AA 智能指数差 20 分以上",
            "修编报告：最快 4 分钟做完，最慢 34 分钟还套错了条目",
            "只靠 OCR 的配置零命中；命中的配置都真的看到了图"]),
        ("所以", 4.24, [
            "别用最便宜的一档判断“AI 行不行”",
            "按任务挑，别拿 A 的短板否定 B",
            "模型一样时，先换壳"]),
        ("自测", 4.98, [
            "换贵一档，跑同一条提示词",
            "同一任务交两家，看谁返工少",
            "换一个能读原文件的壳"]),
    ]
    for tag, yy, texts in rows:
        text(s, ML, yy + 0.02, lab_w, 0.24, tag, 10, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
        for i, t in enumerate(texts):
            col = RED if tag == "所以" else BODY
            bold = (tag == "所以")
            text(s, cols[i], yy, cw - 0.12, 0.86, t, 11.5, col, SANS, bold, PP_ALIGN.LEFT, 1.55)

    hr(s, ML, 5.86, CW, 0.75, RULE)
    text(s, ML, 6.02, 2.2, 0.24, "一句话", 10, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + lab_w + 0.06, 6.00, CW - lab_w - 0.06, 0.32,
         "觉得“AI 不行”之前，按这个顺序查一遍：档位是不是太低、是不是用错了专精、壳是不是不匹配。",
         13, INK, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：第 19 页十份鹈鹕实测；AA 智能指数（2026-10-18 查询）；本次客户端实测记录")


# ---------------------------------------------------------------- 21 同品牌不同档
def p21(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一品牌里，免费版和付费版常常不是同一个模型",
             "拿免费版的体验去评价一个模型，常常低估了它。比较之前，先确认用的是哪个型号。")
    rows = [("ChatGPT", "GPT-6 Luna（上下文 27 K）", "GPT-6.1 Sol", 38, 52),
            ("Claude", "Sonnet / Haiku", "Opus、Claude Code", 43, 58),
            ("Gemini", "Flash-Lite", "Gemini 4 Argon（AI Pro）", 41, 53)]
    x0, x1 = 5.00, 10.40          # 分数轴范围 34 → 62
    lo, hi = 34.0, 62.0

    def fx(v):
        return x0 + (v - lo) / (hi - lo) * (x1 - x0)

    # 轴
    for t in (35, 40, 45, 50, 55, 60):
        vr(s, fx(t), 2.20, 3.02, 0.6, C(0xE4, 0xDF, 0xD3))
        text(s, fx(t) - 0.30, 5.28, 0.60, 0.22, str(t), 8.5, FAINT, MONO, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    text(s, x0 - 0.42, 5.28, 1.6, 0.22, "AA 智能指数", 8.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, x1 - 1.9, 5.52, 1.9, 0.22, "（越高越好）", 8.5, FAINT, SANS, False,
         PP_ALIGN.RIGHT, 1.0, check=False)

    yy = 2.34
    for nm, free, paid, s1, s2 in rows:
        text(s, ML, yy, 3.6, 0.30, nm, 16, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML, yy + 0.36, 3.72, 0.50,
             ("免费 " + free + "\n付费 " + paid), 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.5)
        # 连接线
        hr(s, fx(s1), yy + 0.16, fx(s2) - fx(s1), 2.4, C(0xD5, 0xCF, 0xC1))
        circ(s, fx(s1), yy + 0.16, 0.11, C(0xB9, 0xB2, 0xA3))
        circ(s, fx(s2), yy + 0.16, 0.13, GRN)
        text(s, fx(s1) - 0.24, yy - 0.12, 0.5, 0.22, str(s1), 10, MUT, MONO, True,
             PP_ALIGN.RIGHT, 1.0, check=False)
        text(s, fx(s2) + 0.12, yy - 0.14, 0.62, 0.24, str(s2), 12, GRN, MONO, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        # 差值
        text(s, fx(s2) + 0.78, yy + 0.02, 1.3, 0.28, f"+{s2-s1} 分", 10.5, GRN, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        yy += 1.18

    hr(s, ML, 5.72, CW, 0.75, RULE)
    text(s, ML, 5.86, 11.2, 0.62,
         "同一品牌内部的分差，往往比“国产与国外”的差距还大 —— 所以比较对象必须写清楚到型号。",
         12, INK, SANS, True, PP_ALIGN.LEFT, 1.45)
    footer(s, "来源：chatgpt.com/pricing、claude.com/pricing、gemini.google/subscriptions；Artificial Analysis")


# ---------------------------------------------------------------- 22 模型 × 外壳
def p22(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "实际效果 = 模型 × 外壳",
             "模型决定上限，外壳决定它此刻能干什么。比较两个结果时，这两样都要写清楚。")
    # 左：模型六项
    text(s, ML, 2.16, 4.2, 0.24, "模型（LLM）：看六项", 9.5, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    six = [("上下文", "一次装多少"), ("最大输出", "一次写多长"), ("输入模态", "能不能读图、读文件"),
           ("输出模态", "能产出什么"), ("真能看见图吗", "不是“支持上传”"), ("记忆 / 推理与工具", "跨对话记住啥、想多久")]
    for i, (a, b) in enumerate(six):
        cx = ML + (i % 2) * 2.18
        cy = 2.50 + (i // 2) * 1.06
        vr(s, cx - 0.16, cy + 0.02, 0.62, 1.4, GRN)
        text(s, cx + 0.04, cy - 0.02, 2.05, 0.26, a, 11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, cx + 0.04, cy + 0.30, 2.02, 0.42, b, 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.35)

    # 中：×
    text(s, 4.90, 3.30, 1.0, 0.80, "×", 40, RULE2, SANS, False, PP_ALIGN.CENTER, 1.0, check=False)

    # 右：外壳六类
    x0 = 5.92
    text(s, x0, 2.16, 5.6, 0.24, "外壳（Harness）：在哪跑、能跑多久", 9.5, BLU, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    rows = [("聊天网页 / App", "DeepSeek、千问、ChatGPT", "一问一答"),
            ("编辑器 / 插件", "CodeBuddy、Trae、Cursor、Cline", "在编辑器里改文件"),
            ("终端 Agent", "Codex、Claude Code、OpenCode", "读写整个文件夹、跑程序"),
            ("桌面办公 Agent", "WorkBuddy、Claude 桌面版", "直接处理本机 Excel、PDF"),
            ("常驻 Agent", "Hermes Agent、OpenClaw", "长期运行、接聊天软件、有记忆"),
            ("云端 / 个人代理", "Devin、Jules、Hark、Muse", "云端电脑替你跑任务")]
    yy = 2.48
    for a, b, c in rows:
        text(s, x0, yy, 1.62, 0.24, a, 10.8, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x0 + 1.62, yy, 2.62, 0.24, b, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x0 + 4.28, yy, 1.62, 0.40, c, 9.2, MUT, SANS, False, PP_ALIGN.LEFT, 1.25, check=False)
        yy += 0.62
    # 第二行进第二位（小）
    footer(s, "来源：各产品官网与 GitHub（2026-10-18 查询）；本分享案例")


# ---------------------------------------------------------------- 23 AA 散点
def p23(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "分数和价格不成正比：贵的不一定更强",
             "AA 智能指数 v4.3.2 偏重智能体与办公任务；横轴为每道题成本（美元，对数刻度）。")
    PTS = [("Opus 5.5", 58, 8.34), ("Sonnet 5.5", 56, 11.39), ("GPT-6 Astra", 53, 8.68),
           ("Gemini 4 Argon", 53, 6.02), ("Fable 5.1", 53, 11.35), ("GPT-6.1 Sol", 52, 3.87),
           ("Muse Wonder", 52, 9.27), ("MiMo-V2.6-Pro", 46, 3.49), ("Grok 4.7", 46, 7.35),
           ("GLM-5.3", 45, 6.04), ("Qwen3.8 Max", 45, 11.37), ("Step 5 Preview", 44, 4.70),
           ("Kimi K3", 44, 9.27), ("Haiku 5.5", 43, 1.26), ("GLM-5.3-Flash", 42, 5.11),
           ("Gemini 3.8 Flash", 41, 8.26), ("Qwen3.8-Flash-Next", 41, 5.70),
           ("DeepSeek V4.1", 39, 1.79), ("GPT-6 Luna", 38, 0.50), ("MiniMax-M3", 29, 6.38)]

    def cost(xs):
        return 10 ** ((xs - 7.28) / 4.86)

    PX0, PX1 = 1.42, 8.32
    PY0, PY1 = 2.16, 5.86
    LO, HI = -1.55, 1.15          # log10(cost)
    SLO, SHI = 26.0, 61.0

    def fx(c):
        return PX0 + (math.log10(c) - LO) / (HI - LO) * (PX1 - PX0)

    def fy(sc):
        return PY0 + (SHI - sc) / (SHI - SLO) * (PY1 - PY0)

    # 网格
    for t in (0.05, 0.1, 0.5, 1, 5, 10):
        vr(s, fx(t), PY0, PY1 - PY0, 0.55, C(0xE2, 0xDD, 0xD1))
        lab = f"${t:g}" if t < 1 else f"${t:g}"
        text(s, fx(t) - 0.30, PY1 + 0.08, 0.60, 0.20, lab, 8.2, FAINT, MONO, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    for t in (30, 40, 50, 60):
        hr(s, PX0, fy(t), PX1 - PX0, 0.55, C(0xE2, 0xDD, 0xD1))
        text(s, PX0 - 0.52, fy(t) - 0.09, 0.44, 0.20, str(t), 8.2, FAINT, MONO, False,
             PP_ALIGN.RIGHT, 1.0, check=False)
    hr(s, PX0, PY1, PX1 - PX0, 1.0, RULE2)
    text(s, PX0, PY1 + 0.30, 4.2, 0.22, "每道题成本（美元，每格贵 10 倍）", 8.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, PX0 - 0.86, PY0 - 0.24, 1.4, 0.22, "智能指数", 8.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)

    # 点 + 标签（避让）
    taken = []
    def overlap(a, b):
        return not (a[0] + a[2] <= b[0] or b[0] + b[2] <= a[0] or a[1] + a[3] <= b[1] or b[1] + b[3] <= a[1])
    order = sorted(PTS, key=lambda p: -p[1])
    for nm, sc, xs in order:
        c = cost(xs)
        px, py = fx(c), fy(sc)
        circ(s, px, py, 0.085, INK)
        lw = est_w(nm, 8.0) / 72.0
        cand = [(0.09, -0.085), (0.09, 0.035), (-lw - 0.09, -0.085), (-lw - 0.09, 0.035),
                (0.09, 0.20), (0.09, -0.245)]
        chosen = None
        for dx, dy in cand:
            r = (px + dx, py + dy, lw, 0.15)
            if r[0] < PX0 - 0.28 or r[0] + r[2] > PX1 + 0.30 or r[1] < PY0 - 0.20 or r[1] > PY1:
                continue
            if any(overlap(r, t) for t in taken):
                continue
            chosen = (r, dx, dy); break
        if chosen is None:
            chosen = ((px + 0.09, py - 0.085, lw, 0.15), 0.09, -0.085)
        r, dx, dy = chosen
        taken.append(r)
        col = RED if nm == "GPT-6.1 Sol" else BODY
        bd = (nm == "GPT-6.1 Sol")
        text(s, r[0], r[1], lw + 0.06, 0.16, nm, 8.0, col, SANS, bd, PP_ALIGN.LEFT, 1.0, check=False)

    # 右侧说明
    x0 = 9.30
    vr(s, x0 - 0.42, 2.16, 3.60, 0.75, RULE)
    text(s, x0, 2.16, 3.2, 0.24, "怎么读这张图", 9.5, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    tips = [("越往左上", "越划算 —— 同样分数，成本更低"),
            ("最左一列", "便宜十倍，分数只低几分"),
            ("横轴最右", "贵，但不一定强；用于少数硬任务"),
            ("标注", "GPT-6.1 Sol 52 分只要 $0.72 / 题")]
    yy = 2.52
    for a, b in tips:
        text(s, x0, yy, 3.06, 0.24, a, 10.5, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x0, yy + 0.28, 3.06, 0.50, b, 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.4)
        yy += 0.86
    text(s, x0, 6.16, 3.06, 0.42,
         "榜单只帮你挑候选，不替你做决定。", 10.5, RED, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：artificialanalysis.ai/leaderboards/models；DeepSeek 官方价格页（2026-10-18 查询）")


# ---------------------------------------------------------------- 24 榜单
def p24(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "换一种考法，第一名就换人",
             "2026-10-13 至 10-18 查询；分差在误差范围内的，按“并列”看。")
    rows = [("Epoch 综合指数", "50 多项评测合成", "Claude Opus 5.5", "GPT-6 Astra（几乎并列）"),
            ("FrontierMath", "研究级数学题", "GPT-6 Astra  97.6%", "Erdős 未解题：唯一得分"),
            ("ARC-AGI-3", "全新抽象推理", "GPT-6 Astra  99.9%", "Gemini 4 Argon  97.5%"),
            ("LiveBench 推理", "推理题，定期换题", "GPT-6 Astra  92.7%", "—"),
            ("AA 智能指数", "偏智能体、办公任务", "Claude Opus 5.5  58", "Sonnet 5.5  56，Astra  53"),
            ("LMArena 文字", "人类盲投偏好", "Gemini 4 Argon", "Muse Wonder、Opus 5.5"),
            ("LMArena 智能体", "人类评智能体干活", "Claude Opus 5.5", "GPT-6.1 Sol、Fable 5.1")]
    cw = [3.10, 2.55, 3.30, None]
    xs = [ML, ML + 3.18, ML + 5.86, ML + 9.28]
    text(s, xs[0], 1.98, 3.0, 0.22, "榜单", 9.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, xs[1], 1.98, 2.4, 0.22, "主要测什么", 9.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, xs[2], 1.98, 3.2, 0.22, "第一名", 9.2, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, xs[3], 1.98, 3.2, 0.22, "紧随其后", 9.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    hr(s, ML, 2.26, CW, 1.1, RULE2)
    yy = 2.38
    for i, (a, b, c, d) in enumerate(rows):
        text(s, xs[0], yy, 3.05, 0.26, a, 11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, xs[1], yy + 0.01, 2.55, 0.26, b, 9.8, MUT, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, xs[2], yy, 3.30, 0.26, c, 11.5, RED, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, xs[3], yy + 0.01, 3.10, 0.26, d, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        hr(s, ML, yy + 0.44, CW, 0.6, RULE)
        yy += 0.50
    # 右下角一条判断
    text(s, ML, 6.02, 2.4, 0.24, "一句话", 10, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 1.10, 5.98, 10.30, 0.34,
         "推理看 Astra，干活看 Claude，聊天偏好投 Gemini、Muse —— 不是“总分谁最高”。",
         13, INK, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：epoch.ai/benchmarks；arcprize.org/blog/astra；arena.ai/leaderboard；artificialanalysis.ai；LiveBench")


# ---------------------------------------------------------------- 25 格局
def p25(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "美国两强领跑，国产头部落后约十分",
             "综合 Epoch、AA、LMArena 三个榜单的位置分层；括号内为 AA 智能指数 v4.3.2。")
    tiers = [
        ("第一梯队", "OpenAI · Anthropic", "GPT-6 Astra（53）、GPT-6.1 Sol（52）　Claude Opus 5.5（58）、Sonnet 5.5（56）", 1.00, RED),
        ("紧随其后", "Google · Meta", "Gemini 4 Argon（53，人类盲投第一）　Muse Wonder（52）", 0.86, RED),
        ("第二梯队", "xAI · 国产头部", "Grok 4.7（46）、MiMo（46）、Qwen3.8 Max（45）、GLM-5.3（45）、Kimi K3、Step 5（44）", 0.70, GRN),
        ("其后", "国产追赶", "腾讯 Hy4（智能体榜与 DeepSeek Flash 相近）、MiniMax-M3（29）；豆包无同口径分数", 0.52, GRN),
    ]
    yy = 2.16
    for tag, who, models, frac, col in tiers:
        text(s, ML, yy + 0.02, 1.30, 0.24, tag, 10, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        text(s, ML + 1.36, yy, 3.05, 0.30, who, 13.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML + 4.48, yy + 0.02, 6.92, 0.52, models, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        # 强弱条
        box(s, ML + 1.36, yy + 0.44, 3.05 * frac, 0.05, fill=col)
        hr(s, ML, yy + 0.86, CW, 0.6, RULE)
        yy += 0.98
    # 右侧竖排注
    hr(s, ML, 6.02, CW, 0.75, RULE)
    text(s, ML, 6.16, 6.35, 0.52,
         "AA 开源模型前六名全是中国模型。DeepSeek 旧版 V4.1 Flash 只有 39 分，但最便宜、最快。",
         10.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.45)
    text(s, ML + 6.85, 6.12, 4.68, 0.54,
         "最强选 OpenAI / Claude；国产 GLM、Qwen、Kimi、DeepSeek 够日常用。",
         11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.42)
    footer(s, "来源：artificialanalysis.ai；epoch.ai/eci；arena.ai/leaderboard；huggingface.co/datasets/lmarena-ai（2026-10-16）")


# ---------------------------------------------------------------- 26 脾气
def p26(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "脾气是真的：按任务挑，比按总分挑靠谱",
             "左边是分享人与社区的使用体感，不是评测结论；右边是能对上的公开证据或本次案例。")
    text(s, ML, 1.96, 3.0, 0.22, "模型", 9.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 1.62, 1.96, 4.9, 0.22, "使用体感", 9.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 7.10, 1.96, 4.4, 0.22, "能对上的证据", 9.2, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    hr(s, ML, 2.22, CW, 1.1, RULE2)
    rows = [("GPT", "写代码偏“防御式”，兜底写得很全；推理这一项最硬",
             "Astra 是 FrontierMath Erdős 未解题唯一得分的模型"),
            ("Claude", "最听得懂人话，一句话就办对；会主动查证，也爱纠正你",
             "报告案例 4 分钟出结果、金额全对、承建单位也查对"),
            ("Gemini", "知识面最广，中文最顺；但几乎不反对你，先夸再顺着说",
             "职业作家盲评榜前两名（2026-02，较旧）"),
            ("DeepSeek", "出字最快、最便宜；但先来一段“雷霆大思考”才给答案",
             "报告案例原因最具体，但有套错，耗时 34 分"),
            ("豆包", "彻底顺着你：说 A 对就夸 A 对，说 A 错就立刻改口",
             "未进本次盲测，也无同口径分数；只按日常使用记录")]
    yy = 2.36
    for i, (nm, feel, ev) in enumerate(rows):
        text(s, ML, yy, 1.46, 0.26, nm, 13, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML + 1.62, yy, 5.24, 0.56, feel, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.42)
        text(s, ML + 7.10, yy, 4.42, 0.56, ev, 10.5, GRN, SANS, False, PP_ALIGN.LEFT, 1.42)
        hr(s, ML, yy + 0.74, CW, 0.6, RULE)
        yy += 0.82
    text(s, ML, 6.44, CW, 0.28,
         "体感会因为客户端版本、提示词习惯而变；所以右边一列必须写“证据来自哪”。",
         9.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：分享人体感；本分享两个案例；Artificial Analysis；FrontierMath Erdős；LMArena")


# ---------------------------------------------------------------- 27 免费额度
def p27(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "先用免费积分跑通一个任务，再决定付不付",
             "价格与额度 2026-10-18 查询，活动经常变，以官网为准。")
    rows = [["客户端", "免费能用多少", "想多用：最低付费", "备注"],
            ["WorkBuddy（腾讯）", "每月 500 积分", "¥99/月（连续包月 ¥70）", "国内版，人民币支付"],
            ["Trae 国内版（字节）", "每月 500 积分，功能全开", "¥49/月起", "trae.cn"],
            ["Qoder（阿里）", "新用户 2 周试用 + 300 Credits", "$20/月；国内版 ¥59/月", "可接自己的 API Key"],
            ["Cline", "插件开源免费；限时免费模型轮换", "$9.99/月，或自带 Key 按量", "VS Code 插件"],
            ["Codex（OpenAI）", "ChatGPT 免费账号可用小模型", "$20/月（Plus）", "需海外账号"],
            ["Kiro（AWS）", "50 credits", "$20/月", "需海外账号"],
            ["Claude Code", "无免费版", "$20/月（Claude Pro）", "需海外账号"]]
    tbl_rows = []
    for i, r in enumerate(rows):
        if i == 0:
            tbl_rows.append(r)
        else:
            tbl_rows.append([r[0], r[1], ((r[2], GRN, 10.5, True),), r[3]])
    table(s, ML, 2.10, CW, tbl_rows, [3.05, 3.85, 2.85, 2.55], size=10.2, head_size=9.2,
          first_bold=True, aligns=[PP_ALIGN.LEFT] * 4, minrow=0.40, ls=1.2)
    text(s, ML, 6.06, CW, 0.24,
         "国内四个客户端都能用人民币直接付；海外四个需要海外账号与支付方式 —— 这一步的成本常被低估。",
         9.8, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：workbuddy.cn、trae.cn、qoder.com、docs.cline.bot、learn.chatgpt.com、kiro.dev")


def build(prs, start):
    pno = start
    div02(prs, pno)
    fn = [p19, p20, p21, p22, p23, p24, p25, p26, p27]
    for i, f in enumerate(fn):
        f(prs, pno + 1 + i)
