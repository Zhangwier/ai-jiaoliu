# -*- coding: utf-8 -*-
"""开场 P1-P5"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "开场"


def p01_cover(prs):
    s = slide(prs, DARK)
    # 底部整幅长卷（3:1 正好铺满，不裁切）
    pic(s, "s001_1.png", 0, 3.32, w=W, h=4.18, mode="fill")
    # 压一层极淡的暗角，让图与字衔接
    text(s, ML, 0.62, 8.0, 0.22, "工作常用 AI 的使用观察  ·  2026 年 10 月",
         11, C(0x8D,0x96,0xA4), SANS, False, PP_ALIGN.LEFT, 1.0, cs=2.2, check=False)
    text(s, ML - 0.04, 1.00, 9.6, 1.15, "AI 交流分享", 66, WHITE, SERIF, False,
         PP_ALIGN.LEFT, 1.05, check=False)
    text(s, ML, 2.24, 9.6, 0.42, "从原理到实践", 21, C(0xC8,0xCF,0xD9), SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 2.72, 9.6, 0.30,
         "三个真实测试：AI 能做到哪一步  ·  哪里会出错  ·  哪些地方必须人来核",
         11.5, C(0x93,0x9C,0xAB), SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    return s


def p02_quote(prs):
    s = slide(prs, DARK)
    header(s, CH, 2, dark=True)
    text(s, ML, 1.95, 11.2, 3.1,
         "交付物越像样，\n越要知道它会在哪里错。", 50, WHITE, SERIF, False,
         PP_ALIGN.LEFT, 1.42, check=False)
    text(s, ML, 5.20, 9.6, 0.80,
         "两个真实任务、十组配置、一百多条可核对的结论 ——\n把它们放在一起看，规律就出来了。",
         13, C(0xB2,0xBB,0xC7), SANS, False, PP_ALIGN.LEFT, 1.55, check=False)
    footer(s, "来源：本地分享材料", dark=True)
    return s


def p03_two_results(prs):
    s = slide(prs)
    y = head(s, CH, 3, "两组结果摆在一起，规律就出来了",
             "一个交付物是动画，一个是修编报告。看起来都能用，用起来都会出事。")
    top = 2.28
    colw = 5.35
    x1, x2 = ML, X2 - colw

    for i, (x, tag, big, unit, ttl, body, extra) in enumerate([
        (x1, "案例一  ·  图纸审查", "3", "/ 10", "十个配置审同一张工程图",
         "只有三个配置说出了图上真正存在的关键问题。同一批结果里，还有三十条各不相同的错误结论 —— 它们位置写得很具体，回原图一比，全都对不上。",
         "另外 21 条待核，未计入"),
        (x2, "案例二  ·  修编报告", "1", "/ 4", "四份报告的金额都算对了",
         "四份报告的总额都落在复算区间内，其中一份还多做了单价取值。这份报告金额之外有两处硬伤：总额数字写错一位，承建单位照抄了旧标段。",
         "101 个变化子目，核减 173.7369 万元"),
    ]):
        text(s, x, top, colw, 0.24, tag, 9.5, RED if i == 0 else GRN, SANS, True,
             PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        hr(s, x, top + 0.34, colw, 1.2, RED if i == 0 else GRN)
        text(s, x, top + 0.50, 2.0, 0.92, big, 62, INK, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + est_w(big, 62, True)/72.0 + 0.08, top + 0.78, 1.4, 0.30, unit,
             18, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 1.52, colw, 0.32, ttl, 16, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, top + 1.92, colw, 1.30, body, 12.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.62)
        text(s, x, top + 3.30, colw, 0.22, extra, 9.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)

    vr(s, MID, top + 0.30, 3.30, 0.75, RULE)
    footer(s, "来源：本地图纸审查记录（21 条待核未计）·  修编报告复算基准见表后页")
    return s


def p04_three_things(prs):
    s = slide(prs)
    y = head(s, CH, 4, "今天只讲三件事：怎么选、怎么交代、怎么检查",
             "三个测试各对应一件。只讲能立刻用上的判断，不铺开讲技术。")
    top = 2.34
    items = [
        ("01", "选配置", "模型、客户端、设置要一起看",
         "同一个模型换个外壳，结果能差一个档。先确认你用的是哪一档、有没有开视觉。", RED),
        ("02", "交代任务", "目标、材料、边界、交付物",
         "报告的差错多来自没交代清楚口径，不是模型不会写。", GRN),
        ("03", "检查结果", "回原文、重算、打开文件",
         "三样里最省时间的是打开文件看一眼 —— 假错误常常在这一步露出来。", BLU),
    ]
    gap = 0.55
    cw = (CW - gap*2) / 3
    for i, (no, ttl, sub, body, col) in enumerate(items):
        x = ML + i*(cw+gap)
        text(s, x, top, 1.2, 0.26, no, 11, col, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, x, top + 0.36, cw - 0.25, 1.6, col)
        text(s, x, top + 0.55, cw, 0.42, ttl, 25, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x, top + 1.10, cw, 0.26, sub, 11.5, col, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, top + 1.50, cw - 0.20, 1.05, body, 12, BODY, SANS, False, PP_ALIGN.LEFT, 1.62)
    hr(s, ML, 6.32, CW, 0.75, RULE)
    text(s, ML, 6.45, CW, 0.30,
         "顺序也重要：先用第一类活建立信任，再跑量，最后才把整段综合任务交出去。",
         12.5, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：本地分享材料")
    return s


def p05_toc(prs):
    s = slide(prs)
    y = head(s, CH, 5, "分享分四部分，外加一份附录",
             "前四部分按“认识 → 挑选 → 验证 → 上手”的顺序，附录是备查资料，可以不按顺序看。")
    top = 2.30
    rows = [
        ("01", "认识 AI", "五个短板，每一个都能从原理上讲清", "抽卡 · 幻觉 · 迎合 · 上下文腐烂 · 奖励黑客"),
        ("02", "模型与工具", "同一句提示词，十个配置给出十种结果", "鹈鹕盲测 · 榜单 · 各家脾气 · 价格"),
        ("03", "两个真实案例", "十组配置审一张图，四份报告写一件事", "图纸审查 · 修编报告原文对比"),
        ("04", "怎么用得稳", "三类活、一份任务书、五个办法", "提示词 · 上下文 · 检索 · 工作流 · 分级"),
    ]
    hh = 0.98
    for i, (no, ttl, l1, l2) in enumerate(rows):
        yy = top + i*hh
        text(s, ML, yy + 0.06, 0.75, 0.30, no, 13, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.80, yy, 2.05, 0.36, ttl, 17, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML + 2.95, yy + 0.03, 4.30, 0.34, l1, 12.5, INK, SANS, False, PP_ALIGN.LEFT, 1.35)
        text(s, ML + 7.35, yy + 0.05, 4.15, 0.30, l2, 10.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        hr(s, ML, yy + 0.80, CW, 0.60, RULE)

    text(s, ML, 6.28, 1.4, 0.24, "附录", 10, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    text(s, ML + 0.85, 6.26, CW - 0.85, 0.26,
         "A 常用术语  ·  B 公司档案  ·  C 订阅与省流  ·  D 资讯渠道  ·  E 分享与工具  ·  F 原理与历史  ·  Z 参考资料来源",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：本地分享材料")
    return s


def build(prs):
    p01_cover(prs); p02_quote(prs); p03_two_results(prs)
    p04_three_things(prs); p05_toc(prs)
