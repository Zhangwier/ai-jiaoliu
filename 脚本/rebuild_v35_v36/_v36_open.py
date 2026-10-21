# -*- coding: utf-8 -*-
"""开场 5 页"""
import os
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as I

ROOT = os.path.dirname(os.path.abspath(__file__))


def p01(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    pic(s, "s001_1.png", 4.00, 0, W - 4.00, H, mode="fill")
    box(s, 0.18, 0, 3.82, H, fill=INK)
    text(s, 0.92, 0.72, 3.40, 0.28, "工作常用 AI 的使用观察", 12, TEALL, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.4, check=False)
    text(s, 0.92, 1.00, 3.40, 0.26, "2026 年 10 月", 12, DARKGREY, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, 0.92, 1.66, 3.50, 1.62, "AI\n交流分享", 52, WHITE, SANS, True, PP_ALIGN.LEFT, 1.16,
         check=False)
    box(s, 0.92, 3.38, 0.86, 0.055, fill=TEAL)
    text(s, 0.92, 3.58, 3.20, 0.32, "从原理到实践", 19, TEALL, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    text(s, 0.92, 4.22, 3.20, 1.40,
         "三个真实测试：\nAI 能做到哪一步\n哪里会出错\n哪些地方必须人来核", 14, C(0xC5, 0xCE, 0xDD),
         SANS, False, PP_ALIGN.LEFT, 1.66, check=False)
    hr(s, 0.92, 6.44, 3.20, 1.0, C(0x2C, 0x3A, 0x55))
    text(s, 0.92, 6.58, 3.30, 0.24, "内部交流材料  ·  请勿外传", 10, GREY2, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    return s


def p02(prs, pno):
    """先导片"""
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, "开场", pno, dark=True)
    video = os.path.join(ROOT, "teaser.mp4")
    poster = os.path.join(ROOT, "poster.png")
    vw, vh = 8.30, 4.67
    vx, vy = W - vw - 0.62, (H - vh) / 2
    if os.path.exists(video):
        mv = s.shapes.add_movie(video, I(vx), I(vy), I(vw), I(vh),
                                poster_frame_image=poster, mime_type="video/mp4")
        mv.name = "先导片"
    text(s, ML, 1.52, 3.60, 0.26, "先导片  ·  1 分钟", 11, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.6, check=False)
    text(s, ML, 1.86, 3.80, 1.70, "先看\n一分钟", 40, WHITE, SANS, True, PP_ALIGN.LEFT, 1.24,
         check=False)
    text(s, ML, 3.70, 3.62, 1.30,
         "不介绍产品，也不讲概念。\n只把这半年里，我们用 AI 做的\n四类活、踩过的坑，摆在一起。", 13,
         C(0xC5, 0xCE, 0xDD), SANS, False, PP_ALIGN.LEFT, 1.62, check=False)
    hr(s, ML, 5.42, 3.50, 1.0, C(0x2C, 0x3A, 0x55))
    for i, t in enumerate(["识图", "写报告", "做表", "实战"]):
        text(s, ML + i * 0.92, 5.58, 0.90, 0.24, t, 11.5, DARKGREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             cs=0.8, check=False)
    text(s, ML, 6.16, 3.60, 0.44, "放映时点击画面即可播放。", 10, GREY2, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    return s


def p03(prs, pno):
    s = slide(prs)
    y = head(s, CH_OPEN, pno, "交付物越像样，越要知道它会在哪里错",
             "下面两个结果，来自这半年里两件真实的活。它们决定了后面九十分钟讲什么。")
    top = y + 0.26
    MIDX = W / 2
    # 左：案例一
    text(s, ML, top, 5.42, 0.26, "案例一  ·  挡块图纸审查", 11, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.2, check=False)
    text(s, ML, top + 0.34, 2.30, 1.00, "3", 62, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.06, top + 0.52, 1.40, 0.34, "/ 10", 20, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.06, top + 0.90, 4.30, 0.50, "十个配置里，只有三个找到关键问题", 13, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.3, check=False)
    text(s, ML, top + 1.56, 5.30, 0.80,
         "另有 30 条“错误”说得非常具体 —— 给了位置、给了数字、给了算式，\n回原图一核，对不上。", 12.5,
         GREY, SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
    # 右：案例二
    RX = MIDX + 0.42
    text(s, RX, top, 5.42, 0.26, "案例二  ·  修编报告", 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, RX, top + 0.34, 2.30, 1.00, "1", 62, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, RX + 0.70, top + 0.52, 1.40, 0.34, "/ 4", 20, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, RX + 0.70, top + 0.90, 4.40, 0.50, "四份报告的金额都对，一份藏着两处硬伤", 13, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.3, check=False)
    text(s, RX, top + 1.56, 5.40, 0.80,
         "总额写成 41638.89 万元（应为 41,638.9336）；\n承建单位照抄了旧标段的“中铁十二局”。", 12.5,
         GREY, SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
    vr(s, MIDX - 0.02, top, 2.40, 1.0, LINE2)
    takeaway(s, "两个结果指向同一件事：它出错的方式，和它做事的方式是同一套机制。", y=5.62,
             tag="所以")
    footer(s, "来源：本稿挡块图纸审查记录（2026-10）；TJ4 合同段修编报告四份输出")
    return s


def p04(prs, pno):
    s = slide(prs)
    y = head(s, CH_OPEN, pno, "今天只讲三件事：怎么选、怎么交代、怎么检查",
             "顺序不是随便排的。先选对，再交代清楚，最后才知道该查什么。")
    top = y + 0.30
    IT = [("01", "选配置", "模型、客户端、设置，三样一起看",
           "同一个牌子，网页免费版、会员版、API 常不是同一个模型。\n能不能直接看图，比分数更影响图纸类任务。",
           "鹈鹕动画：一句话，十个配置", ORANGE),
          ("02", "交代任务", "目标、材料、边界、交付物，四样写全",
           "不说范围、不说格式、不说怎么算完成 —— 这是最常见的一类问题。\n“没说”不是它笨，是你没给。",
           "修编报告：同一份清单，四份结果", TEAL),
          ("03", "检查结果", "回原文、重算、打开文件",
           "它是按概率写的，不是按事实核的。\n所以最后一步必须是人：翻回原图、自己按一遍计算器。",
           "图纸审查：30 条假错误", C(0x4A, 0x6C, 0x9E))]
    cw = (CW - 0.90) / 3
    for i, (no, nm, lead, body, ex, col) in enumerate(IT):
        x = ML + i * (cw + 0.45)
        box(s, x, top, cw, 0.055, fill=col)
        text(s, x, top + 0.22, 0.60, 0.28, no, 13, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        text(s, x, top + 0.54, cw, 0.44, nm, 26, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, x, top + 1.06, cw, 0.50, lead, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35,
             check=False)
        hr(s, x, top + 1.62, cw - 0.10, 1.0, LINE)
        text(s, x, top + 1.74, cw - 0.10, 1.10, body, 12, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.55, check=False)
        text(s, x, top + 2.92, cw - 0.10, 0.26, ex, 10.5, col, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
    box(s, ML, 6.02, CW, 0.66, fill=C(0xF3, 0xF1, 0xEB))
    text(s, ML + 0.24, 6.14, CW - 0.48, 0.44,
         "顺序也重要：先用第一件事建立判断力，再用第二件事放大产出，最后才敢把整段活交出去。", 13.5, NAVY,
         SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：本稿三个测试的执行记录")
    return s


def p05(prs, pno):
    s = slide(prs)
    y = head(s, CH_OPEN, pno, "分享分四部分，附录收细节与出处",
             "前面四部分讲判断，附录放可以直接查的东西 —— 术语、公司档案、订阅、渠道、来源。")
    top = y + 0.22
    IT = [("01", "认识 AI", "五个短板，以及它们各自的原理成因", "抽卡 · 幻觉 · 迎合 · 上下文腐烂 · 奖励黑客", ORANGE),
          ("02", "模型与工具", "同一句提示词，十个配置给出十种结果", "档次 · 专精 · 外壳 · 榜单 · 格局 · 额度", TEAL),
          ("03", "两个真实案例", "图纸审查与修编报告，逐条回原文核", "挡块图纸 · 修编报告 · 15.4 条文", C(0x4A, 0x6C, 0x9E)),
          ("04", "怎么用得稳", "三类活、任务书、上下文、资料库、工作流", "找路子 · 跑量 · 综合任务 · 数据分级", C(0x6B, 0x73, 0x85))]
    yy = top
    for i, (no, nm, lead, kw, col) in enumerate(IT):
        if i:
            hr(s, ML, yy - 0.06, CW, 1.0, LINE)
        text(s, ML, yy + 0.06, 0.66, 0.30, no, 12.5, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        text(s, ML + 0.72, yy, 2.90, 0.40, nm, 20, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, ML + 3.70, yy + 0.02, 4.60, 0.32, lead, 13, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML + 8.50, yy + 0.06, 3.62, 0.26, kw, 10.5, GREY2, SANS, False, PP_ALIGN.RIGHT, 1.0,
             check=False)
        yy += 0.92
    hr(s, ML, yy - 0.06, CW, 1.6, LINE2)
    text(s, ML, yy + 0.12, 1.10, 0.26, "附录", 12, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.4,
         check=False)
    AP = ["A 术语", "B 公司档案", "C 订阅与省流", "D 资讯渠道", "E 分享与工具", "F 原理与历史", "Z 参考来源"]
    for i, t in enumerate(AP):
        text(s, ML + 1.10 + i * 1.62, yy + 0.10, 1.58, 0.26, t, 11, GREY, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
    footer(s, "来源：本稿各部分原始记录；附录逐项标注出处")
    return s


CH_OPEN = "开场"
