# -*- coding: utf-8 -*-
"""第 04 章 怎么用得稳"""
import os
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "04  怎么用得稳"


def div(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.24, 1.5, 1.10, "04", 68, C(0x28, 0x3A, 0x59), SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.36, 1.30, 6.0, 0.30, "PART  04", 12, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=2.2, check=False)
    text(s, ML + 1.36, 1.58, 7.4, 0.66, "怎么用得稳", 38, WHITE, SANS, True, PP_ALIGN.LEFT, 1.14,
         check=False)
    text(s, ML + 1.36, 2.32, 8.4, 0.32, "用方法对付前面那些问题", 16, DARKGREY, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    hr(s, ML, 2.94, CW, 1.0, C(0x2C, 0x3A, 0x55))
    items = [("01", "三类活", "先分对，再决定检查强度：找路子 / 跑量 / 综合任务"),
             ("02", "五种方法", "提示词、上下文、RAG、工作流、输入方式，各对付一个短板"),
             ("03", "任务书与交付", "从一句话到四要素；四步留下可检查的结果"),
             ("04", "工作流与 Agent", "颗粒度、三个坑、多 Agent 的收益与代价"),
             ("05", "数据分级", "公开 / 内部 / 密钥，三级处理，付款删除外发由人确认")]
    yy = 3.16
    for no, nm, sub in items:
        text(s, ML, yy + 0.02, 0.62, 0.26, no, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.68, yy - 0.03, 2.50, 0.30, nm, 16, WHITE, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 3.30, yy, 8.10, 0.30, sub, 14, C(0xB6, 0xC0, 0xD0), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.68
    return s


def p40(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三类活，三种检查强度",
             "先看这件事属于哪一类：第一类最容易见效，第三类最需要人盯。")
    top = y + 0.20
    IT = [("01", "找路子", "你不知道有它", "找工具、公式、规范、先例", "看一眼出处",
           "百余个文件里定位一份图纸", 1, C(0x4A, 0x6C, 0x9E)),
          ("02", "跑量", "量大、规则明确", "批量转表、改名、抽字段、套模板", "先抽 3 个验格式",
           "三百份合同抽甲方和金额", 3, TEAL),
          ("03", "综合任务", "要判断、要取舍", "修编报告、审图、合同改写、长任务", "拆步留痕，人签字",
           "看图、看表、看说明再下结论", 5, ORANGE)]
    cw = (CW - 0.86) / 3
    for i, (no, nm, tag, body, how, ex, lvl, col) in enumerate(IT):
        x = ML + i * (cw + 0.43)
        text(s, x, top, 0.62, 0.26, no, 12, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        text(s, x, top + 0.28, cw, 0.42, nm, 24, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, x, top + 0.74, cw, 0.26, tag, 12, col, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        hr(s, x, top + 1.08, cw - 0.12, 1.0, LINE)
        text(s, x, top + 1.18, cw - 0.12, 0.56, body, 12.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.45, check=False)
        text(s, x, top + 1.82, cw - 0.12, 0.24, "怎么核", 10, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=1.2, check=False)
        text(s, x, top + 2.06, cw - 0.12, 0.26, how, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, x, top + 2.40, cw - 0.12, 0.24, ex, 10.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        # 检查强度计量
        text(s, x, top + 2.78, 1.30, 0.24, "检查强度", 10, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             cs=1.0, check=False)
        for k in range(5):
            box(s, x + 0.92 + k * 0.30, top + 2.83, 0.24, 0.13,
                fill=col if k < lvl else C(0xE2, 0xDE, 0xD3))
    # 上手顺序
    hr(s, ML, 5.62, CW, 1.0, LINE)
    text(s, ML, 5.76, 1.16, 0.26, "上手顺序", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    ST = ["先用第一类建立信任", "再交给它跑量", "最后才整段交出综合任务"]
    for i, t in enumerate(ST):
        x = ML + 1.24 + i * 3.66
        text(s, x, 5.72, 0.34, 0.24, f"{i+1}", 12, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.30, 5.72, 2.90, 0.28, t, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        if i < 2:
            text(s, x + 3.16, 5.70, 0.34, 0.24, "›", 14, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
    takeaway(s, "先按类分，再按类给不同的检查强度：第三类永远要人签字。", y=6.14, tag="结论")
    footer(s, "来源：本稿四类真实任务（识图、报告、做表、实战）的复盘记录")
    return s


def p41(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "第一类实例：百余个文件中定位一张图纸",
             "它的检索能力，要一张自己能核对的表来托底 —— 不是它比你聪明，是它肯把一百多个文件翻一遍。")
    top = y + 0.16
    TW = 7.34
    rows = [["序号", "文件名", "类型", "出图编号", "所在册"],
            ["3", "控制成果表.xlsx", "表格", "S1-2-2", "第二册 路线"],
            ["6", "逐桩坐标表.xlsx", "坐标表", "S1-2-7", "第二册 路线"],
            ["8", "砍树挖根数量表.xls", "数量表", "S1-2-10", "第二册 路线"],
            ["…", "……", "……", "……", "……"]]
    table(s, ML, top, TW, rows, [0.66, 2.60, 1.24, 1.52, 1.86], size=12, head_size=12,
          minrow=0.42, head_h=0.40, rowh=0.42, pad=0.11,
          colors=lambda r, c: FAINT if r == 4 else (TEAL if c == 3 else C(0x2B, 0x34, 0x45)))
    text(s, ML, top + 2.30, TW, 0.24, "实际表共 9 列，另含子目录、大小、修改时间，和可点击的相对路径",
         10, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    # 册汇总
    box(s, ML, 4.28, TW, 1.42, fill=C(0xF1, 0xEF, 0xE8))
    text(s, ML + 0.24, 4.40, 2.4, 0.24, "另有一张「册汇总」", 10.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=0.8, check=False)
    text(s, ML + 0.24, 4.68, TW - 0.48, 0.30,
         "第二册 路线 · 26 个文件 · 9 个数量表", 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
         check=False)
    text(s, ML + 0.24, 5.02, TW - 0.48, 0.30,
         "第三册一分册 路基 · 18 个文件 · 11 个数量表", 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
         check=False)
    text(s, ML + 0.24, 5.36, TW - 0.48, 0.24, "先看它决定翻哪一册，再动手", 10.5, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    # 右栏三件事
    RX = ML + TW + 0.42
    RW = X2 - RX
    vr(s, RX - 0.24, top, 4.60, 1.0, LINE)
    text(s, RX, top - 0.02, RW, 0.26, "学会用之后，找图变成三件事", 13, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.1, check=False)
    TT = [("按册筛", "“第三册第一分册有哪些数量表”—— 筛一列就出来"),
          ("按出图编号定位", "图纸目录上写 S1-2-7，直接查这一列，落到具体文件"),
          ("按类型找同类", "要所有“数量表”，筛类型；要所有坐标表，同理")]
    yy = top + 0.40
    for i, (nm, body) in enumerate(TT):
        text(s, RX, yy, RX - RX + 0.26, 0.24, f"0{i+1}", 11, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.34, yy - 0.02, RW - 0.34, 0.28, nm, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT,
             1.1, check=False)
        nl = nlines(body, 11, RW - 0.34)
        text(s, RX + 0.34, yy + 0.28, RW - 0.34, nl * 11 * 1.32 * 1.4 / 72 + 0.05, body, 11, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
        yy += 0.92
    hr(s, RX, yy + 0.02, RW, 1.0, LINE)
    text(s, RX, yy + 0.16, RW, 0.26, "表里没有的，直接问它", 12.5, NAVY, SANS, True, PP_ALIGN.LEFT,
         1.1, check=False)
    text(s, RX, yy + 0.46, RW, 0.72,
         "“有没有 XX 桥台的挡块图” —— 它会去翻目录和分册表，再告诉你在哪本。", 11.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45, check=False)
    takeaway(s, "这就是最土的 RAG：先建索引，再检索。可核对、可重生成 —— 换项目只改一个路径参数。", y=6.06,
             tag="本质")
    footer(s, "来源：本地文件索引脚本与生成结果（2026-10）")
    return s


def p42(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "五种方法，对应五个短板",
             "没有一种方法能包治百病，按问题挑方法。")
    top = y + 0.18
    IT = [("提示词工程", "迎合、瞎猜", "写成任务书：范围、格式、验收，允许“材料不足”",
           "图纸：范围、证据与验收要求没写清", ORANGE),
          ("上下文管理", "上下文腐烂", "一个任务一段对话，长了压缩或新开，只放相关材料",
           "长对话后忘了前面的要求", TEAL),
          ("RAG 检索", "幻觉、知识过时", "先从资料库找片段，再带出处回答",
           "报告：原因要能回到函件和勘误表", C(0x4A, 0x6C, 0x9E)),
          ("工作流与校验", "抽卡、计算错", "固定步骤给程序，数字用脚本重算，多跑几次比",
           "报告：41,638.89 的笔误", C(0x6B, 0x73, 0x85)),
          ("选对输入方式", "看不见就乱猜", "图纸先确认是“看图”还是 OCR",
           "图纸：只靠 OCR 的一个没找到", C(0x8A, 0x6A, 0x3A))]
    # 表头
    XM, XT, XW, XE = ML, ML + 2.28, ML + 4.32, ML + 8.62
    text(s, XM, top, 2.0, 0.24, "方法", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, XT, top, 1.9, 0.24, "对付", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, XW, top, 4.2, 0.24, "怎么做", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, XE, top, 4.0, 0.24, "第 03 部分的例子", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    hr(s, ML, top + 0.28, CW, 1.0, LINE2)
    yy = top + 0.40
    for i, (nm, tgt, how, ex, col) in enumerate(IT):
        if i:
            hr(s, ML, yy - 0.14, CW, 1.0, LINE)
        box(s, ML, yy, 0.05, 0.46, fill=col)
        text(s, ML + 0.20, yy + 0.02, 2.00, 0.26, nm, 14.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, XT, yy + 0.04, 1.86, 0.26, tgt, 12, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, XW, yy + 0.02, 4.16, 0.50, how, 12, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT,
             1.35, check=False)
        text(s, XE, yy + 0.04, 3.90, 0.26, ex, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.82
    takeaway(s, "五种方法不是五套配置，是五个“该动哪一处”的判断。", y=6.16, tag="要点")
    footer(s, "来源：本分享第 01、03 部分；方法说明")
    return s


def p43(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "提示词：从一句话到一份任务书",
             "“有没有问题”本身是开放问法，问题不在问法，在范围、证据与验收都没写清。")
    top = y + 0.18
    # 原话
    box(s, ML, top, 5.10, 0.92, fill=ORANGEL)
    text(s, ML + 0.22, top + 0.12, 1.60, 0.24, "案例里的原话", 10, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.0, check=False)
    text(s, ML + 0.22, top + 0.40, 4.70, 0.34, "“审查这个有没有问题，结果输出 excel 给我”", 12.5,
         C(0x7A, 0x3A, 0x22), SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    # 四要素
    E = [("01", "任务目标", "检查图纸内部是否一致，没有问题也可以"),
         ("02", "材料范围", "先说明读到哪些页、是看图还是 OCR、哪些看不清"),
         ("03", "输出格式", "可编辑 Excel，计算项保留公式"),
         ("04", "验收要求", "每项写位置、依据和把握程度；不补造尺寸")]
    yy = top + 1.14
    for no, nm, body in E:
        text(s, ML, yy + 0.04, 0.52, 0.26, no, 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.58, yy, 1.72, 0.28, nm, 14, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 2.40, yy + 0.02, 3.92, 0.60, body, 12, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.4, check=False)
        hr(s, ML, yy + 0.70, 5.10, 1.0, LINE)
        yy += 0.80
    # 右：完整写法
    RX = ML + 5.58
    RW = X2 - RX
    vr(s, RX - 0.28, top, 3.90, 1.0, LINE)
    text(s, RX, top - 0.02, RW, 0.26, "完整写法（改写示例，未做对照复测）", 11, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    box(s, RX, top + 0.30, RW, 3.36, fill=CARD, line=LINE)
    BODY = ("请审查这份挡块图纸内部是否一致，没有问题也可以。\n"
            "先说明你读到了哪些页、是直接看图还是只做了 OCR、哪些地方看不清。\n"
            "逐项列出：位置（页/图号）、原文依据、判断、把握程度（高中低）。\n"
            "不要补造尺寸；不确定的写进“待确认”，并写明缺什么证据。\n"
            "输出为可编辑 Excel，计算项保留公式。")
    text(s, RX + 0.26, top + 0.50, RW - 0.52, 2.60, BODY, 12, C(0x2B, 0x34, 0x45), SANS, False,
         PP_ALIGN.LEFT, 1.62, check=False)
    text(s, RX, top + 3.78, RW, 0.52,
         "四要素齐了，检查才有落点：范围可核、格式可开、验收可判。", 11.5, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.4, check=False)
    takeaway(s, "主要可改进点不在于“问得更客气”，而在于把范围、证据与验收写进去。", y=6.16, tag="提示")
    footer(s, "来源：原任务提示词，改写为方法示例，未复测")
    return s


def p44(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "交付四步：每一步留下可检查的结果",
             "上一步有看得见的结果，下一步才有可靠的输入。")
    top = y + 0.24
    ST = [("1", "读材料", "列出读了哪些文件、范围、缺页和认不出的地方", "清单 ＋ 函件 ＋ 勘误表", TEAL),
          ("2", "定口径", "说清单位、金额范围、资料优先级、哪些不许猜", "汇总表口径，不含暂列金额",
           C(0x4A, 0x6C, 0x9E)),
          ("3", "做产物", "按模板交可编辑文件，保留计算和出处", "报告 ＋ 变更表", C(0x6B, 0x73, 0x85)),
          ("4", "查结果", "抽查事实、自己重算、打开文件看版面", "金额复算＋原因定位＋模板残留", ORANGE)]
    n = 4
    gap = 0.30
    w = (CW - gap * (n - 1)) / n
    for i, (no, nm, body, ex, col) in enumerate(ST):
        x = ML + i * (w + gap)
        box(s, x, top, w, 2.62, fill=CARD, line=LINE)
        box(s, x, top, w, 0.055, fill=col)
        text(s, x + 0.24, top + 0.20, 0.60, 0.30, no, 13, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.24, top + 0.56, w - 0.48, 0.34, nm, 19, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        nl = nlines(body, 12, w - 0.48)
        text(s, x + 0.24, top + 1.00, w - 0.48, nl * 12 * 1.32 * 1.5 / 72 + 0.05, body, 12,
             C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
        hr(s, x + 0.24, top + 2.00, w - 0.48, 1.0, LINE)
        text(s, x + 0.24, top + 2.08, w - 0.48, 0.24, "修编报告里对应", 9.5, GREY2, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.24, top + 2.32, w - 0.48, 0.26, ex, 11, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        if i < n - 1:
            line(s, x + w + 0.04, top + 1.30, x + w + gap - 0.04, top + 1.30, LINE2, 1.4, head=True)
    hr(s, ML, 5.14, CW, 1.0, LINE)
    text(s, ML, 5.28, 1.30, 0.26, "断链在哪", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    BR = [("④ 查结果被跳过", "文件打不开、版面错位，直到现场才发现"),
          ("③ 做产物不留痕", "数字改了但没留公式，复算无从下手"),
          ("② 定口径没说", "四份报告各用各的口径，金额对不上")]
    for i, (t, b) in enumerate(BR):
        x = ML + 1.34 + i * 3.60
        text(s, x, 5.24, 3.46, 0.26, t, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, 5.54, 3.46, 0.28, b, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    takeaway(s, "四步里最容易被省掉的是第 4 步，而它恰好是唯一能拦住错误的一步。", y=6.16, tag="注意")
    footer(s, "来源：本地修编报告案例")
    return s


def p45(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "上下文管理：控制对话长度与输入量",
             "窗口 1M 不等于要塞满 1M：放得越精，答得越准。")
    top = y + 0.18
    L = [("01", "一个任务一段对话", "换任务就新开，旧材料不带过去"),
         ("02", "长了就压缩", "用客户端的“压缩上下文”，或让它写进度摘要后新开"),
         ("03", "只放相关的", "先让它说需要哪几页、哪张表，再给"),
         ("04", "规则写进文件", "单位口径、格式、禁止事项写成规则文件，每次自动带上"),
         ("05", "要求放开头", "关键要求放最前，必要时结尾再说一遍")]
    yy = top
    for no, nm, body in L:
        text(s, ML, yy + 0.02, 0.52, 0.26, no, 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.58, yy - 0.02, 2.30, 0.28, nm, 14.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 3.00, yy + 0.02, 3.20, 0.52, body, 11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35,
             check=False)
        hr(s, ML, yy + 0.62, 6.20, 1.0, LINE)
        yy += 0.74
    # 右：规则文件
    RX = ML + 6.60
    RW = X2 - RX
    text(s, RX, top - 0.02, RW, 0.26, "规则文件长什么样（示例）", 11, TEALL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    box(s, RX, top + 0.30, RW, 2.92, fill=INK)
    box(s, RX, top + 0.30, 0.05, 2.92, fill=TEAL)
    CODE = ["# 项目规则",
            "- 金额单位：万元，保留 4 位小数",
            "- 汇总口径：不含暂列金额",
            "- 合同条文原样引用，不改写",
            "- 原因必须注明函件编号或勘误表序号",
            "- 找不到依据就写“待核”，不要补"]
    cy = top + 0.50
    for i, ln in enumerate(CODE):
        text(s, RX + 0.28, cy, RW - 0.52, 0.30, ln, 11.5,
             TEALL if i == 0 else C(0xC9, 0xD3, 0xE2), MONO if i == 0 else SANS, i == 0,
             PP_ALIGN.LEFT, 1.0, check=False)
        cy += 0.38
    text(s, RX, top + 3.36, RW, 0.52, "写成文件的好处：每次自动带上，不靠你记。", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.4, check=False)
    takeaway(s, "对付上下文腐烂：控制放进去的东西，比换更长窗口的模型更有效。", y=6.16, tag="结论")
    footer(s, "来源：Chroma《Context Rot》；Claude Code、Codex 官方文档")
    return s


def p46(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "检索增强（RAG）：资料多时先检索、再回答",
             "买的是“少付 token”和“少编”。做法是“先找，再答”，检索可用关键词、向量或两者混合。")
    top = y + 0.16
    # 左：token 对比
    text(s, ML, top, 5.20, 0.24, "① 省一个量级", 11, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    text(s, ML, top + 0.28, 5.20, 0.26, "100 页整包 ≈ 5–8 万 token，问一次付一次", 12,
         C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    box(s, ML, top + 0.62, 5.00, 0.26, fill=ORANGE)
    text(s, ML + 0.16, top + 0.66, 4.70, 0.24, "整包灌进去：5–8 万 token", 10, WHITE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, top + 0.98, 5.20, 0.26, "只检索命中 3–5 段 ≈ 1–2 千 token", 12, C(0x2B, 0x34, 0x45),
         SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    box(s, ML, top + 1.32, 0.44, 0.26, fill=TEAL)
    text(s, ML + 0.52, top + 1.35, 4.10, 0.24, "差几十倍；每追问一次，这个差就再付一遍", 10, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    # ② 少编
    text(s, ML, top + 1.82, 5.20, 0.24, "② 少编：把“凭印象”换成“照着原文”", 11, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML, top + 2.12, 5.10, 0.72,
         "材料不在眼前：只能按“最像”续写，数字、条文最容易编。\n原文摆到眼前：有东西可抄，还能标出处。",
         11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
    text(s, ML, top + 2.92, 5.10, 0.30,
         "再加一句“只依据片段回答，找不到就说没找到”。", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    # 右：embedding 图
    RX = ML + 5.60
    RW = X2 - RX
    text(s, RX, top - 0.02, RW, 0.24, "语义能落到：两段话意思近，向量就靠得近", 10.5, GREY2, SANS,
         False, PP_ALIGN.LEFT, 1.0, check=False)
    pic(s, "s114_1.png", RX, top + 0.28, RW, 1.86, mode="fit")
    text(s, RX, top + 2.20, RW, 0.24, "但“近”≠“对”：近邻里常有看着像、其实不是一回事的段落", 10,
         ORANGE, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    # 流程
    hr(s, ML, 5.62, CW, 1.0, LINE)
    text(s, ML, 5.72, 1.24, 0.24, "五环节", 10, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    FL = ["入库切分建索引（一次）", "关键词 / 向量 / 混合检索取候选", "精排筛证据", "拼进提示词",
          "回答并给出处"]
    for i, t in enumerate(FL):
        x = ML + 1.20 + i * 2.24
        box(s, x, 5.70, 1.96, 0.34, fill=TEALXL if i < 4 else MINT)
        text(s, x + 0.12, 5.76, 1.76, 0.24, t, 10.5, TEAL if i < 4 else C(0x0B, 0x54, 0x4E), SANS,
             True, PP_ALIGN.LEFT, 1.0, check=False)
        if i < 4:
            text(s, x + 2.00, 5.73, 0.24, 0.24, "›", 13, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
    takeaway(s, "关键词空手的例子：“防撞设施”与“挡块”一个字都不重合 —— 所以实践多用混合检索。",
             y=6.20, tag="做法")
    footer(s, "来源：Chroma《Context Rot》；Lewis 等 RAG 原论文（2020）；各厂 embedding 与计费文档")
    return s


def p47(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "工作流：把系统性任务拆成有标准的步骤",
             "不是“换更强的模型”，是“把每一步的标准写下来”。审图与修编报告，骨架其实是同一个。")
    top = y + 0.14
    X1, X2c, X3, X4 = ML, ML + 1.86, ML + 2.62, ML + 8.26
    text(s, X1, top, 1.80, 0.24, "步骤", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, X2c, top, 0.72, 0.24, "谁做", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, X3, top, 5.50, 0.24, "按什么判（标准）", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, X4, top, 4.40, 0.24, "交出什么", 10.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    hr(s, ML, top + 0.26, CW, 1.0, LINE2)
    ST = [("① 定范围", "人", "页数、图号、哪些表；先确认是“看图”还是“只 OCR”", "一页清单"),
          ("② 提取", "OCR＋视觉", "页数、表名、行数对得上；缺页、糊字要点名", "结构化底表"),
          ("③ 逐项核对", "模型", "每条结论要有图中位置＋原文；不确定项进入“待确认”栏", "问题清单"),
          ("④ 复算", "脚本", "关键量能独立复算；算不出来不进正文", "计算底稿"),
          ("⑤ 汇总结论", "人", "分“不一致 / 待确认 / 未检查”，不许混写", "审查意见")]
    yy = top + 0.38
    for i, (nm, who, std, out) in enumerate(ST):
        if i:
            hr(s, ML, yy - 0.12, CW, 1.0, LINE)
        text(s, X1, yy + 0.02, 1.78, 0.28, nm, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, X2c, yy + 0.04, 0.86, 0.24, who, 10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, X3, yy + 0.02, 5.44, 0.48, std, 11, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT,
             1.35, check=False)
        text(s, X4, yy + 0.04, 4.30, 0.24, out, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.66
    # 两个案例的坑
    hr(s, ML, 5.30, CW, 1.0, LINE)
    text(s, ML, 5.42, 1.40, 0.24, "两个案例的坑", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    PT = [("把 mm/cm 当错误", "没先定“看图例”"),
          ("汇总与分项差 112 元", "没定口径"),
          ("核减数算不出", "没复算")]
    for i, (t, b) in enumerate(PT):
        x = ML + 1.46 + i * 3.56
        text(s, x, 5.40, 3.40, 0.26, t, 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, 5.68, 3.40, 0.26, "→ " + b, 11, ORANGE, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    # 右：修编报告同构
    text(s, ML, 6.02, CW, 0.28,
         "修编报告同构：读清单 → 定位变更 → 套模板 → 复算金额 → 写原因 —— 也各标“谁做、按什么判、交出什么”。",
         12, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 6.34, CW, 0.26, "工作流是长出来的：每跑一次，把踩到的坑写成一条标准塞回对应步骤。", 11, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：本稿图纸审查与修编报告两个案例的实际执行与复核记录")
    return s


def p48(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "工作流的颗粒度：哪些写死，哪些留白",
             "工作流首先要让人能验收，而不是先让模型听话。")
    top = y + 0.16
    # 该写死
    LW = 5.90
    text(s, ML, top, LW, 0.28, "该写死的：三条全中", 13.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.1,
         check=False)
    hr(s, ML, top + 0.34, LW, 1.6, TEAL)
    FIX = [("格式固定", "Excel 行列、模板单元格、报告章节都是死的"),
           ("有唯一正确答案", "合计数、差额、核减额，算错就是错"),
           ("出错代价高", "要能复算、能追责，不能“大概对”")]
    yy = top + 0.48
    for nm, body in FIX:
        text(s, ML, yy, 0.20, 0.24, "·", 13, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.26, yy - 0.02, 1.94, 0.26, nm, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 2.26, yy, 3.56, 0.28, body, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.46
    # 该留白
    RX = ML + LW + 0.56
    RW = X2 - RX
    text(s, RX, top, RW, 0.28, "该留白的：中一条就交给模型", 13.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.1, check=False)
    hr(s, RX, top + 0.34, RW, 1.6, ORANGE)
    FREE = [("材料不全", "缺页、口径不清，要临场判断"),
            ("输出是自然语言", "写原因、写信、写说明"),
            ("步骤说不清", "要边做边看结果再决定下一步")]
    yy = top + 0.48
    for nm, body in FREE:
        text(s, RX, yy, 0.20, 0.24, "·", 13, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX + 0.26, yy - 0.02, 1.94, 0.26, nm, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 2.26, yy, RW - 2.26, 0.28, body, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.46
    # 过细的三害
    box(s, ML, 3.86, LW, 1.62, fill=ORANGEL)
    text(s, ML + 0.24, 3.96, LW - 0.48, 0.24, "流程写得过细，三个害处", 10.5, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    HRM = [("拆分反而掉分", "模型本来一次能做对，拆成五步每步掉一点"),
           ("标准写成了操作", "“按这个顺序、用这个模板”——它发现更快的路也不许走"),
           ("给旧模型打的补丁", "新模型一升级，绕坑的规则全成了负担")]
    yy = 4.24
    for nm, body in HRM:
        text(s, ML + 0.24, yy, 1.90, 0.24, nm, 11, C(0x7A, 0x3A, 0x22), SANS, True, PP_ALIGN.LEFT,
             1.0, check=False)
        text(s, ML + 2.22, yy, 3.44, 0.24, body, 10.5, C(0x7A, 0x3A, 0x22), SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
        yy += 0.40
    # 留活口三条
    text(s, RX, 3.90, RW, 0.26, "怎么留活口：三条", 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    LIVE = ["标准写“要达到什么”，不写“怎么做”", "每一步内部留一个判断位，允许它换做法",
            "定期复核规则：确认已过时再删"]
    yy = 4.24
    for t in LIVE:
        text(s, RX, yy, 0.24, 0.24, "–", 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX + 0.26, yy, RW - 0.26, 0.26, t, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.36
    takeaway(s, "两档不是对立的：外层（交付、审批、复算）写死，内层（怎么找、怎么写）留给模型。",
             y=6.06, tag="一句话")
    footer(s, "来源：本稿两案例的执行记录；各模型客户端说明；模型能力变化的观察（2026-10）")
    return s


def p49(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "多 Agent 协作：收益与代价",
             "人多不自动变可靠：拆得开、验得了、有人兜底，才值得拆。")
    top = y + 0.16
    # 左：收益
    LW = 5.70
    text(s, ML, top, LW, 0.26, "什么时候值：两条真实收益", 12.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.0, check=False)
    hr(s, ML, top + 0.30, LW, 1.6, TEAL)
    text(s, ML, top + 0.44, 0.60, 0.24, "并行", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 0.66, top + 0.42, LW - 0.66, 0.46,
         "三个 Agent 各查一份清单的不同章节：串行 40 分钟的活，十几分钟出初稿。", 11, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45, check=False)
    text(s, ML, top + 1.00, 0.60, 0.24, "上下文干净", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 0.66, top + 0.98, LW - 0.66, 0.62,
         "每个只看自己那一摊，主线不被几十页材料冲淡 —— 接着“上下文腐烂”那一条。", 11, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.45, check=False)
    # 左：代价
    text(s, ML, top + 1.74, LW, 0.26, "什么时候只是更贵：三条真实代价", 12.5, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, top + 2.04, LW, 1.6, ORANGE)
    COST = [("成本翻倍", "三个 Agent = 三份 token，还得加上协调的往返"),
            ("写冲突", "两个改同一个文件，后写的把前面的覆盖掉"),
            ("错误前提传染", "三个都用了错的汇总口径 → 一起错，还像“交叉验证过”")]
    yy = top + 2.16
    for nm, body in COST:
        text(s, ML, yy, 1.44, 0.24, nm, 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 1.50, yy, LW - 1.50, 0.26, body, 10.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.48
    # 右：四条门槛
    RX = ML + LW + 0.60
    RW = X2 - RX
    vr(s, RX - 0.30, top, 3.60, 1.0, LINE)
    text(s, RX, top - 0.02, RW, 0.26, "拆之前先过这四条（缺一就别拆）", 12.5, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    GATE = [("任务能切成互不依赖的块", ""), ("每块有独立、可核对的产出", ""),
            ("写文件的范围不重叠", ""), ("有人负责合并并验收", "")]
    yy = top + 0.36
    for i, (t, _) in enumerate(GATE):
        box(s, RX, yy, 0.26, 0.26, fill=TEAL if i < 4 else LINE)
        text(s, RX + 0.06, yy + 0.03, 0.20, 0.20, f"{i+1}", 9, WHITE, SANS, True, PP_ALIGN.LEFT,
             1.0, check=False)
        text(s, RX + 0.40, yy + 0.02, RW - 0.40, 0.26, t, 12, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
        yy += 0.42
    text(s, RX, yy + 0.10, RW, 0.26, "这几种就别拆", 11, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.8, check=False)
    text(s, RX, yy + 0.38, RW, 0.52,
         "步骤前后强依赖（第二步要用第一步的结论）· 总量本来就小 · 没有验收人", 10.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.4, check=False)
    # 假复核
    box(s, RX, 5.06, RW, 0.86, fill=C(0xF1, 0xEF, 0xE8))
    text(s, RX + 0.20, 5.14, RW - 0.40, 0.24, "最常见的一种假复核", 10, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    text(s, RX + 0.20, 5.40, RW - 0.40, 0.44,
         "复核的人只复述执行者的答案，等于没复核；必须回到原材料逐条对。", 10.5, C(0x2B, 0x34, 0x45),
         SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    takeaway(s, "收益来自并行和上下文隔离；代价来自协调、文件冲突，以及“同一个错误前提”。", y=6.16,
             tag="结论")
    footer(s, "来源：本稿多 Agent 协作记录；Subagent 与并行任务说明；各客户端文档")
    return s


def p50(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "工作流与 Agent 编排：三个常见陷阱",
             "前面讲怎么搭；这一页讲它什么时候会坏事 —— 三条都是本稿踩过的。")
    top = y + 0.18
    PIT = [("坑一", "低级模型＋复杂工作流，不如高级模型一句话",
            "工作流不增加理解力，只会把同一个误解重复五遍。本次差的几组不是步骤不够细，是没看懂图。",
            "先确认模型够用，再谈工作流", ORANGE),
           ("坑二", "不逐步确认，跑完查不出哪一步坏了",
            "中间产物没留，结果不对只能整条重跑，时间全耗在重跑上。",
            "每一步留一个能查的产物", C(0x4A, 0x6C, 0x9E)),
           ("坑三", "一个 AI 没看懂，换一个 AI 复核也看不懂",
            "复核者拿到的还是那段模糊描述，只会给出另一个模糊判断。",
            "复核要能回到原文逐条对", TEAL)]
    yy = top
    for i, (tag, ttl, body, fix, col) in enumerate(PIT):
        if i:
            hr(s, ML, yy - 0.16, CW, 1.0, LINE)
        text(s, ML, yy + 0.02, 1.10, 0.26, tag, 10.5, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.4,
             check=False)
        text(s, ML + 1.16, yy - 0.04, 6.60, 0.34, ttl, 17, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, ML + 1.16, yy + 0.40, 6.60, 0.52, body, 11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.45,
             check=False)
        text(s, ML + 8.36, yy + 0.04, 1.00, 0.24, "怎么避", 9.5, GREY2, SANS, True, PP_ALIGN.LEFT,
             1.0, cs=1.0, check=False)
        text(s, ML + 8.36, yy + 0.30, 3.72, 0.44, fix, 12, col, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        yy += 1.16
    box(s, ML, 5.62, CW, 0.70, fill=TEALXL)
    text(s, ML + 0.26, 5.74, CW - 0.52, 0.44,
         "三句话并起来其实是一句：先确认它看懂了，再谈拆几步、用几个 Agent。", 14, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：本稿两案例的执行与返工记录；多 Agent 协作记录")
    return s


def p51(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "数据分三级处理，付款、删除、外发前由人确认",
             "越敏感的数据、越不能撤回的动作，越要人在最后一步把关。")
    top = y + 0.16
    TIER = [("公开资料", "按允许的流程使用", MINT, TEAL),
            ("内部业务材料", "脱敏，用单位允许的工具和账号", C(0xF1, 0xEF, 0xE8), C(0x6B, 0x73, 0x85)),
            ("密钥与高敏感资料", "不贴进普通对话，用专门的凭据机制", ORANGEL, ORANGE)]
    cw = (CW - 0.66) / 3
    for i, (nm, body, bg, col) in enumerate(TIER):
        x = ML + i * (cw + 0.33)
        box(s, x, top, cw, 1.26, fill=bg)
        box(s, x, top, cw, 0.05, fill=col)
        text(s, x + 0.24, top + 0.20, cw - 0.48, 0.30, nm, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, x + 0.24, top + 0.60, cw - 0.48, 0.52, body, 11.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.4, check=False)
    # 操作权限
    text(s, ML, 3.56, 6.10, 0.26, "操作权限", 12, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
         check=False)
    hr(s, ML, 3.86, 6.06, 1.6, ORANGE)
    OP = [("读写本地文件、执行命令", "只给完成任务必需的权限"),
          ("对外发送、付款、删除", "最后一步由人核对对象、内容、金额和后果")]
    yy = 4.00
    for t, b in OP:
        text(s, ML, yy, 6.06, 0.26, t, 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML, yy + 0.28, 6.06, 0.26, b, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.72
    # 误传处置
    RX = ML + 6.46
    RW = X2 - RX
    text(s, RX, 3.56, RW, 0.26, "已经误传了？四步处置", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    hr(s, RX, 3.86, RW, 1.6, TEAL)
    D = ["马上停止扩散", "撤销或轮换相关凭据", "保留必要记录", "按单位流程处理"]
    yy = 4.00
    for i, t in enumerate(D):
        text(s, RX, yy + 0.02, 0.34, 0.24, f"{i+1}", 11, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.34, yy, RW - 0.34, 0.26, t, 12, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
        yy += 0.44
    takeaway(s, "三级的分界不在“数据多重要”，而在“一旦发出去能不能收回”。", y=6.16, tag="判断")
    footer(s, "来源：工作资料与权限管理方法；单位内部资料管理要求")
    return s


def p52(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, CH, pno, dark=True)
    text(s, ML, 0.86, 8.0, 0.32, "带走三件事", 13, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=2.0,
         check=False)
    text(s, ML, 1.24, 10.6, 0.66, "选好配置，交代清楚，检查结果", 40, WHITE, SANS, True,
         PP_ALIGN.LEFT, 1.15, check=False)
    hr(s, ML, 2.10, CW, 1.0, C(0x2C, 0x3A, 0x55))
    IT = [("01", "选配置", "确认型号和入口，图纸类任务先确认能看图"),
          ("02", "交代清楚", "目标、材料、格式、怎么算完成，不替它下结论"),
          ("03", "检查结果", "回原文、重算关键数字、打开实际文件")]
    yy = 2.34
    for no, nm, body in IT:
        text(s, ML, yy + 0.04, 0.62, 0.28, no, 13, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
             check=False)
        text(s, ML + 0.72, yy - 0.04, 2.40, 0.36, nm, 20, WHITE, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 3.30, yy + 0.02, 7.60, 0.30, body, 14, C(0xB6, 0xC0, 0xD0), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.86
    hr(s, ML, 5.16, CW, 1.0, C(0x2C, 0x3A, 0x55))
    text(s, ML, 5.32, 1.30, 0.24, "上手路径", 10.5, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    text(s, ML + 1.36, 5.28, 10.40, 0.30,
         "用免费积分跑通一个小任务  →  核结果  →  再决定换模型或付费", 15, WHITE, SANS, True,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 5.76, CW, 0.26,
         "不要一上来就买最贵的：先确认它在你这件活上到底行不行。", 12, DARKGREY, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 6.42, CW, 0.26,
         "附录  常用术语 · 公司档案 · 订阅与省流版 · 资讯渠道 · 分享与工具 · 原理与历史 · 参考资料来源",
         10.5, C(0x6E, 0x7C, 0x93), SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    return s


def p53(prs, pno):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, CH, pno, dark=True)
    text(s, ML, 1.06, 7.60, 0.30, "谢谢", 13, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=2.0,
         check=False)
    text(s, ML, 1.44, 8.20, 0.80, "提问与交流", 44, WHITE, SANS, True, PP_ALIGN.LEFT, 1.14,
         check=False)
    hr(s, ML, 2.38, 7.30, 1.0, C(0x2C, 0x3A, 0x55))
    IT = [("01", "选配置", "先确认型号、入口，以及它能不能真的看见图"),
          ("02", "交代清楚", "目标、材料、格式、怎么算完成；不替它下结论"),
          ("03", "检查结果", "回原文、重算关键数字、打开实际文件")]
    yy = 2.58
    for no, nm, body in IT:
        text(s, ML, yy + 0.03, 0.60, 0.26, no, 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
             check=False)
        text(s, ML + 0.66, yy - 0.02, 1.76, 0.30, nm, 15.5, WHITE, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 2.52, yy + 0.02, 5.30, 0.28, body, 12, C(0xB6, 0xC0, 0xD0), SANS, False,
             PP_ALIGN.LEFT, 1.15, check=False)
        yy += 0.66
    text(s, ML, 5.62, 7.40, 1.06,
         "附录 A–F：常用术语、公司档案、订阅与省流、资讯渠道、分享与工具、原理与历史；Z 为参考资料与来源。\n"
         "配套材料：十份鹈鹕原始 HTML、图纸与报告的核对材料。", 10.5, C(0x6E, 0x7C, 0x93), SANS,
         False, PP_ALIGN.LEFT, 1.5, check=False)
    # 右侧鹈鹕
    pic(s, "s019_1.gif", 9.04, 1.42, 3.70, 2.78, mode="fit")
    text(s, 9.04, 4.32, 3.70, 0.46,
         "同一句提示词，十份结果。\n剩下的时间，交给提问。", 11, C(0x8B, 0x97, 0xAB), SANS, False,
         PP_ALIGN.CENTER, 1.4, check=False)
    return s


def build(prs, start):
    div(prs, start)
    fn = [p40, p41, p42, p43, p44, p45, p46, p47, p48, p49, p50, p51, p52, p53]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
