# -*- coding: utf-8 -*-
"""第 04 章 怎么用得稳 + 收尾"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE

CH = "04   怎么用得稳"


# ---------------------------------------------------------------- 章节页（深色，编号大字 + 右侧纲目）
def div04(prs, pno):
    s = slide(prs, DARK)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.14, 2.4, 1.30, "04", 104, C(0x2B,0x34,0x43), MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 2.16, 1.42, 4.0, 0.28, "PART", 9.5, C(0x7E,0x88,0x96), SANS, False,
         PP_ALIGN.LEFT, 1.0, cs=2.4, check=False)
    text(s, ML + 2.16, 1.66, 8.2, 1.00, "怎么用得稳", 40, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.1, check=False)
    text(s, ML + 2.16, 2.52, 7.4, 0.34,
         "前面讲了它会在哪里错；这一章讲怎么把错关在笼子里。", 13, C(0xB2,0xBB,0xC7), SANS, False,
         PP_ALIGN.LEFT, 1.35)
    hr(s, ML, 3.10, CW, 0.75, C(0x38,0x40,0x4E))

    items = [("01", "三类活", "先分类，再决定检查到什么程度"),
             ("02", "交代任务", "目标、材料、格式、怎么算完成"),
             ("03", "管住输入", "上下文、检索、工作流、多 Agent"),
             ("04", "守住边界", "数据分级与最后一步的人工确认")]
    for i, (no, nm, sub) in enumerate(items):
        x = ML + (i % 2) * 5.86
        y = 3.44 + (i // 2) * 1.10
        text(s, x, y, 0.6, 0.24, no, 9.5, C(0x6E,0x78,0x86), MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.62, y - 0.05, 4.6, 0.30, nm, 15, WHITE, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x + 0.62, y + 0.30, 4.9, 0.26, sub, 10.2, C(0x9E,0xA8,0xB6), SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
    text(s, ML, 6.30, CW - 1.2, 0.34,
         "方法不是越多越好：五种方法各治一个毛病，用错地方就是白费工夫。", 11, C(0x8D,0x96,0xA4), SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    return s


# ---------------------------------------------------------------- 40 三类活
def p40(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "先分清是哪类活，再决定查多严",
             "第一类最容易见效，第三类最需要人盯 —— 用同一种检查强度对付三类活，是常见的浪费。")
    COLS = [("01", "找路子", "你不知道有它", "找工具、公式、规范、先例",
             "例：百余个文件里定位一份图纸", 1, "看一眼出处"),
            ("02", "跑量", "量大、规则明确", "批量转表、改名、抽字段、套模板",
             "例：三百份合同抽甲方和金额", 3, "先抽 3 个验格式"),
            ("03", "综合任务", "要判断、要取舍", "修编报告、审图、合同改写、长任务",
             "例：看图、看表、看说明再下结论", 5, "拆步留痕，人签字")]
    cw = 3.47
    for i, (no, nm, why, what, eg, lvl, how) in enumerate(COLS):
        x = ML + i * (cw + 0.56)
        text(s, x, 1.94, 1.6, 0.24, no, 9.5, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.52, 1.90, cw - 0.52, 0.34, nm, 19, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        hr(s, x, 2.34, cw, 1.6, RULE2)
        text(s, x, 2.46, cw, 0.24, why, 10.5, RED if i == 2 else GRN, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x, 2.78, cw, 0.50, what, 10.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        text(s, x, 3.36, cw, 0.26, eg, 9.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        # 检查强度计量条
        text(s, x, 3.86, cw, 0.22, "检查强度", 8.4, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        for k in range(5):
            box(s, x + k * 0.30, 4.14, 0.24, 0.11, fill=RED if k < lvl else C(0xDF,0xDA,0xCD))
        text(s, x, 4.36, cw, 0.24, how, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)

    # 顺序带
    hr(s, ML, 5.06, CW, 0.75, RULE)
    text(s, ML, 5.16, 2.0, 0.22, "上手的顺序", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    steps = ["先用第一类建立信任", "再交给它跑量", "最后才整段交出综合任务"]
    for i, t in enumerate(steps):
        x = ML + 1.92 + i * 3.24
        circ(s, x, 5.27, 0.17, RED)
        text(s, x + 0.12, 5.235, 0.4, 0.18, str(i + 1), 8.2, WHITE, MONO, True, PP_ALIGN.CENTER, 1.0, check=False)
        text(s, x + 0.30, 5.18, 2.80, 0.26, t, 11, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        if i < 2:
            line(s, x + 2.62, 5.27, x + 3.12, 5.27, RULE2, 0.9, head=True)
    takeaway(s, "第三类任务永远要人签字：它做得越像样，越需要有人对结果负责。", y=6.06, color=INK)
    footer(s, "来源：本地两案例的执行记录；三类活的分法与检查强度为本稿整理")


# ---------------------------------------------------------------- 41 第一类实例
def p41(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它肯把一百多个文件翻一遍，这就是它的用处",
             "第一类“找路子”的样子：不是它比你聪明，是它肯把一百多个文件挨个打开、列成一张能筛的表。")

    text(s, ML, 1.94, 7.4, 0.22, "生成出来的表长什么样（实际 9 列，此处 6 列）", 8.8, MUT, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    rows = [["序号", "文件名", "类型", "出图编号", "所在册", "子目录"],
            ["3", "控制成果表.xlsx", "表格", "S1-2-2", "第二册 路线", "02 路线"],
            ["6", "逐桩坐标表.xlsx", "坐标表", "S1-2-7", "第二册 路线", "02 路线"],
            ["8", "砍树挖根数量表.xls", "数量表", "S1-2-10", "第二册 路线", "02 路线"],
            ["…", "……", "……", "……", "……", "……"]]
    table(s, ML, 2.22, 7.38, rows, [0.52, 2.32, 0.94, 1.06, 1.44, 1.10], size=9.6, head_size=8.6,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
          first_bold=False, minrow=0.30, line_color=C(0xD8,0xD2,0xC4))

    # 册汇总
    text(s, ML, 4.34, 7.4, 0.22, "另有一张「册汇总」，先看它决定翻哪一册", 8.8, MUT, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    tot = [("第二册 路线", "26", "9"), ("第三册 一分册 路基", "18", "11")]
    for i, (nm, n1, n2) in enumerate(tot):
        x = ML + i * 3.72
        text(s, x, 4.66, 3.30, 0.24, nm, 11, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x, 4.94, 1.5, 0.30, n1, 17, INK, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.62, 5.02, 0.9, 0.20, "个文件", 8.6, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 1.44, 4.94, 1.2, 0.30, n2, 17, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 2.04, 5.02, 1.2, 0.20, "个数量表", 8.6, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 5.46, 7.38, 0.75, RULE)
    text(s, ML, 5.56, 7.38, 0.72,
         "这就是最土的 RAG：先建索引，再检索。好处是可核对、可重生成 —— 换个项目只改一个路径参数重跑脚本。",
         9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
    takeaway(s, "它的检索能力，要一张自己能核对的表来托底。", y=6.40, color=INK)

    x = 8.62
    vr(s, x - 0.28, 1.94, 4.20, 0.75, RULE)
    text(s, x, 1.94, 3.7, 0.22, "学会用之后，找图变成三件事", 9.0, INK, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    three = [("按册筛", "“第三册第一分册有哪些数量表”—— 筛一列就出来"),
             ("按出图编号定位", "图纸目录上写 S1-2-7，直接查这一列，落到具体文件"),
             ("按类型找同类", "要所有数量表，筛类型；要所有坐标表，同理"),
             ("表里没有的，直接问", "“有没有 XX 桥台的挡块图”—— 它去翻目录和分册表，再告诉你哪一本")]
    yy = 2.28
    for i, (a, b) in enumerate(three):
        text(s, x, yy, 3.7, 0.24, a, 11.2, RED if i == 3 else INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, yy + 0.28, 3.68, 0.62, b, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
        yy += 0.98
    footer(s, "来源：本地文件索引脚本与生成结果（2026-10）")


# ---------------------------------------------------------------- 42 五种方法
def p42(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "五种方法，各治一个毛病",
             "没有一种方法包治百病。先认出这是哪种毛病，再挑对应的那一味。")
    text(s, ML, 1.94, 1.9, 0.22, "毛病", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.9, check=False)
    text(s, 3.52, 1.94, 2.4, 0.22, "方法", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.9, check=False)
    text(s, 5.96, 1.94, 4.2, 0.22, "怎么做", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.9, check=False)
    text(s, 9.18, 1.94, 3.2, 0.22, "第 03 部分的例子", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.9,
         check=False)
    hr(s, ML, 2.20, CW, 1.1, RULE2)

    R = [("迎合、瞎猜", "提示词工程", "写成任务书：范围、格式、验收，允许它说“材料不足”",
          "图纸：范围、证据与验收要求没写清"),
         ("上下文腐烂", "上下文管理", "一个任务一段对话，长了压缩或新开，只放相关材料",
          "长对话后忘了前面的要求"),
         ("幻觉、知识过时", "RAG 检索", "先从资料库找片段，再带出处回答",
          "报告：原因要能回到函件和勘误表"),
         ("抽卡、计算错", "工作流与校验", "固定步骤给程序，数字用脚本重算，多跑几次比对",
          "报告：41,638.89 的笔误"),
         ("看不见就乱猜", "选对输入方式", "图纸先确认是“看图”还是“只 OCR”", "图纸：只靠 OCR 的一个没找到")]
    yy = 2.34
    for i, (ill, mth, how, eg) in enumerate(R):
        text(s, ML, yy + 0.03, 2.46, 0.26, ill, 10.6, RED, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        line(s, 2.92, yy + 0.16, 3.34, yy + 0.16, REDM, 0.9, head=True)
        text(s, 3.52, yy, 2.30, 0.28, mth, 12.4, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, 5.96, yy + 0.02, 3.04, 0.44, how, 10.0, BODY, SANS, False, PP_ALIGN.LEFT, 1.35)
        text(s, 9.18, yy + 0.02, 3.24, 0.44, eg, 9.6, MUT, SANS, False, PP_ALIGN.LEFT, 1.35)
        hr(s, ML, yy + 0.60, CW, 0.55, C(0xE3,0xDE,0xD2))
        yy += 0.68
    takeaway(s, "五种方法不是清单，是配对：认出毛病，才知道该动哪一处。", y=6.10, color=INK)
    footer(s, "来源：本分享第 01 部分的五个短板与第 03 部分的两案例对应关系")


# ---------------------------------------------------------------- 43 任务书
def p43(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "从一句话，到一份能验收的任务书",
             "案例里的原话只有一句。把它拆成四件事，结果的可检查程度就完全不同。")

    # 原话
    box(s, ML, 1.92, 6.10, 1.34, fill=CARD, line=C(0xDE,0xD8,0xCB))
    box(s, ML, 1.92, 0.055, 1.34, fill=RED)
    text(s, ML + 0.30, 2.06, 5.5, 0.22, "案例里的原话", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    text(s, ML + 0.30, 2.34, 5.5, 0.66, "“审查这个有没有问题，结果输出 excel 给我”", 14.5, INK, SERIF, False,
         PP_ALIGN.LEFT, 1.35)
    text(s, ML + 6.36, 1.98, 5.10, 1.24,
         "没说范围、没说格式、没说怎么算完成 —— 这是主要的可改进点。\n“有没有问题”本身不是错问法，缺的是范围、证据与验收。",
         10.0, BODY, SANS, False, PP_ALIGN.LEFT, 1.5)

    # 四要素
    F = [("01", "任务目标", "检查图纸内部是否一致；没有问题也可以"),
         ("02", "材料范围", "先说明读到哪几页、是看图还是 OCR、哪些看不清"),
         ("03", "输出格式", "可编辑 Excel，计算项保留公式"),
         ("04", "验收要求", "每项写位置、依据和把握程度；不补造尺寸")]
    for i, (no, nm, ex) in enumerate(F):
        x = ML + (i % 2) * 5.86
        yy = 3.62 + (i // 2) * 1.30
        text(s, x, yy, 0.52, 0.22, no, 9.2, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.52, yy - 0.05, 4.9, 0.28, nm, 14.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x + 0.52, yy + 0.30, 4.86, 0.28, ex, 10.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    hr(s, ML, 3.48, CW, 0.75, RULE)
    takeaway(s, "写清这四件事，模型的第一句回答就会变成“我先读了哪几页、哪几处看不清”。", y=6.22,
             color=GRN)
    footer(s, "来源：本地图纸审查案例的原始提示词与后续改写；任务书四要素为本稿归纳")


# ---------------------------------------------------------------- 44 交付四步
def p44(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "每一步都留下能查的东西，下一步才有可靠输入",
             "同一条流水线在审图和写报告里都成立。区别只在每步交出来的东西长什么样。")
    S = [("1", "读材料", "列出读了哪些文件、范围、缺页和认不出的地方",
          "清单 ＋ 函件 ＋ 勘误表"),
         ("2", "定口径", "说清单位、金额范围、资料优先级、哪些不许猜",
          "汇总表口径，不含暂列金额"),
         ("3", "做产物", "按模板交可编辑文件，保留计算过程和出处",
          "报告 ＋ 变更表"),
         ("4", "查结果", "抽查事实、自己重算、打开文件看版面",
          "金额复算 ＋ 原因定位 ＋ 模板残留")]
    cw = 2.68
    for i, (no, nm, what, eg) in enumerate(S):
        x = ML + i * (cw + 0.30)
        text(s, x, 1.92, 0.5, 0.26, no, 10, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.36, 1.88, cw - 0.36, 0.30, nm, 15, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        hr(s, x, 2.28, cw, 1.4, INK if i == 3 else RULE2)
        text(s, x, 2.40, cw, 0.84, what, 10.0, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
        text(s, x, 3.34, cw, 0.20, "修编报告里对应", 8.2, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.7,
             check=False)
        text(s, x, 3.58, cw, 0.60, eg, 9.6, GRN, SANS, True, PP_ALIGN.LEFT, 1.4)
        if i < 3:
            line(s, x + cw + 0.04, 2.40, x + cw + 0.26, 2.40, RULE2, 0.9, head=True)

    # 下游依赖
    hr(s, ML, 4.50, CW, 0.75, RULE)
    text(s, ML, 4.62, 3.0, 0.22, "这条链最容易断在两处", 8.8, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    W = [("口径没定就开做", "后面所有数字都要重算一遍，返工量最大。"),
         ("产物不可编辑", "只能看不能改，等于把复核的机会一起交出去了。")]
    for i, (a, b) in enumerate(W):
        x = ML + i * 5.86
        text(s, x, 4.96, 5.30, 0.26, a, 11.5, RED, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, 5.26, 5.30, 0.28, b, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    takeaway(s, "上一步有看得见的结果，下一步才有可靠的输入。", y=6.06, color=INK)
    footer(s, "来源：本地修编报告案例的实际执行顺序与返工记录")


# ---------------------------------------------------------------- 45 上下文
def p45(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "控制放进去的东西，比换更长窗口更有效",
             "对付上下文腐烂：窗口 1M 不等于要塞满 1M —— 放得越精，答得越准。")
    R = [("01", "一个任务一段对话", "换任务就新开，旧材料不带过去"),
         ("02", "长了就压缩", "用客户端的“压缩上下文”，或让它写进度摘要后新开"),
         ("03", "只放相关的", "先让它说需要哪几页、哪张表，再给"),
         ("04", "规则写进文件", "单位口径、格式、禁止事项写成规则文件，每次自动带上"),
         ("05", "要求放开头", "关键要求放最前，必要时结尾再说一遍")]
    yy = 1.94
    for no, nm, what in R:
        text(s, ML, yy + 0.02, 0.44, 0.22, no, 9.2, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.46, yy - 0.03, 5.06, 0.28, nm, 13, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML + 0.46, yy + 0.30, 5.02, 0.26, what, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.35,
             check=False)
        hr(s, ML, yy + 0.64, 5.52, 0.55, C(0xE3,0xDE,0xD2))
        yy += 0.86

    # 规则文件（深色代码块）
    x = 6.92
    box(s, x, 1.92, 5.60, 3.36, fill=DARK)
    text(s, x + 0.32, 2.06, 4.9, 0.22, "规则文件长什么样（示例）", 8.8, C(0x8D,0x96,0xA4), SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.7, check=False)
    code = ["# 项目规则", "- 金额单位：万元，保留 4 位小数", "- 汇总口径：不含暂列金额",
            "- 合同条文原样引用，不改写", "- 原因必须注明函件编号或勘误表序号",
            "- 找不到依据就写“待核”，不要补"]
    cols = [C(0x8D,0x96,0xA4), C(0xDD,0xE2,0xE8), C(0xDD,0xE2,0xE8), C(0xDD,0xE2,0xE8), C(0xDD,0xE2,0xE8),
            C(0xF0,0xC6,0xB4)]
    for i, ln in enumerate(code):
        text(s, x + 0.32, 2.40 + i * 0.42, 5.0, 0.30, ln, 10.0 if i else 10.0, cols[i], MONO, False,
             PP_ALIGN.LEFT, 1.0, check=False)
    text(s, x + 0.32, 4.94, 5.0, 0.22, "每次新开对话，先把这段贴进去或放到项目根目录", 8.8, C(0x7E,0x88,0x96),
         SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    takeaway(s, "把规则写成文件，就不用每轮重新叮嘱一遍。", y=5.72, color=INK)
    footer(s, "来源：Chroma《Context Rot》；Claude Code、Codex 官方文档中的上下文与规则文件说明")


# ---------------------------------------------------------------- 46 RAG
def p46(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "资料多的时候，先找再答",
             "两个理由：省一个量级的 token，少一批凭空编出来的答案。做法只有一句 —— 先找，再答。")
    # 两个理由
    R = [("① 省 token", "省一个量级",
          "100 页整包 ≈ 5–8 万 token，问一次付一次；\n只检索命中 3–5 段 ≈ 1–2 千 token。\n差几十倍，每追问一次就再付一遍。"),
         ("② 少编", "把“凭印象”换成“照着原文”",
          "材料不在眼前时，它只能按“最像”续写，数字和条文最容易编；\n原文摆到眼前，就有东西可抄，还能标出处。")]
    for i, (t1, t2, body) in enumerate(R):
        x = ML + i * 3.30
        text(s, x, 1.90, 3.06, 0.22, t1, 9.4, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
        text(s, x, 2.16, 3.10, 0.26, t2, 12.4, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x, 2.48, 3.08, 0.90, body, 9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.42)

    # 流程五步
    hr(s, ML, 3.58, 8.18, 0.75, RULE)
    text(s, ML, 3.68, 4.0, 0.22, "怎么做：五个环节，只有第一步是一次性的", 8.8, MUT, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    P = ["入库\n切分建索引", "检索\n关键词/向量取候选", "精排\n筛出真正的证据", "拼装\n证据放进提示词", "回答\n每句带出处"]
    for i, t in enumerate(P):
        x = ML + i * 1.66
        box(s, x, 3.96, 1.44, 0.62, fill=CARD if i else GRNL, line=C(0xDE,0xD8,0xCB))
        text(s, x + 0.12, 4.04, 1.22, 0.48, t, 8.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        if i < 4:
            line(s, x + 1.46, 4.27, x + 1.62, 4.27, RULE2, 0.9, head=True)

    # 关键词 vs 向量
    text(s, ML, 4.86, 8.18, 0.22, "靠什么找：关键词、向量，或者两者混合", 8.8, MUT, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    K = [("关键词空手", "“防撞设施”与“挡块”一个字都不重合，关键词搜索直接空手", RED),
         ("语义能落到", "两段话意思近，向量就靠得近；取最近三五段，正好把“挡块”那段捞上来", GRN),
         ("但“近”≠“对”", "近邻里常有看着像、其实不是一回事的段落：要再精排一次，并给页码回原文核", AMB)]
    yy = 5.12
    for a, b, col in K:
        text(s, ML, yy, 1.62, 0.22, a, 9.4, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 1.70, yy, 6.44, 0.42, b, 9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.35)
        yy += 0.50

    # 右：嵌入散点
    pic(s, "s114_1.png", 9.34, 1.90, 3.18, 2.28, mode="fit")
    text(s, 9.34, 4.24, 3.18, 0.22, "词向量示例：意思近的词自动聚成一簇", 8.2, MUT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    box(s, 9.34, 4.64, 3.18, 1.52, fill=GRNL)
    text(s, 9.52, 4.76, 2.82, 1.30,
         "实践多用混合检索：编号、专有名词靠关键词；换了说法靠向量；两路结果再精排一次，回答必须带出处。",
         9.2, GRN, SANS, False, PP_ALIGN.LEFT, 1.45)
    footer(s, "来源：Chroma《Context Rot》；Lewis 等 RAG 原论文（2020）；各厂 embedding 与计费文档")


# ---------------------------------------------------------------- 47 工作流
def p47(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "把“谁做、按什么判、交出什么”写进每一步",
             "审图和写报告，骨架其实是同一个：拆步骤，每步定三件事。")
    rows = [["步骤", "谁做", "按什么判", "交出什么"],
            ["① 定范围", "人", "页数、图号、哪些表；先确认是“看图”还是“只 OCR”", "一页清单"],
            ["② 提取", "OCR ＋ 视觉", "页数、表名、行数对得上；缺页、糊字要点名", "结构化底表"],
            ["③ 逐项核对", "模型", "每条要有图中位置 ＋ 原文；不确定进“待确认”", "问题清单"],
            ["④ 复算", "脚本", "关键量能独立复算；算不出来不进正文", "计算底稿"],
            ["⑤ 汇总结论", "人", "分“不一致 / 待确认 / 未检查”，不许混写", "审查意见"]]
    table(s, ML, 1.90, 7.72, rows, [1.16, 1.10, 4.12, 1.34], size=9.2, head_size=8.6,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT, PP_ALIGN.LEFT],
          minrow=0.52, line_color=C(0xD8,0xD2,0xC4))

    x = 8.82
    text(s, x, 1.90, 3.60, 0.22, "换成修编报告，同一条链", 8.8, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6,
         check=False)
    M = [("读清单", "行数、口径、单位对得上"),
         ("定位变更", "每条变更能指到函件序号"),
         ("套模板", "章节、单元格位置不变"),
         ("复算金额", "112 元差额能说清来源"),
         ("写原因", "原因必须有据，不许编")]
    yy = 2.18
    for i, (a, b) in enumerate(M):
        text(s, x, yy, 1.30, 0.22, a, 9.6, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 1.24, yy, 2.42, 0.36, b, 9.0, BODY, SANS, False, PP_ALIGN.LEFT, 1.25)
        hr(s, x, yy + 0.40, 3.62, 0.5, C(0xE3,0xDE,0xD2))
        yy += 0.46

    # 三个坑
    hr(s, ML, 4.96, CW, 0.75, RULE)
    text(s, ML, 5.06, 3.4, 0.22, "两个案例踩过的坑，都出在“标准没定死”", 8.8, RED, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=0.6, check=False)
    pits = [("把 mm/cm 当错误", "没先定“看图例”"),
            ("汇总与分项差 112 元", "没定口径"),
            ("核减数算不出", "没做复算")]
    for i, (a, b) in enumerate(pits):
        x2 = ML + i * 3.86
        text(s, x2, 5.38, 3.62, 0.24, a, 10.8, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x2, 5.66, 3.62, 0.22, b, 9.2, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    takeaway(s, "工作流是长出来的：每跑一次，把踩到的坑写成一条标准塞回对应步骤。", y=6.14, color=INK)
    footer(s, "来源：本稿图纸审查与修编报告两案例的实际执行与复核记录")


# ---------------------------------------------------------------- 48 颗粒度
def p48(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "哪些该写死，哪些该留白",
             "这也解释了一件事：为什么“整理成工作流”和“直接交给 Agent”是同一条线上的两档。")

    # 写死
    text(s, ML, 1.94, 5.30, 0.24, "该写死的：三条全中", 9.4, INK, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    hr(s, ML, 2.22, 5.30, 1.4, INK)
    A = [("格式固定", "Excel 行列、模板单元格、报告章节都是死的"),
         ("有唯一正确答案", "合计数、差额、核减额，算错就是错"),
         ("出错代价高", "要能复算、能追责，不能“大概对”")]
    yy = 2.38
    for a, b in A:
        text(s, ML, yy, 5.30, 0.24, a, 11.4, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML, yy + 0.28, 5.28, 0.26, b, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
        yy += 0.82

    # 留白
    x = 6.86
    text(s, x, 1.94, 5.60, 0.24, "该留白的：中一条就交给模型", 9.4, GRN, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.8, check=False)
    hr(s, x, 2.22, 5.60, 1.4, GRN)
    B = [("材料不全", "缺页、口径不清，要临场判断"),
         ("输出是自然语言", "写原因、写信、写说明"),
         ("步骤说不清", "要边做边看结果再决定下一步")]
    yy = 2.38
    for a, b in B:
        text(s, x, yy, 5.60, 0.24, a, 11.4, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, yy + 0.28, 5.56, 0.26, b, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
        yy += 0.72
    text(s, x, 4.62, 5.60, 0.22, "怎么留活口：三条", 8.8, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6,
         check=False)
    for i, t in enumerate(["标准写“要达到什么”，不写“怎么做”", "每一步内部留一个判断位，允许它换做法",
                           "定期复核规则：确认已过时再删"]):
        text(s, x, 4.90 + i * 0.30, 5.56, 0.24, "·  " + t, 9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)

    # 过细的害处
    box(s, ML, 5.92, CW, 0.66, fill=REDL)
    text(s, ML + 0.24, 6.02, 1.28, 0.22, "写过细", 8.8, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    harm = ["拆分反而掉分：本来一次能做对，拆成五步每步掉一点",
            "标准写成操作：它发现更快的路也不许走",
            "给旧模型打的补丁：一升级全成负担"]
    for i, t in enumerate(harm):
        text(s, ML + 1.62 + i * 3.34, 6.02, 3.20, 0.44, t, 9.2, C(0x6E,0x2A,0x1C), SANS, False,
             PP_ALIGN.LEFT, 1.3)
    footer(s, "来源：本稿两案例的执行记录；各模型客户端说明；模型能力变化的观察（2026-10）")


# ---------------------------------------------------------------- 49 多 Agent
def p49(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "人多不自动变可靠，拆得开才值得拆",
             "收益来自并行和上下文隔离；代价来自协调、文件冲突，以及“同一个错误前提”。")

    # 左：收益 / 代价
    text(s, ML, 1.92, 5.40, 0.22, "什么时候值：两条真实收益", 9.0, GRN, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.6, check=False)
    G = [("并行", "三个 Agent 各查一份清单的不同章节：串行 40 分钟的活，十几分钟出初稿"),
         ("上下文干净", "每个只看自己那一摊，主线不被几十页材料冲淡")]
    yy = 2.22
    for a, b in G:
        text(s, ML, yy, 5.40, 0.24, a, 11.4, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML, yy + 0.28, 5.36, 0.50, b, 9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.4)
        yy += 0.92
    text(s, ML, 4.10, 5.40, 0.22, "什么时候只是更贵：三条真实代价", 9.0, RED, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.6, check=False)
    C3 = [("成本翻倍", "三个 Agent ＝ 三份 token，还得加上协调的往返"),
          ("写冲突", "两个改同一个文件，后写的把前面的覆盖掉"),
          ("错误前提传染", "三个都用了错的汇总口径 → 一起错，还像“交叉验证过”")]
    yy = 4.38
    for a, b in C3:
        text(s, ML, yy, 1.34, 0.22, a, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 1.40, yy, 4.00, 0.40, b, 9.2, BODY, SANS, False, PP_ALIGN.LEFT, 1.3)
        yy += 0.46

    # 右：门槛
    x = 6.86
    vr(s, x - 0.24, 1.92, 3.30, 0.75, RULE)
    text(s, x, 1.92, 5.60, 0.22, "拆之前先过这四条，缺一就别拆", 9.0, INK, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.6, check=False)
    Q = ["任务能切成互不依赖的块", "每块有独立、可核对的产出", "写文件的范围不重叠", "有人负责合并并验收"]
    for i, t in enumerate(Q):
        yy = 2.24 + i * 0.46
        circ(s, x + 0.10, yy + 0.13, 0.26, None, RULE2, 0.9)
        text(s, x + 0.05, yy + 0.045, 0.30, 0.18, str(i + 1), 8.4, MUT, MONO, True, PP_ALIGN.CENTER, 1.0,
             check=False)
        text(s, x + 0.44, yy, 4.90, 0.26, t, 10.4, INK, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, x, 4.10, 5.50, 0.44, "这几种就别拆：步骤前后强依赖 · 总量本来就小 · 没有验收人", 9.4, MUT, SANS,
         False, PP_ALIGN.LEFT, 1.35)

    # 假复核
    box(s, x, 4.70, 5.52, 1.20, fill=AMBL)
    text(s, x + 0.24, 4.82, 5.04, 0.22, "最常见的一种假复核", 8.8, AMB, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=0.6, check=False)
    text(s, x + 0.24, 5.08, 5.04, 0.72,
         "复核的人只复述执行者的答案，等于没复核；必须回到原材料逐条对。", 9.8, C(0x53,0x3E,0x12), SANS,
         False, PP_ALIGN.LEFT, 1.4)
    takeaway(s, "拆得开、验得了、有人兜底，才值得多开一个 Agent。", y=6.18, color=INK)
    footer(s, "来源：本稿多 Agent 协作记录；Subagent 与并行任务说明；客户端文档")


# ---------------------------------------------------------------- 50 三个陷阱
def p50(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三个坑，都是我们自己踩过的",
             "前面讲怎么搭，这一页讲它什么时候会坏事。")
    T = [("坑一", "低级模型 ＋ 复杂工作流，\n不如高级模型一句话",
          "工作流不增加理解力，只会把同一个误解重复五遍。本次差的几组不是步骤不够细，是没看懂图。",
          "先确认模型够用，再谈工作流"),
         ("坑二", "不逐步确认，\n跑完查不出哪一步坏了",
          "中间产物没留，结果不对只能整条重跑，时间全耗在重跑上。",
          "每一步留一个能查的产物"),
         ("坑三", "一个 AI 没看懂，\n换一个 AI 复核也看不懂",
          "复核者拿到的还是那段模糊描述，只会给出另一个模糊判断。",
          "复核要能回到原文逐条对")]
    cw = 3.47
    for i, (tag, title, why, how) in enumerate(T):
        x = ML + i * (cw + 0.56)
        text(s, x, 1.92, 1.2, 0.22, tag, 9.2, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
        text(s, x, 2.20, cw, 0.86, title, 14.5, INK, SANS, True, PP_ALIGN.LEFT, 1.35)
        hr(s, x, 3.14, cw, 0.75, RULE2)
        text(s, x, 3.26, cw, 1.10, why, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.5)
        text(s, x, 4.56, cw, 0.20, "怎么避", 8.2, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        text(s, x, 4.80, cw, 0.28, how, 10.4, GRN, SANS, True, PP_ALIGN.LEFT, 1.25, check=False)

    hr(s, ML, 5.44, CW, 0.75, RULE)
    text(s, ML, 5.56, CW, 0.62,
         "三条坑指向同一句话：先确认它看得懂，再决定拆几步、用几个 Agent。", 14, INK, SERIF, False,
         PP_ALIGN.LEFT, 1.4)
    footer(s, "来源：本稿两案例的执行与返工记录；多 Agent 协作记录")


# ---------------------------------------------------------------- 51 数据分级
def p51(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "数据分三级，最后一步永远由人确认",
             "越敏感的数据、越不能撤回的动作，越要人在最后一步把关。")

    LV = [("公开资料", "按允许的流程使用", GRNL, GRN),
          ("内部业务材料", "脱敏，用单位允许的工具和账号", AMBL, AMB),
          ("密钥与高敏感资料", "不贴进普通对话，用专门的凭据机制", REDL, RED)]
    yy = 1.94
    for i, (nm, how, bg, fg) in enumerate(LV):
        box(s, ML, yy, 6.34, 1.00, fill=bg)
        text(s, ML + 0.28, yy + 0.16, 2.40, 0.26, nm, 12.4, fg, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML + 0.28, yy + 0.50, 5.80, 0.40, how, 9.8, C(0x3A,0x3F,0x4A), SANS, False, PP_ALIGN.LEFT,
             1.35)
        yy += 1.12

    # 操作权限
    x = 7.56
    text(s, x, 1.94, 4.90, 0.24, "操作权限：只给完成任务必需的那部分", 9.4, INK, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=0.6, check=False)
    hr(s, x, 2.22, 4.90, 1.4, RULE2)
    acts = ["读写本地文件、执行命令", "只给完成任务必需的权限", "对外发送、付款、删除",
            "最后一步由人核对对象、内容、金额和后果"]
    for i, t in enumerate(acts):
        yy = 2.40 + i * 0.50
        text(s, x, yy, 0.30, 0.22, "·", 12, RED if i >= 2 else MUT, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.24, yy, 4.66, 0.44, t, 9.6, BODY if i < 3 else INK, SANS, i == 3, PP_ALIGN.LEFT, 1.35)

    # 误传处置
    hr(s, ML, 5.34, CW, 0.75, RULE)
    text(s, ML, 5.44, 3.0, 0.22, "已经误传了怎么办", 8.8, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    FX = ["马上停止扩散", "撤销或轮换凭据", "保留必要的记录", "按单位流程处理"]
    for i, t in enumerate(FX):
        xx = ML + i * 2.90
        text(s, xx, 5.72, 0.40, 0.22, str(i + 1), 9.2, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, xx + 0.30, 5.70, 2.42, 0.26, t, 10.4, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        if i < 3:
            line(s, xx + 2.30, 5.82, xx + 2.72, 5.82, RULE2, 0.9, head=True)
    takeaway(s, "权限、付款、删除、外发这四件事，人签字之前不自动执行。", y=6.28, color=INK)
    footer(s, "来源：工作资料与权限管理方法（本稿整理）")


# ---------------------------------------------------------------- 52 带走三件事
def p52(prs, pno):
    s = slide(prs, DARK)
    header(s, "收尾", pno, dark=True)
    text(s, ML, 1.16, 8.0, 0.30, "带走三件事", 9.5, C(0x8D,0x96,0xA4), SANS, False, PP_ALIGN.LEFT, 1.0,
         cs=2.2, check=False)
    text(s, ML, 1.56, 11.0, 0.72, "选好配置，交代清楚，检查结果", 40, WHITE, SERIF, False, PP_ALIGN.LEFT,
         1.15, check=False)

    T = [("01", "选配置", "确认型号和入口；图纸类任务先确认它真能看见图"),
         ("02", "交代清楚", "目标、材料、格式、怎么算完成；不替它下结论"),
         ("03", "检查结果", "回原文、重算关键数字、打开实际文件")]
    for i, (no, nm, body) in enumerate(T):
        x = ML + i * 3.86
        text(s, x, 3.04, 1.0, 0.24, no, 9.5, C(0x7E,0x88,0x96), MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, x, 3.34, 3.56, 1.4, C(0x4A,0x54,0x63))
        text(s, x, 3.50, 3.54, 0.34, nm, 19, WHITE, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x, 3.94, 3.50, 0.70, body, 10.4, C(0xA8,0xB2,0xC0), SANS, False, PP_ALIGN.LEFT, 1.5)

    hr(s, ML, 5.28, CW, 0.75, C(0x38,0x40,0x4E))
    text(s, ML, 5.44, 2.0, 0.22, "上手路径", 8.8, C(0x7E,0x88,0x96), SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, ML + 1.52, 5.40, 9.9, 0.30,
         "用免费积分跑通一个小任务  →  核结果  →  再决定换模型或付费", 12.5, WHITE, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    return s


# ---------------------------------------------------------------- 53 谢谢
def p53(prs, pno):
    s = slide(prs, DARK)
    header(s, "收尾", pno, dark=True)
    text(s, ML, 1.62, 7.4, 1.10, "谢谢", 60, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.1, check=False)
    text(s, ML, 2.72, 7.4, 0.40, "提问与交流", 17, C(0xB2,0xBB,0xC7), SANS, False, PP_ALIGN.LEFT, 1.3,
         check=False)
    hr(s, ML, 3.42, 7.10, 0.75, C(0x38,0x40,0x4E))
    R = [("01", "选配置", "先确认型号、入口，以及它能不能真的看见图"),
         ("02", "交代清楚", "目标、材料、格式、怎么算完成；不替它下结论"),
         ("03", "检查结果", "回原文、重算关键数字、打开实际文件")]
    yy = 3.66
    for no, nm, body in R:
        text(s, ML, yy + 0.02, 0.5, 0.22, no, 9.2, C(0x7E,0x88,0x96), MONO, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.54, yy - 0.03, 1.30, 0.26, nm, 12.5, WHITE, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, ML + 2.02, yy, 5.10, 0.26, body, 10.2, C(0xA8,0xB2,0xC0), SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
        yy += 0.60
    text(s, ML, 5.72, 7.30, 1.10,
         "附录 A–F：常用术语、公司档案、订阅与省流、资讯渠道、分享与工具、原理与历史；Z 为参考资料与来源。\n配套材料里有十份鹈鹕原始 HTML、图纸与报告的核对材料。",
         9.4, C(0x7E,0x88,0x96), SANS, False, PP_ALIGN.LEFT, 1.55)
    # 右侧留白装饰：一条竖线与素材缩略
    pic(s, "s019_1.gif", 8.96, 1.36, 3.52, 2.64, mode="fit")
    text(s, 8.96, 4.10, 3.52, 0.22, "同一句提示词，十个配置的鹈鹕之一", 8.4, C(0x7E,0x88,0x96), SANS,
         False, PP_ALIGN.LEFT, 1.0, check=False)
    return s


def build(prs, start):
    div04(prs, start)
    fn = [p40, p41, p42, p43, p44, p45, p46, p47, p48, p49, p50]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
    p51(prs, start + 12)
    p52(prs, start + 13)
    p53(prs, start + 14)
