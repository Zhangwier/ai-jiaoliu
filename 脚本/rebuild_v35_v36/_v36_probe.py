# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\_v36")
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def p_cover(prs):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    pic(s, "s001_1.png", 4.20, 0, W - 4.20, H, mode="fill")
    # 左侧文字区压一层墨蓝
    box(s, 0.18, 0, 4.02, H, fill=INK)
    text(s, 0.90, 0.70, 3.2, 0.30, "工作常用 AI 的使用观察", 12, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.4, check=False)
    text(s, 0.90, 0.98, 3.2, 0.28, "2026 年 10 月", 12, DARKGREY, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, 0.90, 1.62, 3.30, 1.60, "AI\n交流分享", 54, WHITE, SANS, True, PP_ALIGN.LEFT, 1.16,
         check=False)
    box(s, 0.90, 3.34, 0.90, 0.055, fill=TEAL)
    text(s, 0.90, 3.56, 3.10, 0.34, "从原理到实践", 20, TEALL, SANS, True, PP_ALIGN.LEFT, 1.2,
         check=False)
    text(s, 0.90, 4.28, 3.10, 1.50,
         "三个真实测试：\nAI 能做到哪一步\n哪里会出错\n哪些地方必须人来核", 15, C(0xC5, 0xCE, 0xDD), SANS,
         False, PP_ALIGN.LEFT, 1.62, check=False)
    hr(s, 0.90, 6.42, 3.10, 1.0, C(0x2C, 0x3A, 0x55))
    text(s, 0.90, 6.56, 3.10, 0.26, "内部交流材料  ·  请勿外传", 10.5, GREY2, SANS, False, PP_ALIGN.LEFT,
         1.0, check=False)
    return s


def p_divider(prs):
    s = slide(prs, INK)
    box(s, 0, 0, 0.18, H, fill=TEAL)
    header(s, "章节", 6, dark=True)
    text(s, ML, 1.28, 1.4, 1.10, "01", 66, C(0x28, 0x3A, 0x59), SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.34, 1.34, 6.0, 0.30, "PART", 12, TEALL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=2.2,
         check=False)
    text(s, ML + 1.34, 1.60, 8.0, 0.66, "认识 AI", 38, WHITE, SANS, True, PP_ALIGN.LEFT, 1.14,
         check=False)
    text(s, ML + 1.34, 2.34, 7.4, 0.32, "五个短板，每一个都能从原理上讲清", 16, DARKGREY, SANS, False,
         PP_ALIGN.LEFT, 1.3, check=False)
    hr(s, ML, 2.96, CW, 1.0, C(0x2C, 0x3A, 0x55))
    items = [("01", "五个短板", "抽卡、幻觉、迎合、上下文腐烂、奖励黑客"),
             ("02", "从原理看", "学习三态、注意力、训练三步、逐 token 生成"),
             ("03", "五个原因", "为什么它会编、会顺、会忘、会钻空子"),
             ("04", "怎么防", "多跑、要出处、开放提问、控制输入、查过程")]
    yy = 3.22
    for no, nm, sub in items:
        text(s, ML, yy + 0.02, 0.60, 0.26, no, 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.66, yy - 0.03, 2.60, 0.30, nm, 16, WHITE, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 3.30, yy, 8.20, 0.30, sub, 14, C(0xB6, 0xC0, 0xD0), SANS, False, PP_ALIGN.LEFT,
             1.2, check=False)
        yy += 0.86
    return s


def p_five(prs):
    """原稿式表格页：青绿表头 + 白底行 + 橙色问题名 + 青绿处置"""
    s = slide(prs)
    y = head(s, "01  认识 AI", 7, "AI 的五个短板：每一个都能从原理上讲清",
             "说“AI 不可靠”没有用。把不可靠拆成五种，每一种都能对应到原理上的一处成因，也都能对应一条处置办法。")
    rows = [["问题", "表现", "来自哪一步", "怎么防"],
            ["抽卡", "同一句话问两次，结果时好时坏", "按概率抽 token", "多跑几次，关键结论复核"],
            ["幻觉", "编出不存在的条款、判例、数字", "只挑“最像”的下一个字", "要出处，回原文核"],
            ["迎合", "你说有错它就找错；你给的前提它照用", "偏好训练奖励“让人满意”", "开放提问，先让它复述口径"],
            ["上下文腐烂", "聊久了忘了前面提过的要求", "注意力被摊薄", "新开对话，要求放开头"],
            ["奖励黑客", "钻空子、走捷径、多要权限", "奖励只看结果指标", "写清红线，过程也查"]]
    table(s, ML, TOP, CW, rows, [1.45, 4.30, 3.10, 3.20], size=15, head_size=15,
          aligns=[PP_ALIGN.LEFT] * 4, minrow=0.62,
          colors=lambda r, c: ORANGE if c == 0 else (TEAL if c == 3 else C(0x2B, 0x34, 0x45)))
    takeaway(s, "五个短板不是五种“心情”，是五个机制 —— 认出是哪一种，才知道该动哪一处。", y=6.16,
             tag="要点")
    footer(s, "来源：本部分各页原理说明与公开研究；五个短板的分类为本稿归纳")
    return s


def p_cards(prs):
    """三卡对照（原稿常用），但信息密度与层级重排"""
    s = slide(prs)
    y = head(s, "03  两个真实案例", 34, "金额几乎都对，错在数量、原因和旧模板",
             "四份报告列的条目范围不同，不能按条数排名；按“哪一类错”看，差别立刻清楚。")
    items = [("Opus · Hark", "4 分 · 43 条", ["只列 ≥1 万元的条目", "没交代这个筛选口径", "承建单位与中标公告一致"],
              ORANGE, "格式最省事"),
             ("WorkBuddy Auto", "14 分 · 101 条", ["总额写成 41638.89", "承建单位照抄旧标段", "101 条原因全是同一句套话"],
              TEAL, "最像正式报告"),
             ("DeepSeek Flash", "34 分 · 99 条", ["原因写得具体、能落到条目", "但透水管一项套错了条目", "总额与核减数正确"],
              C(0x4A, 0x6C, 0x9E), "原因最具体")]
    cw = 3.86
    for i, (nm, meta, pts, col, tag) in enumerate(items):
        x = ML + i * (cw + 0.29)
        card(s, x, 1.94, cw, 3.72, accent=col)
        text(s, x + 0.28, 2.14, cw - 0.56, 0.30, nm, 17, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.28, 2.46, cw - 0.56, 0.26, meta, 13, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, x + 0.28, 2.82, cw - 0.56, 1.0, LINE)
        bullets(s, x + 0.28, 2.96, cw - 0.56, pts, size=13.5, dot=col, gap=0.20)
        text(s, x + 0.28, 5.20, cw - 0.56, 0.26, tag, 12, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    takeaway(s, "四份都做到了金额基本对；四份都没做到的，是把“为什么变”写回到函件和勘误表。", y=5.96,
             tag="判断")
    footer(s, "来源：四份修编报告原文（截图为各份对应段落）；复算底稿")
    return s


def p_bignum(prs):
    s = slide(prs)
    y = head(s, "02  模型与工具", 20, "十种结果的三个来源：档次、专精、外壳",
             "同一句提示词，十个配置。差别不在“AI 行不行”，在这三件事上。")
    group = [("档次", "GPT-6.1 Sol 与 GPT-6 Luna 在同一榜单上差 20 分以上", "1", ORANGE),
             ("专精", "修编报告最快 4 分钟、最慢 34 分钟，还有一份套错了条目", "2", TEAL),
             ("外壳", "只靠 OCR 的配置，两个检查点零命中", "3", C(0x4A, 0x6C, 0x9E))]
    for i, (nm, body, no, col) in enumerate(group):
        x = ML + i * 4.10
        box(s, x, 1.94, 3.86, 2.44, fill=CARD, line=LINE)
        text(s, x + 0.30, 2.14, 0.50, 0.30, no, 13, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.72, 2.10, 2.90, 0.34, nm, 20, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, x + 0.30, 2.66, 3.26, 1.50, body, 14, GREY, SANS, False, PP_ALIGN.LEFT, 1.55)
    stat(s, ML, 4.66, "20", "分", "同一榜单上 Astra 与 Luna 的差距（AA 智能指数）", ORANGE, 40, 3.6, 12)
    stat(s, ML + 3.90, 4.66, "4→34", "分钟", "修编报告一份的耗时跨度（Opus 4 分、DeepSeek 34 分）", TEAL, 40,
         4.2, 12)
    stat(s, ML + 8.40, 4.66, "0", "命中", "只靠 OCR 的两组配置，两个检查点都没找到", C(0x4A, 0x6C, 0x9E), 40,
         3.4, 12)
    footer(s, "来源：本地鹈鹕盲测与图纸审查记录；AA 智能指数 v4.3.2")
    return s


def main():
    prs = new_deck()
    p_cover(prs)
    p_divider(prs)
    p_five(prs)
    p_cards(prs)
    p_bignum(prs)
    prs.save(r"D:\MG\program\ai交流\_v36\probe.pptx")
    print("ok", len(prs.slides))
    for w in WARN:
        print(w)


main()
