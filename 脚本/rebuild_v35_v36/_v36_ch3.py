# -*- coding: utf-8 -*-
"""第 03 章 两个真实案例"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "03  两个真实案例"

RES = [("Opus 5.5 · Hark", "3分03秒", 1, 1, 2, 2, 11, "直接看图（自述人工识读）"),
       ("GPT 6.1 Sol high · Codex", "14分34秒", 1, 0, 4, 1, 8, "OCR 定位 + 回看原图"),
       ("Qwen 3.8 Flash · Qoder", "43分09秒", 0, 1, 2, 5, 24, "渲染成图，逐格读"),
       ("Auto fast · WorkBuddy", "2分38秒", 0, 0, 1, 2, 7, "未说明"),
       ("DeepSeek Flash · WorkBuddy", "32分05秒", 0, 0, 1, 2, 19, "自述“图面识读”"),
       ("Step 5 Preview · WorkBuddy", "29分40秒", 0, 0, 1, 8, 26, "未说明"),
       ("MiniMax M3 · WorkBuddy", "12分27秒", 0, 0, 0, 5, 10, "未说明"),
       ("Seed 2.1 Turbo · Trae", "12分47秒", 0, 0, 0, 3, 10, "OCR"),
       ("GLM 5.3 Flash · Trae", "1时24分", 0, 0, 0, 0, 1, "渲染后放大看"),
       ("HY4 Preview · WorkBuddy", "3分55秒", 0, 0, 0, 2, 2, "视觉未启用")]


def div(prs, pno):
    s = slide(prs)
    box(s, 0, 0, W, H, fill=PAPER)
    pic(s, "s029_1.png", 6.62, 0.0, W - 6.62, H, mode="fill")
    box(s, 0, 0, 6.62, H, fill=PAPER)
    box(s, 0, 0, 0.18, H, fill=ORANGE)
    header(s, "章节", pno)
    text(s, ML, 1.30, 1.5, 1.10, "03", 68, C(0xD8, 0xD3, 0xC5), SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 1.36, 1.36, 5.0, 0.30, "PART  03", 12, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=2.2, check=False)
    text(s, ML + 1.36, 1.64, 4.90, 1.40,
         "两件事，\n一百多条可核对的结论。", 32, NAVY, SANS, True, PP_ALIGN.LEFT, 1.30,
         check=False)
    text(s, ML + 1.36, 3.10, 4.80, 0.62,
         "两次都是真实交付物：同一张工程图、同一份清单，\n交给十个配置，然后逐条回去核。", 13.5, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    hr(s, ML, 3.98, 5.90, 1.0, LINE2)
    CA = [("案例一", "挡块图纸审查", "十个配置看同一张图，三个找到关键问题"),
          ("案例二", "修编报告", "四个配置写同一份报告，金额几乎都对")]
    yy = 4.16
    for tag, nm, body in CA:
        text(s, ML, yy + 0.02, 0.86, 0.26, tag, 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             cs=1.0, check=False)
        text(s, ML + 0.92, yy - 0.04, 4.90, 0.30, nm, 16.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 0.92, yy + 0.30, 4.96, 0.28, body, 12, GREY, SANS, False, PP_ALIGN.LEFT, 1.25,
             check=False)
        yy += 0.92
    return s


def p29(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一张挡块图，十个配置，只有三个找到了这两处",
             "提示词只有一句：“审查这个有没有问题，结果输出 excel 给我”。两个检查点，是后来人工复核确认的。")
    top = y + 0.14
    # 左：整图
    pic(s, "s029_1.png", ML, top, 6.10, 2.92, mode="fit")
    text(s, ML, top + 3.00, 6.10, 0.24, "挡块图纸全图：立面、平面、钢筋大样", 10.5, FAINT, SANS,
         False, PP_ALIGN.LEFT, 1.0, check=False)
    # 右：两处放大
    RX = ML + 6.36
    pic(s, "s029_2.png", RX, top, 2.34, 1.86, mode="fit")
    pic(s, "s029_3.png", RX + 2.44, top, 1.66, 1.86, mode="fit")
    text(s, RX, top + 1.94, 2.34, 0.24, "1a 大样：各段之和", 10, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, RX + 2.44, top + 1.94, 1.66, 0.24, "3a 引线：sinα / sinβ", 10, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    # 两个检查点
    KP = [("K1", "分段与总长对不上", "1a 大样各段加起来，比印刷的总长公式多 45 cm", ORANGE),
          ("K2", "符号不一致", "3a 图上写 sinα，引线公式里写 sinβ", TEAL)]
    yy = top + 2.32
    for no, nm, body, col in KP:
        box(s, RX, yy, 4.10, 0.62, fill=CARD, line=LINE)
        box(s, RX, yy, 0.05, 0.62, fill=col)
        text(s, RX + 0.20, yy + 0.08, 0.42, 0.24, no, 12, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 0.66, yy + 0.05, 3.30, 0.24, nm, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, RX + 0.66, yy + 0.31, 3.34, 0.24, body, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.72
    # 口径
    box(s, ML, 6.00, CW, 0.66, fill=C(0xF3, 0xF1, 0xEB))
    text(s, ML + 0.22, 6.10, 1.00, 0.24, "口径", 11, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.6,
         check=False)
    vr(s, ML + 1.12, 6.12, 0.42, 1.0, LINE2)
    text(s, ML + 1.30, 6.09, CW - 1.50, 0.48,
         "单位 mm/cm 经工程复核不算问题，只看 K1、K2 两项：分母是 2，不是“准确率”。另有 21 条待核，未计入。",
         12.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：本稿挡块图纸审查记录（2026-10）；原图由项目方提供")
    return s


def p30(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三个配置命中，假错误却累计三十条",
             "按命中数排序。点表示命中；右侧堆叠条是“有效条目 + 假错误”，统一按 26 条为满格。")
    ty = y + 0.10
    CX = [ML, ML + 3.02, ML + 4.62, ML + 5.14, ML + 6.62, ML + 7.14, ML + 8.66,
          ML + 10.86]
    HD = ["配置", "耗时", "K1", "K2", "其他", "堆叠（有效 / 假错误）", "共", "怎么读的图"]
    for i, h in enumerate(HD):
        al = PP_ALIGN.CENTER if i in (2, 3, 4, 6) else PP_ALIGN.LEFT
        text(s, CX[i], ty, (CX[i + 1] - CX[i] - 0.10) if i < 7 else 1.60, 0.24, h, 10, GREY2, SANS,
             True, al, 1.0, cs=0.6, check=False)
    hr(s, ML, ty + 0.28, CW, 1.2, LINE2)
    BW = 1.66     # 堆叠条满格宽度（26 条）
    yy = ty + 0.36
    RH = 0.382
    for i, (nm, tm, k1, k2, oth, fake, tot, how) in enumerate(RES):
        hit = k1 + k2
        if i:
            hr(s, ML, yy - 0.02, CW, 0.8, LINE)
        ncol = NAVY if hit else GREY
        text(s, CX[0], yy + 0.02, 2.94, 0.26, nm, 10.8, ncol, SANS, hit > 0, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, CX[1], yy + 0.03, 1.50, 0.24, tm, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        for j, v in enumerate((k1, k2)):
            cx = CX[2 + j] + 0.22
            cy = yy + 0.15
            if v:
                circ(s, cx, cy, 0.135, TEAL)
            else:
                circ(s, cx, cy, 0.135, CARD, LINE2, 1.0)
        text(s, CX[4], yy + 0.03, 0.44, 0.24, str(oth), 10.5, GREY2, SANS, False, PP_ALIGN.CENTER,
             1.0, check=False)
        # 堆叠条
        bx = CX[5]
        by = yy + 0.10
        eff = tot - fake
        w1 = BW * (eff / 26.0)
        w2 = BW * (fake / 26.0)
        box(s, bx, by, BW, 0.16, fill=C(0xF0, 0xEE, 0xE8))
        if w1 > 0.01:
            box(s, bx, by, max(0.02, w1), 0.16, fill=TEAL)
        if w2 > 0.01:
            box(s, bx + max(0.02, w1), by, max(0.02, w2), 0.16, fill=ORANGE)
        text(s, CX[6], yy + 0.03, 0.44, 0.24, str(tot), 10.8, NAVY, SANS, True, PP_ALIGN.CENTER,
             1.0, check=False)
        text(s, CX[7], yy + 0.03, X2 - CX[7], 0.24, how, 9.8,
             ORANGE if "未" in how or how in ("OCR", "视觉未启用") else GREY2, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
        yy += RH
    hr(s, ML, yy - 0.02, CW, 1.2, LINE2)
    # 图例
    ly = yy + 0.12
    box(s, ML, ly + 0.06, 0.20, 0.10, fill=TEAL)
    text(s, ML + 0.26, ly, 1.30, 0.22, "有效条目", 9.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    box(s, ML + 1.42, ly + 0.06, 0.20, 0.10, fill=ORANGE)
    text(s, ML + 1.68, ly, 3.10, 0.22, "假错误（说得具体，和原图对不上）", 9.5, GREY2, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 5.60, ly, 6.60, 0.22,
         "“假错误”含把单位 mm/cm 列为错误的情况。", 9.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    box(s, ML, 6.06, CW, 0.62, fill=TEALXL)
    text(s, ML + 0.24, 6.16, CW - 0.48, 0.44,
         "命中率高不等于可靠：Opus 3 分 03 秒命中两项，Step 5 用了近半小时，26 条里 8 条是假的。", 14,
         NAVY, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：本稿十份审查输出逐条回原图核对；21 条待核未计")
    return s


def p31(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "只靠 OCR 的配置，一条都没命中",
             "把十个配置按“怎么读图”重新分组，命中与不命中的分界就出现了。")
    top = y + 0.16
    # 左组
    LW = 5.60
    text(s, ML, top, LW, 0.26, "真的看到了图", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2,
         check=False)
    hr(s, ML, top + 0.30, LW, 2.2, TEAL)
    G1 = [("Opus 5.5 · Hark", "自述人工识读整图", "K1 + K2"),
          ("GPT 6.1 Sol · Codex", "OCR 定位后回看原图", "K1"),
          ("Qwen 3.8 Flash · Qoder", "文字提取无效，改渲染成图逐格读", "K2"),
          ("GLM 5.3 Flash · Trae", "渲染后放大读图，用时 1 小时 24 分", "0")]
    yy = top + 0.44
    for nm, how, k in G1:
        text(s, ML, yy, 3.06, 0.26, nm, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, ML, yy + 0.26, 3.06, 0.22, how, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 3.30, yy + 0.06, 1.90, 0.26, k, 12, TEAL if k != "0" else GREY2, SANS, True,
             PP_ALIGN.RIGHT, 1.0, check=False)
        yy += 0.76
    # 右组
    RX = ML + 6.30
    text(s, RX, top, 5.50, 0.26, "只靠 OCR / 没开视觉", 11.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.2, check=False)
    hr(s, RX, top + 0.30, 5.50, 2.2, ORANGE)
    G2 = [("Seed 2.1 Turbo · Trae", "自述“通过 OCR 识别图纸内容”", "0"),
          ("HY4 Preview · WorkBuddy", "自述无法看图，只读元数据", "0")]
    yy = top + 0.44
    for nm, how, k in G2:
        text(s, RX, yy, 3.60, 0.26, nm, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, RX, yy + 0.26, 3.70, 0.22, how, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, RX + 3.90, yy + 0.06, 1.60, 0.26, k, 12, ORANGE, SANS, True, PP_ALIGN.RIGHT, 1.0,
             check=False)
        yy += 0.76
    text(s, RX + 3.90, top + 2.00, 1.60, 0.24, "命中数", 9.5, FAINT, SANS, False, PP_ALIGN.RIGHT,
         1.0, check=False)
    # 下半：OCR 误读
    hr(s, ML, 4.28, CW, 1.0, LINE2)
    text(s, ML, 4.42, 5.60, 0.26, "OCR 读错的三个样子", 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    O = [("陈双喜", "陈d喜", "人名，两个字被粘成一个"), ("β（°）", "β（）", "单位丢了，符号还在"),
         ("−9/sinβ", "3-9", "公式被读成一串数字")]
    for i, (a, b, note) in enumerate(O):
        x = ML + i * 4.10
        text(s, x, 4.74, 1.60, 0.28, a, 13, GREY, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, 5.04, 1.30, 0.26, "→", 12, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.34, 5.00, 1.60, 0.32, b, 15, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x, 5.36, 3.80, 0.24, note, 10.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    # 右下结论
    box(s, ML + 6.30, 4.42, 5.50, 1.62, fill=CARD, line=LINE)
    text(s, ML + 6.52, 4.54, 5.06, 0.26, "所以，第一句该问的是", 11, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 6.52, 4.82, 5.10, 0.32, "“你看到的是原图，还是识别出来的文字？”", 14.5, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML + 6.52, 5.20, 5.10, 0.72,
         "能看图是前提，但不是保证：Qwen 能看图，仍只命中一项；\nOpus 能看图，也只命中两项。", 11.5, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：十份审查输出的自述读图方式；OCR 误读样本取自对应输出原文")
    return s


def p32(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "假错误说得很具体，但和原图对不上",
             "最容易被骗的地方就在这里：它给了位置、给了数字、给了算式，读起来毫无破绽。")
    top = y + 0.14
    F = [("MiniMax M3", "“α+γ 不闭合，85.3+84.7=170°”", "依据不成立：图上没有这个约束关系"),
         ("Auto fast", "“2#、3# 墩仅有左侧挡块”", "原图第 2、3 页都画了右侧挡块"),
         ("Step 5 Preview", "七条“1a 共重偏大 0.1 kg”", "按 2.986 kg/m 复算是对的"),
         ("DeepSeek Flash", "“缺少 β（°）列”", "原图第 3 页有该列"),
         ("Qwen 3.8 Flash", "“分段之和比公式大 1.0 cm”", "漏了两端 22 cm，实际差 45 cm"),
         ("Seed 2.1 Turbo", "“陈双喜识别为陈d喜”", "是它自己 OCR 读错，不是图错")]
    cw = (CW - 0.46) / 3
    for i, (nm, said, truth) in enumerate(F):
        c, r = i % 3, i // 3
        x = ML + c * (cw + 0.23)
        yy = top + r * 2.24
        text(s, x, yy, cw, 0.24, nm, 9.5, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        hr(s, x, yy + 0.26, cw, 1.0, LINE2)
        text(s, x, yy + 0.36, cw, 0.92, said, 13.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.34,
             check=False)
        text(s, x, yy + 1.34, cw, 0.22, "原图", 9.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0,
             check=False)
        text(s, x, yy + 1.58, cw, 0.54, truth, 11.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.3, check=False)
    box(s, ML, 6.02, CW, 0.70, fill=TEALXL)
    text(s, ML + 0.24, 6.12, CW - 0.48, 0.52,
         "“60 项表格全对得上”只说明表和公式一致，不说明公式和图一致 —— 这类错，只能回原图核。", 14,
         NAVY, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：十份审查输出原文与挡块图纸原图逐条比对")
    return s


def p33(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "先算出一份基准，才判得了四份报告的对错",
             "提示词：“根据提供的工程量清单完成修编报告，其余文件为查阅变更原因参考。”")
    top = y + 0.16
    pic(s, "s033_1.png", ML, top, 4.06, 2.98, mode="fit")
    text(s, ML, top + 3.06, 4.06, 0.52,
         "材料：旧标段报告模板、修编前后清单、设计后服务函与勘误回复函。\n基准由人工复算，不看任何一份 AI 输出。", 10.5,
         FAINT, SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
    RX = ML + 4.42
    vr(s, RX - 0.28, top, 3.20, 1.0, LINE)
    S = [("41,638.9336", "万元", "修编前清单总额（汇总表，不含暂列金额）", GREY2),
         ("41,465.1967", "万元", "修编后清单总额", NAVY),
         ("−173.7369", "万元", "核减金额", ORANGE)]
    yy = top
    for big, unit, cap, col in S:
        text(s, RX, yy, 3.60, 0.50, big, 28, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX + est_w(big, 28, True) / 72 + 0.10, yy + 0.22, 0.90, 0.24, unit, 12, GREY2,
             SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX, yy + 0.54, 5.40, 0.24, cap, 11, GREY, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        yy += 0.92
    hr(s, RX, yy + 0.02, 7.20, 1.0, LINE2)
    text(s, RX, yy + 0.16, 2.00, 0.24, "101", 24, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, RX + 0.76, yy + 0.26, 6.40, 0.24, "个变化子目", 12, GREY, SANS, False, PP_ALIGN.LEFT,
         1.0, check=False)
    text(s, RX, yy + 0.62, 7.20, 0.56,
         "其中 99 项金额变化（核增 41、核减 58），另 2 项只变数量。\n四份报告的条目数分别是 43、24、101、99 —— 范围不同，不能按条数排名。",
         11.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    footer(s, "来源：TJ4 合同段修编前后清单汇总表；人工复算记录")
    return s


def p34(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "金额几乎都对，错在数量、原因和旧模板",
             "四份报告的截图都是它们自己写的那一段。用同一份基准去核，问题各不相同。")
    top = y + 0.12
    R = [("s034_1.png", "Opus 5.5 · Hark", "4 分 · 43 条", "只列了 ≥1 万元的变化，没交代筛选口径", TEAL),
         ("s034_2.png", "Trae Auto", "7 分 · 24 行", "18 行数量错、正负号写反、合同条文被改写", ORANGE),
         ("s034_3.png", "WorkBuddy Auto", "14 分 · 101 条", "总额笔误、承建单位照抄旧标段、原因全是套话", ORANGE),
         ("s034_4.png", "DeepSeek Flash", "34 分 · 99 条", "原因具体，但有的套错了条目", GREY)]
    for i, (fn, nm, meta, prob, col) in enumerate(R):
        yy = top + i * 1.10
        if i:
            hr(s, ML, yy - 0.10, CW, 1.0, LINE)
        frame(s, ML, yy, 3.10, 0.94, LINE, fill=CARD)
        pic(s, fn, ML, yy, 3.10, 0.94, mode="fit")
        text(s, ML + 3.28, yy + 0.02, 3.00, 0.26, nm, 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 3.28, yy + 0.30, 2.80, 0.24, meta, 11, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 6.24, yy + 0.06, CW - 6.24, 0.72, prob, 12.5, C(0x2B, 0x34, 0x45), SANS,
             False, PP_ALIGN.LEFT, 1.35, check=False)
    hr(s, ML, 6.02, CW, 1.2, LINE2)
    text(s, ML, 6.14, 5.90, 0.26, "四份都做到了", 10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    text(s, ML, 6.40, 5.90, 0.24, "总额与核减额都算对；都交出了可打开的文件", 11.5, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML + 6.30, 6.14, 5.80, 0.26, "四份都没做到", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.0, check=False)
    text(s, ML + 6.30, 6.40, 5.80, 0.24, "交代自己的口径；核对模板里的旧标段信息", 11.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：四份报告的原始输出截图；条目数按各自实际列出的行数统计")
    return s


def p35(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "最像样的那一份，反而藏着两处硬伤",
             "格式、语气、结构都像正式报告。正因为像，人才最容易放过它。")
    top = y + 0.12
    text(s, ML, top, 6.00, 0.24, "WorkBuddy Auto：总额写成 41638.89 万元", 11, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    frame(s, ML, top + 0.28, 5.96, 1.02, LINE, fill=CARD)
    pic(s, "s035_2.png", ML, top + 0.28, 5.96, 1.02, mode="fit")
    text(s, ML, top + 1.36, 5.96, 0.24, "↑ 应为 41,638.9336 万元，两位小数是 41638.93", 10.5, ORANGE,
         SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, top + 1.78, 6.00, 0.24, "Opus 5.5：同一句照抄正确", 11, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    frame(s, ML, top + 2.06, 5.96, 0.98, LINE, fill=CARD)
    pic(s, "s035_1.png", ML, top + 2.06, 5.96, 0.98, mode="fit")
    # 右：承建单位对照
    RX = ML + 6.24
    text(s, RX, top, 5.80, 0.24, "承建单位：模板里写的是旧标段", 11, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    rows = [["来源", "写法", "判定"],
            ["旧模板 TJ1", "中铁十二局集团有限公司", "旧标段，须替换"],
            ["Opus 5.5", "广州公路工程集团有限公司", "与中标公告一致"],
            ["WorkBuddy Auto", "中铁十二局集团有限公司", "照抄旧模板"],
            ["DeepSeek Flash", "确定施工单位承建", "中性，可接受"],
            ["Trae Auto", "（未写）", "缺失"]]
    table(s, RX, top + 0.28, 5.86, rows, [1.50, 2.70, 1.66], size=10, head_size=10,
          minrow=0.40, head_h=0.30, rowh=0.40, first_bold=True, pad=0.10, ls=1.15,
          colors=lambda r, c: (ORANGE if (c == 2 and r in (1, 3)) else
                               (TEAL if (c == 2 and r == 2) else
                                (GREY if c in (1, 2) else C(0x2B, 0x34, 0x45)))))
    hr(s, RX, top + 2.68, 5.86, 1.0, LINE2)
    text(s, RX, top + 2.80, 5.86, 0.24, "为什么最容易漏", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.0, check=False)
    WHY = ["格式、语气、标题都像正式报告，读起来就放松了",
           "金额全对的时候，人会把注意力全放在数字上",
           "模板里的旧信息藏在正文中段，不在显眼处"]
    for i, t in enumerate(WHY):
        text(s, RX, top + 3.08 + i * 0.26, 5.86, 0.24, "· " + t, 10.8, GREY, SANS, False,
             PP_ALIGN.LEFT, 1.0, check=False)
    box(s, RX, top + 3.96, 5.86, 0.60, fill=ORANGEL)
    text(s, RX + 0.18, top + 4.04, 5.50, 0.44,
         "对策：拿模板逐句比 —— 旧标段名、单位名、日期，都要换。", 11.5, C(0x7A, 0x3A, 0x22), SANS,
         True, PP_ALIGN.LEFT, 1.3)
    footer(s, "来源：两份报告原文截图；中标公告（TJ4 合同段）")
    return s


def p36(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一条变更，四份报告给出四种原因",
             "挑两条金额较大的核减子目。原因写得好不好，一看能不能回到函件，二看数量对不对。")
    top = y + 0.14
    CX = [ML, ML + 1.86, ML + 5.86, ML + 9.62]
    HD = ["", "透水管  −71.8663 万元", "支座  −23.6094 万元"]
    text(s, CX[1], top, 3.90, 0.26, HD[1], 11, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6,
         check=False)
    text(s, CX[2], top, 3.66, 0.26, HD[2], 11, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6,
         check=False)
    text(s, CX[3], top, 3.08, 0.26, "判定", 11, GREY2, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6,
         check=False)
    hr(s, ML, top + 0.30, CW, 1.4, LINE2)
    D = [("Opus 5.5", "纵向填挖交界渗沟间距由每 4m 调整为每 8m",
          "乳子石大桥支座汇总表与细表类型及数量不一致", "两项都有据，对得上勘误表", TEAL),
         ("DeepSeek Flash", "环向渗沟透水管按每道 1 根计列", "由 120 个调整为 60 个",
          "支座对；透水管套错条目", ORANGE),
         ("Trae Auto", "根据设计变更调整排水工程数量", "桥梁支座数量调整",
          "两项数量都写错", ORANGE),
         ("WorkBuddy Auto", "依据设计后服务函及勘误回复函核对修改清单数量", "（同左，101 条同一句）",
          "等于没写原因", ORANGE)]
    yy = top + 0.40
    for i, (nm, a, b, j, col) in enumerate(D):
        if i:
            hr(s, ML, yy - 0.10, CW, 1.0, LINE)
        text(s, ML, yy + 0.02, 1.80, 0.26, nm, 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, CX[1], yy, 3.88, 0.76, a, 11.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT,
             1.34, check=False)
        text(s, CX[2], yy, 3.64, 0.76, b, 11.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT,
             1.34, check=False)
        text(s, CX[3], yy, 3.10, 0.76, j, 11, col, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
        yy += 0.96
    box(s, ML, 5.86, CW, 0.78, fill=TEALXL)
    text(s, ML + 0.24, 5.97, CW - 0.48, 0.58,
         "“原因写得好”的判断标准只有一条：能不能顺着这句话，翻到那份函件的第几条。", 14, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：设计后服务函、勘误回复函；四份报告对应条目的原文")
    return s


def p37(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "写得具体，不等于有出处",
             "子目 203-1-2-2-1：395.9711 → 573.1093 万元，增加 177.1382 万元。看四份报告怎么解释这一条。")
    top = y + 0.12
    # 三个数字
    N = [("395.9711", "原合同（万元）", GREY2), ("+177.1382", "核增（万元）", ORANGE),
         ("573.1093", "修编后（万元）", NAVY)]
    for i, (v, cap, col) in enumerate(N):
        x = ML + i * 2.04
        text(s, x, top, 1.96, 0.44, v, 25, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, top + 0.50, 1.96, 0.24, cap, 10.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
    # 参考函
    RX = 6.98
    text(s, RX, top - 0.02, X2 - RX, 0.24, "参考函原文（设计后服务函 第 1 页第 8 项）", 10, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    pic(s, "s037_1.png", RX, top + 0.26, X2 - RX, 0.66, mode="fit")
    hr(s, ML, top + 0.98, CW, 1.0, LINE2)
    # 四行摘录
    D = [("Opus 5.5", "“下路床换填材料修改为石渣，土石方调配调整，利用石方增加，审核予以调整。”",
          "对得上原文", TEAL),
         ("DeepSeek Flash", "“利用方增加石渣利用量，下路床换填材料调整为石渣。”", "对得上原文", TEAL),
         ("Trae Auto", "“石方利用量增加，根据土石方调配优化调整。”", "只说了一半", ORANGE),
         ("WorkBuddy Auto", "“依据设计后服务函及勘误回复函核对修改清单数量。”", "泛化，等于没写", ORANGE)]
    yy = top + 1.14
    for i, (nm, txt, j, col) in enumerate(D):
        if i:
            hr(s, ML, yy - 0.08, CW, 1.0, LINE)
        text(s, ML, yy + 0.02, 1.84, 0.26, nm, 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 1.90, yy, 8.42, 0.44, txt, 11.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.3, check=False)
        text(s, ML + 10.46, yy + 0.02, 1.66, 0.26, j, 10.5, col, SANS, True, PP_ALIGN.RIGHT, 1.0,
             check=False)
        yy += 0.66
    # 检查顺序
    hr(s, ML, 5.90, CW, 1.0, LINE2)
    text(s, ML, 6.02, 1.30, 0.26, "检查顺序", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    ST = ["金额", "数量", "原因能否回到源文件", "模板有没有改干净"]
    for i, t in enumerate(ST):
        x = ML + 1.30 + i * 2.72
        text(s, x, 6.00, 0.34, 0.24, f"{i+1}", 10.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.28, 6.00, 2.34, 0.26, t, 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        if i < 3:
            text(s, x + 2.36, 6.00, 0.30, 0.24, "›", 12, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
    footer(s, "来源：设计后服务函第 1 页第 8 项；四份报告的对应条目原文")
    return s


def p38(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "合同条文不让改写，只能原样引用",
             "同一段 15.4 变更估价原则，四份报告出现了三种处理方式。")
    top = y + 0.14
    text(s, ML, top, 5.30, 0.24, "三份逐字一致：Opus / WorkBuddy / DeepSeek", 10.5, TEAL, SANS,
         True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    frame(s, ML, top + 0.26, 5.30, 2.10, LINE, fill=CARD)
    pic(s, "s038_2.png", ML, top + 0.26, 5.30, 2.10, mode="fit")
    text(s, ML + 5.60, top, 6.60, 0.24, "Trae：压成三条，条件被删掉", 10.5, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    frame(s, ML + 5.60, top + 0.26, 4.10, 2.10, LINE, fill=CARD)
    pic(s, "s038_1.png", ML + 5.60, top + 0.26, 4.10, 2.10, mode="fit")
    text(s, ML + 9.86, top + 0.30, 2.86, 1.90,
         "删掉的是条件，不是废话：\n15.4.1、15.4.3、15.4.5，\n“经发包人同意”，\n以及最后两段。",
         11, GREY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    # 三栏
    hr(s, ML, top + 2.56, CW, 1.0, LINE2)
    B = [("Trae 删掉的", "把“有相同或相似子目单价，经发承包双方协商一致可直接套用”压成“可直接套用” —— 前提没了。",
          ORANGE),
         ("Opus 多做的", "按 15.4.2 取 2 标 3 标平均值，算出新增单价 389.77 元/m² —— 多做了一步，但方法在条文里有。",
          TEAL),
         ("三份都没核的", "“2025 年 04 月”是模板里的月份，四份报告都照抄了 —— 这是模板残留，不是条文。",
          GREY2)]
    cw = (CW - 0.52) / 3
    for i, (t, b, col) in enumerate(B):
        x = ML + i * (cw + 0.26)
        box(s, x, top + 2.68, cw, 0.05, fill=col)
        text(s, x, top + 2.82, cw, 0.26, t, 12, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        text(s, x, top + 3.14, cw, 0.90, b, 11.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.38, check=False)
    footer(s, "来源：TJ4 合同段施工合同 15.4 条原文；四份报告对应段落")
    return s


def build(prs, start):
    div(prs, start)
    fn = [p29, p30, p31, p32, p33, p34, p35, p36, p37, p38]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
