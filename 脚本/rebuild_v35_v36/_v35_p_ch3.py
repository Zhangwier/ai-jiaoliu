# -*- coding: utf-8 -*-
"""第 03 章 两个真实案例"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

CH = "03   两个真实案例"


# ---------------------------------------------------------------- 章节页（浅色 + 右侧图纸）
def div03(prs, pno):
    s = slide(prs)
    pic(s, "s029_1.png", 6.55, 0, 6.78, 7.5, mode="fill")
    vr(s, 6.55, 0, 7.5, 0.75, RULE)
    header(s, "章节", pno)
    text(s, ML, 1.28, 4.0, 0.26, "PART  03", 9.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, cs=2.6, check=False)
    text(s, ML, 1.66, 5.35, 1.80, "两件事，\n一百多条可核对的结论。", 36, INK, SERIF, False,
         PP_ALIGN.LEFT, 1.34, check=False)
    text(s, ML, 3.62, 4.95, 0.30, "同一张图、同一份清单，交给不同的配置，结果差得很远。", 12.5, BODY,
         SANS, False, PP_ALIGN.LEFT, 1.45)
    hr(s, ML, 4.08, 4.95, 0.75, RULE2)

    yy = 4.28
    cases = [("案例一", "挡块图纸审查", "10 个配置 · 2 个检查点 · 30 条假错误"),
             ("案例二", "修编报告", "4 个配置 · 101 个变化子目 · 1 份复算基准")]
    for k, nm, sub in cases:
        text(s, ML, yy, 1.0, 0.24, k, 9.5, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
        text(s, ML + 0.98, yy - 0.05, 3.9, 0.30, nm, 16, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, ML + 0.98, yy + 0.30, 3.9, 0.24, sub, 9.8, MUT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        yy += 0.98

    text(s, ML, 6.34, 4.95, 0.50,
         "这里不放结论，只摆可核对的东西：谁做的、用了多久、说对了哪一条、错在哪一行。",
         9.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.5, check=False)
    text(s, 11.94, 7.10, 1.0, 0.2, "", 8, FAINT, SANS, False, PP_ALIGN.RIGHT, 1.0, check=False)
    return s


# ---------------------------------------------------------------- 29 案例一 设定
def p29(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一张挡块图，十个配置，只有三个找到了这两处",
             "提示词只有一句：「审查这个有没有问题结果输出 excel 给我」。所以先替它把“要看什么”定成两个可核对的点。")

    # 左：整张图纸
    pic(s, "s029_1.png", ML, 2.12, 6.30, 3.13, mode="fit")
    hr(s, ML, 5.31, 6.30, 0.75, RULE)
    text(s, ML, 5.44, 6.30, 0.24, "原图：挡块构造图（含 1a 大样、3a 大样与 A–A 剖面）", 8.6, MUT,
         SANS, False, PP_ALIGN.LEFT, 1.2, check=False)

    # 右：两个检查点
    x = 7.55
    pic(s, "s029_2.png", x, 2.12, 2.42, 1.88, mode="fit")
    pic(s, "s029_3.png", x + 2.62, 2.12, 1.58, 1.88, mode="fit")
    text(s, x, 4.06, 2.42, 0.22, "1a 大样：总长公式与分段", 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.1, check=False)
    text(s, x + 2.62, 4.06, 1.58, 0.22, "3a 引线与 A–A", 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.1, check=False)

    hr(s, x, 4.44, 4.80, 0.75, RULE2)
    blocks = [("K1", "分段与总长对不上", "1a 大样各段加起来，比图上印的总长公式多 45 cm。"),
              ("K2", "符号不一致", "3a 图上竖标写 sin α，同一处引线公式写 sin β。")]
    yy = 4.58
    for k, nm, ds_ in blocks:
        text(s, x, yy, 0.52, 0.24, k, 11, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.60, yy - 0.03, 4.2, 0.24, nm, 11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x + 0.60, yy + 0.26, 4.2, 0.24, ds_, 9.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.25, check=False)
        yy += 0.72

    hr(s, ML, 5.86, CW, 0.75, RULE)
    text(s, ML, 5.98, 1.1, 0.22, "口径", 9.5, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    text(s, ML + 1.10, 5.96, 10.4, 0.52,
         "单位 mm/cm 这类问题经工程复核不算问题，不计入；分母是两个检查点，所以叫“命中几个”，不叫准确率。另有 21 条待核未计。",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
    footer(s, "来源：本地图纸审查记录；原图与十份输出均为本次实测材料")


# ---------------------------------------------------------------- 30 结果矩阵
def p30(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "三个配置命中，假错误却累计三十条",
             "按命中数排序。深色段是有效条目，红色段是假错误 —— 列得多，不等于找得准。")

    DATA = [("Opus 5.5 · Hark", "3:03", 1, 1, 2, 2, 11, "看图（自述人工识读）"),
            ("GPT 6.1 Sol high · Codex", "14:34", 1, 0, 4, 1, 8, "OCR 定位 + 回看原图"),
            ("Qwen 3.8 Flash · Qoder", "43:09", 0, 1, 2, 5, 24, "渲染成图、逐格读"),
            ("Auto fast · WorkBuddy", "2:38", 0, 0, 1, 2, 7, "未说明"),
            ("DeepSeek Flash · WorkBuddy", "32:05", 0, 0, 1, 2, 19, "自述“图面识读”"),
            ("Step 5 Preview · WorkBuddy", "29:40", 0, 0, 1, 8, 26, "未说明"),
            ("MiniMax M3 · WorkBuddy", "12:27", 0, 0, 0, 5, 10, "未说明"),
            ("Seed 2.1 Turbo · Trae", "12:47", 0, 0, 0, 3, 10, "OCR"),
            ("GLM 5.3 Flash · Trae", "1:24:00", 0, 0, 0, 0, 1, "渲染后放大看"),
            ("HY4 Preview · WorkBuddy", "3:55", 0, 0, 0, 2, 2, "视觉未启用")]

    XN, XT = ML, 3.46
    XK1, XK2, XO = 4.38, 4.80, 5.28
    XB, XBW = 5.66, 2.30
    XC, XF, XR = 8.06, 8.54, 9.20

    hd = [("配置", XN, 2.50, PP_ALIGN.LEFT), ("耗时", XT, 0.60, PP_ALIGN.RIGHT),
          ("K1", XK1 - 0.20, 0.40, PP_ALIGN.CENTER), ("K2", XK2 - 0.20, 0.40, PP_ALIGN.CENTER),
          ("其他", XO - 0.24, 0.48, PP_ALIGN.CENTER),
          ("条数（深色 = 有效，红色 = 假错误）", XB, 2.36, PP_ALIGN.LEFT),
          ("共", XC, 0.40, PP_ALIGN.RIGHT), ("假错", XF, 0.46, PP_ALIGN.RIGHT),
          ("怎么读的图", XR, 3.20, PP_ALIGN.LEFT)]
    for t, xx, ww, al in hd:
        text(s, xx, 1.96, ww, 0.22, t, 8.2, MUT, SANS, True, al, 1.0, cs=0.5, check=False)
    hr(s, ML, 2.20, CW, 1.1, RULE2)

    yy = 2.32
    for nm, tm, k1, k2, ot, fk, tot, rd in DATA:
        text(s, XN, yy + 0.02, 2.50, 0.24, nm, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, XT, yy + 0.03, 0.60, 0.22, tm, 9.2, MUT, MONO, False, PP_ALIGN.RIGHT, 1.0, check=False)
        for kx, kv in ((XK1, k1), (XK2, k2)):
            if kv:
                circ(s, kx, yy + 0.14, 0.135, RED)
            else:
                circ(s, kx, yy + 0.14, 0.115, None, FAINT, 0.9)
        text(s, XO - 0.30, yy + 0.03, 0.60, 0.22, str(ot) if ot else "—", 9.2, BODY, MONO, False,
             PP_ALIGN.CENTER, 1.0, check=False)
        # 堆叠条
        vr(s, XB - 0.07, yy - 0.01, 0.30, 0.6, C(0xE6, 0xE1, 0xD5))
        wv = (tot - fk) / 26.0 * XBW
        wf = fk / 26.0 * XBW
        if wv > 0.01:
            box(s, XB, yy + 0.055, wv, 0.14, fill=INK)
        if wf > 0.01:
            box(s, XB + wv, yy + 0.055, wf, 0.14, fill=RED)
        text(s, XC, yy + 0.03, 0.40, 0.22, str(tot), 9.4, INK, MONO, True, PP_ALIGN.RIGHT, 1.0, check=False)
        text(s, XF, yy + 0.03, 0.46, 0.22, str(fk) if fk else "0", 9.4, RED if fk else FAINT, MONO, True,
             PP_ALIGN.RIGHT, 1.0, check=False)
        text(s, XR, yy + 0.03, 3.20, 0.22, rd, 9.0, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, ML, yy + 0.325, CW, 0.55, C(0xE3, 0xDE, 0xD2))
        yy += 0.375

    # 图例 + 结论
    box(s, ML, 6.22, 0.16, 0.09, fill=INK)
    text(s, ML + 0.24, 6.15, 1.2, 0.22, "有效条目", 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    box(s, ML + 1.32, 6.22, 0.16, 0.09, fill=RED)
    text(s, ML + 1.56, 6.15, 1.8, 0.22, "假错误（与原图对不上）", 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 3.72, 6.15, 7.80, 0.24,
         "Step 5 列了 26 条、其中 8 条是假的；Opus 只列 11 条，两个检查点全中。",
         9.6, RED, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：本地图纸审查记录（十份输出逐条复核；另有 21 条待核未计）")


# ---------------------------------------------------------------- 31 读图方式
def p31(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "只靠 OCR 的两组，一条都没命中",
             "十份输出的差别，一半在“图到底进没进模型”。左边四组真的读到了图，右边两组只读到了文字。")

    def group(x, w, title, sub, rows, col):
        text(s, x, 1.92, w, 0.24, title, 9.6, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
        text(s, x, 2.18, w, 0.22, sub, 8.8, MUT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        hr(s, x, 2.44, w, 1.6, col)
        yy = 2.58
        for nm, how, hit in rows:
            text(s, x, yy, w - 0.70, 0.24, nm, 10.6, INK, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
            text(s, x, yy + 0.25, w - 0.70, 0.34, how, 9.0, MUT, SANS, False, PP_ALIGN.LEFT, 1.25)
            text(s, x + w - 0.70, yy - 0.01, 0.70, 0.26, hit, 11.5, RED if hit != "0" else FAINT, MONO, True,
                 PP_ALIGN.RIGHT, 1.0, check=False)
            yy += 0.74

    group(ML, 5.05, "真的看到了图", "命中标记：K1 / K2；0 表示两个都没命中",
          [("Opus 5.5 · Hark", "自述人工识读整张图", "K1 K2"),
           ("GPT 6.1 Sol · Codex", "OCR 定位，再回看原图核对", "K1"),
           ("Qwen 3.8 Flash · Qoder", "文字提取无效，渲染成图逐格读", "K2"),
           ("GLM 5.3 Flash · Trae", "渲染后放大读图，看了 1 小时 24 分", "0")], GRN)
    vr(s, 6.28, 1.92, 3.34, 0.75, RULE)
    group(6.72, 5.05, "只靠 OCR / 没开视觉", "连图里的符号都读成了别的字",
          [("Seed 2.1 Turbo · Trae", "自述“通过 OCR 识别图纸内容”", "0"),
           ("HY4 Preview · WorkBuddy", "自述无法看图，只读文件元数据", "0")], RULE2)

    # 误读三例
    hr(s, ML, 5.52, CW, 0.75, RULE)
    text(s, ML, 5.62, 3.2, 0.22, "OCR 读错的三个字，被当成图上的错", 9.6, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    rows = [("陈双喜", "陈d喜", "被当成“图上写错人名”报上来"),
            ("β（°）", "β（）", "被判“缺少单位”"),
            ("−9 / sin β", "3-9", "被判“疑为 1-9 写错”")]
    for i, (a, b, c) in enumerate(rows):
        text(s, ML, 5.94 + i * 0.34, 1.7, 0.24, a, 10.5, INK, MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, 2.55, 5.94 + i * 0.34, 2.0, 0.24, "→  " + b, 10.5, RED, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, 4.66, 5.94 + i * 0.34, 3.30, 0.24, c, 9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    vr(s, 8.16, 5.62, 0.98, 0.75, RULE)
    text(s, 8.36, 5.62, 4.08, 1.00,
         "能看图是前提，但不是保证：GLM 看了 1 小时 24 分，仍判“无实质错误”。\n用之前先确认——图是真的进了模型，不是只传了个文件名。",
         9.4, GRN, SANS, False, PP_ALIGN.LEFT, 1.48)
    footer(s, "来源：十份输出中的自述与原图比对（OCR 误读三例均为输出原文）")


# ---------------------------------------------------------------- 32 假错误六例
def p32(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "假错误说得很具体，但和原图对不上",
             "六条被报上来的“问题”，逐条回到原图都站不住 —— 越具体的描述，越容易被当成真的。")
    ITEMS = [("MiniMax M3", "“α + γ 不闭合：85.3 + 84.7 = 170°”", "这个依据本身不成立，它据此判“不通过”。"),
             ("Auto fast", "“2#、3# 墩仅有左侧挡块”", "第 2、3 页原图都画了右侧挡块。"),
             ("Step 5 Preview", "7 条“1a 共重偏大 0.1 kg”", "按 2.986 kg/m 复算，原值是对的。"),
             ("DeepSeek Flash", "“缺少 β（°）列”", "第 3 页原图有 β（°）列，是 OCR 丢了。"),
             ("Qwen 3.8 Flash", "“分段之和比公式大 1.0 cm”", "漏算两端 22 cm，实际差 45 cm。"),
             ("Seed 2.1 Turbo", "“陈双喜识别为陈d喜”", "OCR 自己读错的字，被当成图上的错。")]
    cw = (CW - 0.56 * 2) / 3
    for i, (nm, said, truth) in enumerate(ITEMS):
        cx = ML + (i % 3) * (cw + 0.56)
        cy = 2.02 + (i // 3) * 2.10
        text(s, cx, cy, cw, 0.22, nm, 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        text(s, cx, cy + 0.26, cw, 0.62, said, 11.5, RED, SERIF, False, PP_ALIGN.LEFT, 1.42)
        hr(s, cx, cy + 1.00, cw, 0.6, RULE)
        text(s, cx, cy + 1.10, 0.5, 0.20, "原图", 8.2, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cx + 0.44, cy + 1.08, cw - 0.44, 0.56, truth, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.42)
    hr(s, ML, 6.16, CW, 0.75, RULE)
    text(s, ML, 6.28, CW, 0.30,
         "“60 项表格全对得上”，只说明表和公式一致，不说明公式和图一致。",
         13, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：十份输出逐条复核；原图为挡块构造图")


# ---------------------------------------------------------------- 33 案例二 设定
def p33(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "先算出一份基准，才判得了四份报告的对错",
             "提示词：「根据提供的工程量清单完成修编报告，其余文件为查阅变更原因参考」。判断之前，先把汇总表复算一遍。")

    pic(s, "s033_1.png", ML, 2.10, 4.06, 3.05, mode="fit")
    hr(s, ML, 5.26, 4.06, 0.75, RULE)
    text(s, ML, 5.38, 4.06, 0.62,
         "材料：旧标段报告模板（TJ1）、修编前后清单、设计后服务函与勘误回复函。",
         9.0, MUT, SANS, False, PP_ALIGN.LEFT, 1.4)

    x = 5.30
    text(s, x, 1.98, 3.2, 0.22, "复算基准（汇总表，不含暂列金额）", 9.2, GRN, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.0, check=False)
    nums = [("修编前", "41,638.9336", "万元"), ("修编后", "41,465.1967", "万元"), ("核减", "−173.7369", "万元")]
    yy = 2.32
    for i, (lab, v, u) in enumerate(nums):
        text(s, x, yy, 1.2, 0.22, lab, 9.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 1.32, yy - 0.14, 4.0, 0.46, v, 27, INK if i < 2 else RED, MONO, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 5.38 - 0.62, yy + 0.10, 0.62, 0.22, u, 8.8, MUT, SANS, False, PP_ALIGN.RIGHT, 1.0, check=False)
        hr(s, x, yy + 0.48, 5.38, 0.55, C(0xE3, 0xDE, 0xD2))
        yy += 0.78

    hr(s, x, 4.96, 5.38, 0.75, RULE)
    text(s, x, 5.08, 5.38, 0.24, "101 个变化子目", 10.5, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    rich(s, x, 5.36, 5.38, 1.30, [
        [("99 项", GRN, True, SANS, 12), (" 金额有变化：核增 ", BODY, False, SANS, 9.8), ("41", RED, True, MONO, 9.8),
         (" 项、核减 ", BODY, False, SANS, 9.8), ("58", RED, True, MONO, 9.8), (" 项；另 ", BODY, False, SANS, 9.8),
         ("2", GRN, True, MONO, 9.8), (" 项只变数量、金额不变。", BODY, False, SANS, 9.8)],
        {"runs": [("有了能复算的基准，四份报告的金额、数量、原因才有同一把尺子；", MUT, False, SANS, 9.4)],
         "sp": 9},
        {"runs": [("没有基准，就只能凭哪份读起来更顺。", INK, True, SANS, 9.4)]}], ls=1.5)
    footer(s, "来源：修编前后工程量清单、汇总表与复算底稿；四份修编报告原文")


# ---------------------------------------------------------------- 34 四份报告
def p34(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "金额几乎都对，错在数量、原因和旧模板",
             "各份列的范围不同，不能按条目多少排名 —— 按“哪一类错”看，差别才清楚。")
    R = [("Opus · Hark", "s034_1.png", "4 分 · 43 条", "只列 ≥1 万元的条目，没交代筛选"),
         ("Trae Auto", "s034_2.png", "7 分 · 24 行", "18 行数量错；正负号写反；合同条文被改写"),
         ("WorkBuddy Auto", "s034_3.png", "14 分 · 101 条", "总额笔误；承建单位照抄旧标段；原因全是套话"),
         ("DeepSeek Flash", "s034_4.png", "34 分 · 99 条", "原因写得具体，但有的套错了条目")]
    yy = 2.00
    for nm, im, meta, prob in R:
        pic(s, im, ML, yy, 3.34, 1.01, mode="fit")
        text(s, ML + 3.52, yy + 0.02, 2.95, 0.26, nm, 12, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, ML + 3.52, yy + 0.30, 2.95, 0.22, meta, 9.2, RED, MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 6.62, yy + 0.06, 4.90, 0.60, prob, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.42)
        if yy > 2.10:
            hr(s, ML, yy - 0.10, CW, 0.55, C(0xE3, 0xDE, 0xD2))
        yy += 1.12
    hr(s, ML, 6.42, CW, 0.75, RULE)
    text(s, ML, 6.54, 6.0, 0.24, "四份都做到的：金额算得基本对。", 9.8, GRN, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 6.20, 6.54, 5.33, 0.24, "四份都没做到的：把“为什么变”写回到函件和勘误表。", 9.8, RED, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    footer(s, "来源：四份修编报告原文（截图为各份对应段落）；复算底稿")


# ---------------------------------------------------------------- 35 两处硬伤
def p35(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "最像样的那一份，反而藏着两处硬伤",
             "WorkBuddy Auto：101 条、格式最工整、读起来最像正式报告 —— 错的地方也最不容易被看见。")

    text(s, ML, 1.96, 5.0, 0.22, "一 · 总额写成 41638.89", 9.4, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    pic(s, "s035_2.png", ML, 2.24, 3.50, 1.27, mode="fit")
    text(s, ML, 3.60, 3.50, 0.22, "WorkBuddy Auto 报告原文", 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    pic(s, "s035_1.png", ML + 3.72, 2.24, 3.50, 1.27, mode="fit")
    text(s, ML + 3.72, 3.60, 3.50, 0.22, "Opus 报告原文（41,638.93 / 173.74 / 41,465.20）", 8.4, MUT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, 3.92, 7.22, 0.24,
         "两位小数之内的笔误，格式和语气全对，肉眼一扫就过去了。", 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.2,
         check=False)

    text(s, ML, 4.44, 6.0, 0.22, "二 · 承建单位照抄旧标段", 9.4, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    rows = [("旧模板 TJ1", "中铁十二局集团有限公司", "旧标段的单位", MUT),
            ("Opus", "广州公路工程集团有限公司", "✓ 与中标公告一致", GRN),
            ("WorkBuddy Auto", "中铁十二局集团有限公司", "✕ 照抄旧模板", RED),
            ("DeepSeek Flash", "“确定施工单位承建”", "✓ 中性写法", GRN),
            ("Trae Auto", "没写", "— 缺项", FAINT)]
    yy = 4.72
    xs = [ML, ML + 2.28, ML + 6.10]
    for i, (a, b, c, col) in enumerate(rows):
        text(s, xs[0], yy, 2.1, 0.22, a, 9.6, MUT if i == 0 else INK, SANS, i != 0, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, xs[1], yy, 3.8, 0.22, b, 9.6, INK, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, xs[2], yy, 3.0, 0.22, c, 9.4, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        hr(s, ML, yy + 0.28, 7.22, 0.5, C(0xE3, 0xDE, 0xD2))
        yy += 0.35

    x = 8.62
    vr(s, x - 0.30, 1.96, 4.44, 0.75, RULE)
    text(s, x, 1.96, 3.7, 0.22, "为什么最容易漏", 9.4, INK, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
    items = [("格式、语气都像正式报告", "读起来很顺，眼睛不会停在数字上"),
             ("金额全对时人会放松", "前面对了，后面就默认也对"),
             ("旧模板的痕迹不会自己消失", "标段名、单位、日期、月份都得逐个换")]
    yy = 2.30
    for a, b in items:
        text(s, x, yy, 3.72, 0.24, a, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, x, yy + 0.28, 3.72, 0.50, b, 9.2, MUT, SANS, False, PP_ALIGN.LEFT, 1.35)
        yy += 0.90
    hr(s, x, 5.10, 3.72, 0.75, RULE)
    text(s, x, 5.22, 3.72, 0.92,
         "对策：拿模板逐句比；材料外补进来的信息，一律注明出处。",
         9.6, GRN, SANS, True, PP_ALIGN.LEFT, 1.5)
    footer(s, "来源：四份修编报告原文；变更修编报告模板（TJ1）；TJ4 中标结果公告（2026-03-23）")


# ---------------------------------------------------------------- 36 原因对比
def p36(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "同一条变更，四份报告给出四种原因",
             "两个子目的金额四份都写对了；把原因摆在一起，差别立刻出来。")
    xs = [ML, ML + 1.62, ML + 7.06]
    text(s, xs[0], 1.96, 1.4, 0.22, "报告", 8.8, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    text(s, xs[1], 1.96, 5.2, 0.22, "透水管　−71.8663 万元", 9.2, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, xs[2], 1.96, 4.4, 0.22, "支座　−23.6094 万元", 9.2, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 2.22, CW, 1.1, RULE2)
    R = [("Opus", "“纵向填挖交界渗沟间距由每 4 m…调整为每 8 m”", "“乳子石大桥支座汇总表与细表类型及数量不一致”"),
         ("DeepSeek", "“环向渗沟透水管按每道 1 根计列”", "“由 120 个调整为 60 个”"),
         ("Trae", "“根据设计变更调整排水工程数量”", "“桥梁支座数量调整”"),
         ("WorkBuddy", "“依据设计后服务函及勘误回复函核对修改清单数量”", "同左（101 条都是这一句）")]
    yy = 2.36
    for nm, a, b in R:
        text(s, xs[0], yy + 0.04, 1.5, 0.24, nm, 10.5, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, xs[1], yy, 5.30, 0.72, a, 10.2, BODY, SERIF, False, PP_ALIGN.LEFT, 1.4)
        text(s, xs[2], yy, 4.40, 0.72, b, 10.2, BODY, SERIF, False, PP_ALIGN.LEFT, 1.4)
        hr(s, ML, yy + 0.82, CW, 0.55, C(0xE3, 0xDE, 0xD2))
        yy += 0.92
    # 判断
    hr(s, ML, 6.06, CW, 0.75, RULE)
    verd = [("Opus", "有据，对得上勘误表", GRN), ("DeepSeek", "透水管套错了条目", RED),
            ("Trae", "两项数量都写错", RED), ("WorkBuddy", "一句套话，等于没写", RED)]
    xx = [ML, ML + 3.00, ML + 6.00, ML + 8.60]
    for i, (nm, v, col) in enumerate(verd):
        text(s, xx[i], 6.18, 3.0, 0.20, nm, 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, xx[i], 6.40, 3.0, 0.24, v, 10, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    footer(s, "来源：四份修编报告原文；勘误回复函（序号 49、70、73、81）")


# ---------------------------------------------------------------- 37 单条核对
def p37(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "写得具体，不等于有出处",
             "清单 203-1-2-2-1：395.9711 万元增至 573.1093 万元，增加 177.1382 万元 —— 四份报告对同一条变化的说法。")
    nums = [("395.9711", "原合同（万元）", INK), ("+177.1382", "核增（万元）", RED), ("573.1093", "修编后（万元）", INK)]
    for i, (v, lab, col) in enumerate(nums):
        xx = ML + i * 2.72
        text(s, xx, 1.94, 2.6, 0.40, v, 20, col, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, xx, 2.34, 2.6, 0.20, lab, 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)

    hr(s, ML, 2.68, CW, 0.75, RULE)
    xs = [ML, ML + 1.62, ML + 8.02]
    text(s, xs[0], 2.78, 1.4, 0.22, "配置", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    text(s, xs[1], 2.78, 6.2, 0.22, "报告里的原因", 8.6, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    text(s, xs[2], 2.78, 4.0, 0.22, "怎么判", 8.6, RED, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.6, check=False)
    R = [("Opus", "“下路床换填材料修改为石渣……利用石方增加”", "具体，仍需核来源"),
         ("DeepSeek", "“利用方增加石渣利用量……调整为石渣”", "仍需核对应关系"),
         ("Trae", "“石方利用量增加，根据土石方调配优化调整”", "方向清楚，仍需依据"),
         ("WorkBuddy", "“依据设计后服务函及勘误回复函核对修改清单数量”", "泛化，没解释为何增加")]
    yy = 3.10
    for nm, txt, v in R:
        text(s, xs[0], yy + 0.02, 1.5, 0.24, nm, 10.2, INK, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, xs[1], yy, 6.30, 0.44, txt, 9.8, BODY, SERIF, False, PP_ALIGN.LEFT, 1.4)
        text(s, xs[2], yy + 0.02, 4.0, 0.44, v, 9.4, RED if "泛化" in v else GRN, SANS, False, PP_ALIGN.LEFT, 1.3)
        hr(s, ML, yy + 0.54, CW, 0.55, C(0xE3, 0xDE, 0xD2))
        yy += 0.62

    text(s, ML, 5.72, 4.4, 0.22, "参考函原文（设计后服务函 第 1 页第 8 项）", 8.8, GRN, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    pic(s, "s037_1.png", ML, 5.98, 6.60, 0.72, mode="fit")
    text(s, 7.78, 5.72, 4.56, 1.00,
         "函件支持“下路床换填材料为石渣”这条原因成立，但完整金额归因还要结合数量表 —— 写得详细，也得对上条目。",
         9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
    footer(s, "来源：四份修编报告与设计后服务函原件")


# ---------------------------------------------------------------- 38 合同条文
def p38(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "合同条文不让改写，只能原样引用",
             "专用条款 15.4 变更估价原则（新增单价怎么定）：三份逐字照抄，一份被压成了三条。")
    pic(s, "s038_1.png", ML, 1.94, 5.20, 2.07, mode="fit")
    text(s, ML, 4.08, 5.20, 0.22, "Opus / WorkBuddy / DeepSeek：逐字一致", 8.6, GRN, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    pic(s, "s038_2.png", ML + 5.68, 1.94, 4.06, 2.07, mode="fit")
    text(s, ML + 5.68, 4.08, 4.20, 0.22, "Trae：压成三条，删掉了条件", 8.6, RED, SANS, True, PP_ALIGN.LEFT, 1.0,
         check=False)
    text(s, ML + 10.00, 1.94, 2.42, 2.20,
         "合同条文是有法律效力的文字。压成三条读起来更清楚，但删掉的正是“什么情况下才适用”。",
         9.4, BODY, SANS, False, PP_ALIGN.LEFT, 1.5)

    hr(s, ML, 4.52, CW, 1.1, RULE2)
    cols = [("Trae 删掉的", RED, "15.4.1、15.4.3、15.4.5 等条款，“经发包人同意”，以及最后两段。"),
            ("Opus 多做的", GRN, "写明本次新增单价按 15.4.2 款取 2 标、3 标平均值 —— 389.77 元/m²。"),
            ("三份都没核的", AMB, "“2025 年 04 月”是模板带来的月份，本标段是否适用没有依据。")]
    for i, (h, col, body) in enumerate(cols):
        xx = ML + i * 3.86
        text(s, xx, 4.66, 3.56, 0.22, h, 9.4, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
        hr(s, xx, 4.92, 3.56, 1.4, col)
        text(s, xx, 5.04, 3.56, 0.90, body, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)
    hr(s, ML, 6.16, CW, 0.75, RULE)
    text(s, ML, 6.28, CW, 0.30,
         "要求它原样引用条款，再单独写一段“本次怎么适用”。",
         12.5, INK, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：四份报告第三节；变更修编报告模板（TJ1）")


def build(prs, start):
    div03(prs, start)
    fn = [p29, p30, p31, p32, p33, p34, p35, p36, p37, p38]
    for i, f in enumerate(fn):
        f(prs, start + 1 + i)
