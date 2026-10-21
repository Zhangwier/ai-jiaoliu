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


