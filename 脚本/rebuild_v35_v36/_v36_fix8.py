# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

i0 = s.index('def p13(prs, pno):')
i1 = s.index('def p14(prs, pno):')

new = '''def p13(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "编造不是故障，是“挑最像的下一个字”的副作用",
             "它没有“查证”这个动作。材料里没有，它就按最像的样子造一个 —— 语感越好，越像真的。")
    HDR = [("时间", ML, 1.30), ("谁", ML + 1.36, 3.60), ("编了什么", ML + 5.02, 3.10),
           ("代价", ML + 8.24, 2.90), ("出处", X2 - 0.90, 0.90)]
    for t, x, w in HDR:
        al = PP_ALIGN.RIGHT if t == "出处" else PP_ALIGN.LEFT
        text(s, x, 1.92, w, 0.24, t, 10.5, FAINT, SANS, True, al, 1.0, cs=1.0, check=False)
    hr(s, ML, 2.20, CW, 1.0, LINE2)
    EV = [("2023-06-23", "美国一名律师用 ChatGPT 写诉状", "引用 6 个不存在的判例", "被罚款 5,000 美元",
           "《卫报》"),
          ("2025-10-06", "德勤为澳大利亚政府做的报告", "虚构文献与法院引语", "退还部分费用，共 44 万澳元",
           "《卫报》"),
          ("2026-10", "本稿实测：HY4 读挡块图纸", "把 −9/sinβ 读成“3-9”", "判为“疑为 1-9 写错”",
           "本地记录"),
          ("2026-10", "本稿实测：Qwen 读同一张图", "把剖面标注 13×9 读成“1349”", "当成真实尺寸继续算",
           "本地记录")]
    for i, (dt, who, what, res, src) in enumerate(EV):
        yy = 2.30 + i * 0.86
        if i:
            hr(s, ML, yy - 0.10, CW, 1.0, LINE)
        text(s, ML, yy + 0.04, 1.30, 0.26, dt, 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 1.36, yy, 3.60, 0.28, who, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 5.02, yy + 0.02, 3.10, 0.28, what, 12.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML + 8.24, yy + 0.02, 2.86, 0.28, res, 12.5, ORANGE, SANS, True, PP_ALIGN.LEFT,
             1.2, check=False)
        text(s, X2 - 0.90, yy + 0.04, 0.90, 0.24, src, 9.5, FAINT, SANS, False, PP_ALIGN.RIGHT,
             1.0, check=False)
    hr(s, ML, 5.72, CW, 1.0, LINE2)
    box(s, ML, 5.90, CW, 0.72, fill=MINT)
    text(s, ML + 0.24, 6.02, CW - 0.48, 0.50,
         "公开案例与本地案例的错法完全一致：不是不知道，是“照着最像的样子补齐”。", 14.5, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.25, check=False)
    footer(s, "来源：《卫报》2023-06-23、2025-10-06；本稿图纸审查十份输出的原文")
    return s


'''
s = s[:i0] + new + s[i1:]
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
