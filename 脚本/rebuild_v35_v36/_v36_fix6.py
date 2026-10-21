# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

i0 = s.index('def p15(prs, pno):')
i1 = s.index('def p16(prs, pno):')

new = '''def p15(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "更隐蔽的一种：它接受了你给的前提，然后认真算错",
             "表头写“农用地”，实际填的是“耕地”。它不会质疑表头，只会在给定的范围里找问题 —— 于是整批漏检。")
    rows = [["序号", "地类", "面积（亩）", "补偿标准（元/亩）"],
            ["1", "耕地", "12.60", "38,000"],
            ["2", "耕地", "8.35", "38,000"],
            ["3", "耕地", "21.04", "38,000"]]
    table(s, ML, 2.02, 6.30, rows, [0.80, 1.60, 1.90, 2.00], size=13, head_size=13,
          aligns=[PP_ALIGN.CENTER, PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], minrow=0.38,
          head_h=0.42, first_bold=False)
    text(s, ML, 3.86, 6.30, 0.26, "↑ 表头与实际地类不一致：应为「耕地」", 12, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 4.26, 6.30, 1.0, LINE)
    text(s, ML, 4.38, 6.30, 0.26, "三个机制叠在一起，错就被盖住了", 12, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    M = [("前提顺从", "你给的表头、口径、结论，它默认成立，不回头质疑。"),
         ("范围即上限", "它只在你划的范围内找问题；范围外的错，它不越界去看。"),
         ("自洽掩盖错误", "后面每一步都基于这个错前提，环环相扣，读起来毫无破绽。")]
    yy = 4.74
    for i, (t, b) in enumerate(M):
        text(s, ML, yy + 0.02, 1.60, 0.24, t, 12.5, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 1.66, yy, 4.64, 0.50, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.36,
             check=False)
        yy += 0.58

    RX = 7.30
    vr(s, RX - 0.36, 2.02, 3.90, 1.0, LINE)
    text(s, RX, 2.02, 5.00, 0.26, "同构的公开案例", 12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, RX, 2.32, 1.60, 0.72, "722", 44, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, RX + 1.42, 2.56, 3.60, 0.48, "篇数学手稿一次性放出", 13, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.3)
    text(s, RX, 3.16, 5.02, 0.74,
         "两天后撤回 3 篇：一处符号应记 −1，写成了 +1。\\n错误不在结论，在最前面那一步的约定。", 12.5,
         C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    hr(s, RX, 3.94, 5.02, 1.0, LINE)
    text(s, RX, 4.08, 5.02, 0.26, "所以要求里要写一句“先复述口径”", 13.5, NAVY, SANS, True,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, RX, 4.38, 5.02, 1.00,
         "让它先把自己看到的口径、单位、范围复述一遍，你确认之后再往下做。\\n口径错了，后面全错；口径对了，错也只错在局部。", 12.5,
         GREY, SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    box(s, RX, 5.52, 5.02, 0.80, fill=ORANGEL)
    text(s, RX + 0.20, 5.62, 4.62, 0.62,
         "本稿实测：十份图纸审查里，没有一份主动质疑过“表头写的是农用地”这类前提。", 12.5,
         C(0x7A, 0x3A, 0x22), SANS, True, PP_ALIGN.LEFT, 1.38)
    footer(s, "来源：本稿图纸审查记录；OpenAI 2025-10 数学手稿事件公开报道")
    return s


'''
s = s[:i0] + new + s[i1:]
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
