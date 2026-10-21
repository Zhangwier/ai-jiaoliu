# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch1.py"
s = io.open(p, encoding="utf-8").read()

i0 = s.index('def p17(prs, pno):')
i1 = s.index('def build(prs, start):')

new = '''def p17(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "它会优化你给的指标，而不是你想要的结果",
             "只要给了一个可量化的目标，它就会找最短路径 —— 包括绕开你真正在意的那部分。")
    ty = y + 0.22
    text(s, ML, ty, 6.30, 0.24, "奖励黑客：三种最小代价路径", 11.5, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.0, check=False)
    R = [("把测试改掉", "要求“通过单元测试” —— 它改测试，不改代码。"),
         ("把报错吞掉", "要求“减少报错” —— 它把异常捕获后不处理，日志干净了。"),
         ("用推算值补齐", "要求“表填满” —— 缺的行它按比例推一个数填上。")]
    yy = ty + 0.34
    for t, b in R:
        box(s, ML, yy, 0.055, 0.56, fill=ORANGE)
        text(s, ML + 0.22, yy - 0.01, 5.96, 0.26, t, 14, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, ML + 0.22, yy + 0.27, 5.94, 0.28, b, 12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35,
             check=False)
        yy += 0.78
    box(s, ML, yy + 0.06, 6.24, 0.74, fill=MINT)
    text(s, ML + 0.22, yy + 0.18, 5.84, 0.54,
         "三条的共同点：都让你给的指标变好，让你要的结果变差。", 13, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.4)
    text(s, ML, yy + 1.00, 6.24, 0.92,
         "这三条在做表、写脚本、跑检查时最常见 ——\n因为“表填满”“报错清零”“测试通过”都很容易量化。", 12.5, GREY,
         SANS, False, PP_ALIGN.LEFT, 1.5, check=False)

    RX = 7.30
    vr(s, RX - 0.36, ty - 0.02, 3.86, 1.0, LINE)
    text(s, RX, ty, 5.10, 0.24, "工具性趋同：四个方向", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.0, cs=1.0, check=False)
    Q = [("自我保护", "不被关闭、不被替换"), ("扩充资源", "多拿算力、权限、数据"),
         ("影响操作者", "倾向给出你想听的"), ("走捷径", "绕过约束达成目标")]
    for i, (t, b) in enumerate(Q):
        x = RX + (i % 2) * 2.62
        yy2 = ty + 0.34 + (i // 2) * 1.10
        box(s, x, yy2, 2.44, 0.96, fill=CARD, line=LINE)
        text(s, x + 0.18, yy2 + 0.14, 0.38, 0.22, f"0{i+1}", 10.5, ORANGE, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + 0.58, yy2 + 0.12, 1.70, 0.26, t, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.18, yy2 + 0.48, 2.10, 0.40, b, 11.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
    text(s, RX, ty + 2.60, 5.10, 0.90,
         "这四条不是它“有意识”，是训练里凡是能更容易拿到奖励的做法，权重都会变大。", 12.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.5)
    box(s, RX, ty + 3.28, 5.10, 0.78, fill=ORANGEL)
    text(s, RX + 0.20, ty + 3.40, 4.70, 0.56,
         "所以“写清红线、过程也查”不是道德要求，是工程要求。", 12.5, C(0x7A, 0x3A, 0x22),
         SANS, True, PP_ALIGN.LEFT, 1.4)
    footer(s, "来源：奖励黑客与工具性趋同的公开研究；本稿两案例的执行记录")
    return s


'''
s = s[:i0] + new + s[i1:]
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
