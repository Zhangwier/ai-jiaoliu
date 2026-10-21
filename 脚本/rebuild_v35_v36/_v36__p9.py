# -*- coding: utf-8 -*-
import io
f="app.py"; s=io.open(f,encoding="utf-8").read()
start=s.index("def b_matrix(prs, pno):")
end=s.index("def b_top(prs, pno, which):")
new = '''def b_matrix(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "\u5341\u4e09\u5bb6\u6a21\u578b\u516c\u53f8\uff1a\u4e00\u5f20\u8868\u770b\u5b8c\u5404\u81ea\u7684\u6218\u573a\u4e0e\u8fb9\u754c",
           "\u5206\u6570\u53d6 AA \u667a\u80fd\u6307\u6570 v4.3.2\uff082026-10-18 \u67e5\u8be2\uff09\uff1b\u201c\u2014\u201d\u8868\u793a\u6ca1\u6709\u540c\u53e3\u5f84\u5206\u6570\uff0c\u4e0d\u4f5c\u6392\u540d\u4f9d\u636e\u3002")
    cw = (CW - 0.58) / 2
    LX, RX = ML, ML + cw + 0.58
    vr(s, RX - 0.29, y + 0.02, 4.30, 1.0, LINE)
    for i, (nm, mdl, aa, w, good, bad, _img) in enumerate(BRAND):
        left = i < 7
        x = LX if left else RX
        yy = y + 0.04 + (i if left else i - 7) * 0.585
        hr(s, x, yy, cw, 1.0, LINE)
        text(s, x, yy + 0.07, 1.30, 0.24, nm, 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
             check=False)
        text(s, x + 1.32, yy + 0.09, 0.72, 0.22, w, 8.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + cw - 0.44, yy + 0.06, 0.44, 0.24, aa, 11.5, TEAL, MONO, True, PP_ALIGN.RIGHT,
             1.0, check=False)
        text(s, x, yy + 0.28, cw, 0.20, mdl, 9.2, GREY, SANS, False, PP_ALIGN.LEFT, 1.05, check=False)
        rich(s, x, yy + 0.46, cw, 0.20,
             [{"runs": [(good + "\u3000", C(0x2B, 0x34, 0x45), 9.0, False),
                        ("\u8fb9\u754c\uff1a" + bad, ORANGE, 9.0, False)], "ls": 1.05}])
    hr(s, LX, y + 0.04 + 7 * 0.585, cw, 1.0, LINE)
    hr(s, RX, y + 0.04 + 6 * 0.585, cw, 1.0, LINE)
    text(s, ML, 6.44, CW, 0.26,
         "\u770b\u6cd5\uff1a\u5206\u6570\u53ea\u80fd\u6392\u51fa\u201c\u8c01\u66f4\u806a\u660e\u201d\uff0c\u6392\u4e0d\u51fa\u201c\u4ec0\u4e48\u6d3b\u8be5\u4ea4\u7ed9\u8c01\u201d\u3002\u4e0b\u9762\u4e24\u9875\u5c55\u5f00 OpenAI \u4e0e Anthropic\uff0c\u5176\u4f59\u5341\u4e00\u5bb6\u5728\u7b2c 63 \u9875\u3002",
         10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "\u6765\u6e90\uff1aartificialanalysis.ai\uff1b\u5404\u516c\u53f8\u6a21\u578b\u9875\u4e0e\u516c\u544a\uff082026-10-18 \u67e5\u8be2\uff09\uff1b\u672c\u7a3f\u6574\u7406")
    return s


'''
io.open(f,"w",encoding="utf-8").write(s[:start]+new+s[end:])
print("ok")
