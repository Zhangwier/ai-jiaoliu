# -*- coding: utf-8 -*-
import io
f="app.py"; s=io.open(f,encoding="utf-8").read()
start=s.index("def b_matrix(prs, pno):")
end=s.index("def b_top(prs, pno, which):")
new = '''MX = [("OpenAI", "GPT-6 Astra \\u00b7 6.1 Sol \\u00b7 6 Luna", "53", "\u5168\u95ed\u6e90",
       "\u63a8\u7406\u4e0e\u6570\u5b66\u7b2c\u4e00\u68af\u961f\uff1bCodex \u56db\u79cd\u5f62\u6001", "$10 / $50 \u6bcf\u767e\u4e07 token"),
      ("Anthropic", "Opus 5.5 \\u00b7 Sonnet 5.5 \\u00b7 Haiku 5.5", "58", "\u5168\u95ed\u6e90",
       "\u957f\u4efb\u52a1\u4e0d\u8dd1\u504f\u3001\u4f1a\u81ea\u5df1\u67e5\u8bc1", "\u65e0\u514d\u8d39\u7248\uff1b\u5730\u533a\u9650\u5236\u6700\u786c"),
      ("Google", "Gemini 4 Argon \\u00b7 3.8 Flash", "53", "\u5168\u95ed\u6e90",
       "\u77e5\u8bc6\u9762\u3001\u4e2d\u6587\u3001\u957f\u6587\uff1b\u5355\u6b21 1M \u8f93\u51fa", "\u667a\u80fd\u4f53\u699c\u504f\u540e\uff1b\u4e0d\u53cd\u5bf9\u4f60"),
      ("Meta", "Muse Wonder \\u00b7 Wonder Full", "52", "\u95ed\u6e90",
       "\u957f\u671f\u8bb0\u5fc6\u4e0e\u4e2a\u4eba\u4e8b\u52a1", "\u53ea\u5bf9\u7f8e\u52a0\uff1b\u7f16\u7a0b\u4e0d\u662f\u4e3b\u573a"),
      ("xAI", "Grok 4.7 \\u00b7 Grok Bot", "46", "\u4e0d\u5f00\u6e90",
       "\u5e38\u9a7b\u578b Agent\uff1a\u6bcf\u4e2a Bot \u4e00\u53f0\u4e91\u7535\u8111", "\u5e38\u9a7b\uff1d\u4ea4\u51fa\u8d26\u53f7\u6743\u9650"),
      ("Mistral", "Mistral Large 4", "\\u2014", "\u5f00\u653e\u6743\u91cd",
       "\u7ea6 1T MoE\u30011M \u4e0a\u4e0b\u6587\uff1b\u6b27\u6d32\u5408\u89c4", "\u751f\u6001\u4e0e\u6559\u7a0b\u504f\u5c11"),
      ("DeepSeek", "V4.1 Flash", "39", "\u5f00\u653e\u6743\u91cd",
       "\u6700\u4fbf\u5b9c\u3001\u51fa\u5b57\u6700\u5feb\uff1b\u4e94\u4e2a\u5e95\u5c42\u5e93\u5f00\u6e90", "\u5148\u5927\u6bb5\u601d\u8003\uff1b\u590d\u6742\u5224\u65ad\u8981\u6838"),
      ("Qwen", "Qwen3.8 Max \\u00b7 Flash", "45", "\u5f00\u653e\u6743\u91cd",
       "\u5c3a\u5bf8\u68af\u961f 0.6B\\u21922.4T \u6700\u5168", "\u65d7\u8230\u4e0e\u9876\u5c16\u5dee\u7ea6 10 \u5206"),
      ("GLM", "GLM-5.3 \\u00b7 5.3-Flash", "45", "\u5f00\u653e\u6743\u91cd",
       "\u7f16\u7a0b\u63d0\u5347 50%\uff1b\u5f00\u6e90 SOTA", "\u901a\u7528\u5bf9\u8bdd\u4e0d\u662f\u5f3a\u9879"),
      ("Kimi", "K3", "44", "\u5f00\u653e\u6743\u91cd",
       "2.8T \u53c2\u6570\u30011M \u4e0a\u4e0b\u6587\uff1b\u957f\u6587\u6863", "\u751f\u6001\u4e0e\u5de5\u5177\u94fe\u8f83\u65b0"),
      ("MiniMax", "M3 \\u00b7 \u6d77\u87ba 3.0", "29", "\u5f00\u653e\u6743\u91cd",
       "\u591a\u6a21\u6001\u4e0e\u8bed\u97f3\uff1b\u6210\u672c\u4f4e", "\u7efc\u5408\u5206\u5dee\u8ddd\u660e\u663e"),
      ("\\u817e\\u8baf\\u6df7\\u5143", "Hy4 preview", "\\u2014", "\u4e0d\u5f00\u6e90",
       "770B / 49B \u6fc0\u6d3b\uff1b\u5143\u5b9d\u3001ima\u3001WorkBuddy", "\u667a\u80fd\u4f53\u699c\u7b2c 20\uff1b\u9ed8\u8ba4\u672a\u5f00\u89c6\u89c9"),
      ("\\u767e\\u5ea6\\u6587\\u5fc3", "\\u6587\\u5fc3 5.0 \\u00b7 5.1", "\\u2014", "\u4e0d\u5f00\u6e90",
       "\u9884\u8bad\u7ec3\u6210\u672c\u7ea6\u540c\u89c4\u6a21 6%\uff1b\u5343\u5e06\u751f\u6001", "\u65e0\u540c\u53e3\u5f84\u7b2c\u4e09\u65b9\u5206\u6570")]


def b_matrix(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "\u5341\u4e09\u5bb6\u6a21\u578b\u516c\u53f8\uff1a\u4e00\u5f20\u8868\u770b\u5b8c\u5404\u81ea\u7684\u6218\u573a\u4e0e\u8fb9\u754c",
           "\u5206\u6570\u53d6 AA \u667a\u80fd\u6307\u6570 v4.3.2\uff082026-10-18 \u67e5\u8be2\uff09\uff1b\u201c\u2014\u201d\u8868\u793a\u6ca1\u6709\u540c\u53e3\u5f84\u5206\u6570\uff0c\u4e0d\u4f5c\u6392\u540d\u4f9d\u636e\u3002")
    cw = (CW - 0.58) / 2
    LX, RX = ML, ML + cw + 0.58
    vr(s, RX - 0.29, y + 0.02, 4.24, 1.0, LINE)
    P = 0.60
    for i, (nm, mdl, aa, w, good, bad) in enumerate(MX):
        left = i < 7
        x = LX if left else RX
        yy = y + 0.02 + (i if left else i - 7) * P
        hr(s, x, yy, cw, 1.0, LINE)
        text(s, x, yy + 0.06, 1.70, 0.22, nm, 11.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
             check=False)
        text(s, x + 1.74, yy + 0.08, 1.10, 0.20, w, 8.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + cw - 0.42, yy + 0.05, 0.42, 0.22, aa, 12, TEAL, MONO, True, PP_ALIGN.RIGHT,
             1.0, check=False)
        text(s, x, yy + 0.25, cw, 0.18, mdl, 9.2, GREY, SANS, False, PP_ALIGN.LEFT, 1.05, check=False)
        rich(s, x, yy + 0.42, cw, 0.18,
             [{"runs": [(good + "\u3000", {"size": 8.8, "color": C(0x2B, 0x34, 0x45)}),
                        ("\u8fb9\u754c\uff1a" + bad, {"size": 8.8, "color": ORANGE})], "ls": 1.02}])
    hr(s, LX, y + 0.02 + 7 * P, cw, 1.0, LINE)
    hr(s, RX, y + 0.02 + 6 * P, cw, 1.0, LINE)
    text(s, ML, 6.46, CW, 0.24,
         "\u770b\u6cd5\uff1a\u5206\u6570\u53ea\u80fd\u6392\u51fa\u201c\u8c01\u66f4\u806a\u660e\u201d\uff0c\u6392\u4e0d\u51fa\u201c\u4ec0\u4e48\u6d3b\u8be5\u4ea4\u7ed9\u8c01\u201d\u3002\u4e0b\u4e24\u9875\u5c55\u5f00 OpenAI \u4e0e Anthropic\uff0c\u5176\u4f59\u5341\u4e00\u5bb6\u5728\u7b2c 63 \u9875\u3002",
         10.2, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "\u6765\u6e90\uff1aartificialanalysis.ai\uff1b\u5404\u516c\u53f8\u6a21\u578b\u9875\u4e0e\u516c\u544a\uff082026-10-18 \u67e5\u8be2\uff09\uff1b\u672c\u7a3f\u6574\u7406")
    return s


'''
io.open(f,"w",encoding="utf-8").write(s[:start]+new+s[end:])
print("ok")
