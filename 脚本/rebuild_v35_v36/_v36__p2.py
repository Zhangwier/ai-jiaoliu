# -*- coding: utf-8 -*-
import io
f = "app2.py"
s = io.open(f, encoding="utf-8").read()
start = s.index("def b_others")
new = '''def b_others(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "\u5176\u4f59\u5341\u5bb6\uff1a\u5404\u81ea\u5b88\u4f4f\u4e00\u6bb5\uff0c\u4e5f\u5404\u6709\u4e00\u5904\u786c\u8fb9\u754c",
           "\u4e0d\u662f\u6392\u540d\uff0c\u662f\u201c\u4ec0\u4e48\u60c5\u51b5\u4e0b\u8be5\u60f3\u5230\u5b83\u201d\u3002\u5206\u6570\u53d6 AA \u667a\u80fd\u6307\u6570 v4.3.2\uff1b\u6ca1\u6709\u540c\u53e3\u5f84\u5206\u6570\u7684\u5199\u201c\u2014\u201d\u3002")
    rows = [["\u516c\u53f8", "\u4e3b\u529b\u6a21\u578b\u4e0e\u89c4\u683c", "AA", "\u5f3a\u9879", "\u8fb9\u754c"],
            ["Google", "Gemini 4 Argon \u00b7 3.8 Flash\uff1b\u5355\u6b21\u8f93\u51fa 1M", "53", "\u77e5\u8bc6\u9762\u6700\u5e7f\u3001\u4e2d\u6587\u6700\u987a", "\u667a\u80fd\u4f53\u699c\u504f\u540e\uff1b\u4e0d\u53cd\u5bf9\u4f60"],
            ["Meta", "Muse Wonder / Wonder Full\uff082026-10-12\uff09", "52", "\u957f\u671f\u8bb0\u5fc6\u4e0e\u4e13\u6ce8\uff08Active Memory\uff09", "\u53ea\u5bf9\u7f8e\u52a0\u5f00\u653e\uff1b$20 / $100"],
            ["xAI", "Grok 4.7\uff1bGrok Bot\uff082026-08-11 \u516c\u6d4b\uff09", "46", "\u6bcf\u4e2a Bot \u4e00\u53f0\u4e91\u7aef\u7535\u8111\uff0c\u5e38\u9a7b\u5e72\u6d3b", "\u5e38\u9a7b\uff1d\u4ea4\u51fa\u8d26\u53f7\u6743\u9650"],
            ["Mistral", "Mistral Large 4\uff1a\u7ea6 1T MoE\u30011M \u4e0a\u4e0b\u6587", "\u2014", "\u5f00\u653e\u6743\u91cd\u3001\u6b27\u6d32\u5408\u89c4\u3001\u53ef\u79c1\u6709\u5316", "\u751f\u6001\u4e0e\u6559\u7a0b\u504f\u5c11"],
            ["DeepSeek", "V4.1 Flash 1017\uff1a1M\uff1b\u8f93\u51fa 8 \u5143/\u767e\u4e07", "39", "\u6700\u4fbf\u5b9c\u6700\u5feb\uff08\u7ea6 217 token/s\uff09", "\u5148\u5927\u6bb5\u601d\u8003\uff1b\u590d\u6742\u5224\u65ad\u8981\u6838"],
            ["Qwen", "Qwen3.8 Max\uff1a2.4T MoE\uff1bFlash", "45", "\u5c3a\u5bf8\u68af\u961f 0.6B\u21922.4T \u6700\u5168", "\u65d7\u8230\u4e0e\u9876\u5c16\u5dee\u7ea6 10 \u5206"],
            ["GLM", "GLM-5.3\uff1b5.3-Flash", "45", "\u7f16\u7a0b\u8f83 5.2 \u63d0\u5347 50%\uff0c\u5f00\u6e90 SOTA", "\u901a\u7528\u5bf9\u8bdd\u4e0d\u662f\u5f3a\u9879"],
            ["Kimi", "K3\uff1a2.8T \u53c2\u6570\u30011M\u3001\u5f00\u653e\u6743\u91cd", "44", "\u957f\u6587\u6863\uff1bKimi Delta Attention \u7701\u663e\u5b58", "\u751f\u6001\u4e0e\u5de5\u5177\u94fe\u8f83\u65b0"],
            ["MiniMax", "M3\uff1a\u7ea6 428B \u603b / 23B \u6fc0\u6d3b\uff1b\u6d77\u87ba 3.0", "29", "\u591a\u6a21\u6001\u4e0e\u8bed\u97f3\uff0c\u6210\u672c\u4f4e", "\u7efc\u5408\u5206\u5dee\u8ddd\u660e\u663e"],
            ["\u817e\u8baf\u6df7\u5143", "Hy4 preview\uff1a770B \u603b / 49B \u6fc0\u6d3b\u30011M", "\u2014", "\u5143\u5b9d\u3001ima\u3001WorkBuddy \u4e00\u5957\u5916\u58f3", "LMArena \u667a\u80fd\u4f53\u7b2c 20\uff1b\u9ed8\u8ba4\u672a\u5f00\u89c6\u89c9"],
            ["\u767e\u5ea6\u6587\u5fc3", "\u6587\u5fc3 5.0\uff082.4T\uff09\u30015.1\uff1b\u9884\u8bad\u7ea6 6%", "\u2014", "\u5343\u5e06\u751f\u6001\u3001\u5355\u4f4d\u91c7\u8d2d\u8def\u5f84\u6e05\u695a", "\u65e0\u540c\u53e3\u5f84\u7b2c\u4e09\u65b9\u5206\u6570"]]
    table(s, ML, y + 0.10, CW, rows, [1.16, 4.24, 0.58, 3.62, 2.54], size=10.5, head_size=10.5,
          aligns=[PP_ALIGN.LEFT] * 5, minrow=0.325, head_h=0.34, pad=0.10,
          colors=lambda r, c: (NAVY if c == 0 else (TEAL if c == 2 else (C(0x2B, 0x34, 0x45)))),
          rowlines=True)
    footer(s, "\u6765\u6e90\uff1a\u5404\u516c\u53f8\u6a21\u578b\u9875\u4e0e\u516c\u544a\u3001artificialanalysis.ai\uff1b2026-10-18 \u67e5\u8be2")
    return s
'''
io.open(f, "w", encoding="utf-8").write(s[:start] + new)
print("b_others replaced")
