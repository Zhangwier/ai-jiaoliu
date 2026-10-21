# -*- coding: utf-8 -*-
import io
f = "app5.py"
s = io.open(f, encoding="utf-8").read()
start = s.index("def z2(prs, pno):")
new = '''def z2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "\u94fe\u63a5\u6e05\u5355\uff1a\u653f\u7b56\u3001\u5224\u4f8b\u3001\u5de5\u5177\u4e0e\u793e\u533a",
           "\u53ea\u5217\u5b98\u65b9\u9875\u9762\u4e0e\u4e00\u624b\u6765\u6e90\uff1b\u793e\u533a\u5e16\u6807\u6ce8\u4e3a\u793e\u533a\u3002\u67e5\u8be2\u65e5\u671f\u7edf\u4e00\u4e3a 2026-10-18\u3002")
    G = [("\u653f\u7b56\u4e0e\u5408\u89c4", ["openai.com/policies/usage-policies/", "anthropic.com/legal/aup",
                     "genai.owasp.org/llm-top-10/", "cac.gov.cn\uff08\u751f\u6210\u5f0f AI \u5907\u6848\u516c\u544a\uff09"]),
         ("\u5de5\u5177\u4e0e\u5f00\u6e90", ["github.com/yetone/magpie", "github.com/router-for-me/CLIProxyAPI",
                    "github.com/YuJunZhiXue/dsh-purge", "github.com/himovo/movo",
                    "github.com/Stonewuu/agenteam"]),
         ("\u793e\u533a", ["linux.do\uff08\u9080\u8bf7\u5236\uff09", "bbs.kanxue.com/thread-255826.htm",
                "daily.juya.uk\uff08\u4e2d\u6587\u65e5\u62a5\uff09"]),
         ("\u4e8b\u4ef6\u4e0e\u5224\u4f8b", ["huggingface.co/blog/security-incident-july-2026",
                    "openai.com/index/hugging-face-model-evaluation-security-incident/",
                    "bbc.com/news/articles/cd9w22n9e4go",
                    "cnbc.com/2026/09/30/openai-sued-cyberattack.html",
                    "therecord.media/south-korean-bank-hacks-ai-agents"]),
         ("\u699c\u5355\u4e0e\u68c0\u6d4b", ["artificialanalysis.ai", "epoch.ai/frontiermath", "arcprize.org",
                    "lmarena.ai", "livebench.ai", "lm.ikale.io", "marginlab.ai"])]
    cw = (CW - 0.62) / 2
    TOPY = y + 0.14
    # \u5de6\u5217
    yy = TOPY
    for gi in (0, 1, 2):
        t, links = G[gi]
        text(s, ML, yy, cw, 0.22, t, 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
        ly = yy + 0.28
        for ln in links:
            text(s, ML + 0.02, ly, cw - 0.04, 0.22, ln, 9, BODY, MONO, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
            ly += 0.25
        yy = ly + 0.22
    # \u53f3\u5217
    RX = ML + cw + 0.62
    yy = TOPY
    for gi in (3, 4):
        t, links = G[gi]
        text(s, RX, yy, cw, 0.22, t, 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
        ly = yy + 0.28
        for ln in links:
            text(s, RX + 0.02, ly, cw - 0.04, 0.22, ln, 9, BODY, MONO, False, PP_ALIGN.LEFT, 1.0,
                 check=False)
            ly += 0.25
        yy = ly + 0.22
    vr(s, RX - 0.31, TOPY - 0.04, 4.62, 1.0, LINE)
    hr(s, ML, 6.10, CW, 1.0, LINE)
    text(s, ML, 6.24, CW, 0.30,
         "\u5f15\u7528\u539f\u5219\uff1a\u5148\u770b\u5b98\u65b9\u539f\u5e16\u4e0e\u539f\u59cb\u6587\u6863\uff0c\u518d\u770b\u4e8c\u624b\u6574\u7406\uff1b\u793e\u533a\u5e16\u53ea\u5f53\u7ebf\u7d22\uff0c\u4e0d\u5f53\u7ed3\u8bba\u3002", 11, TEAL, SANS,
         True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "\u6765\u6e90\uff1a\u4ee5\u4e0a\u5404\u7ad9\u70b9\uff082026-10-18 \u67e5\u8be2\uff09\uff1b\u793e\u533a\u5185\u5bb9\u4e3a\u4e2a\u4eba\u8868\u8ff0\uff0c\u672a\u7ecf\u72ec\u7acb\u6838\u9a8c")
    return s
'''
io.open(f, "w", encoding="utf-8").write(s[:start] + new)
print("z2 replaced")
