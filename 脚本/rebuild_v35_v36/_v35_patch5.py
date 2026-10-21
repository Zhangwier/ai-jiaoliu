# -*- coding: utf-8 -*-
"""修正第 02 章的 5 处问题"""
p = "p_ch2.py"
s = open(p, encoding="utf-8").read()
o = s

# 1) P19：底部多余说明并入来源行
s = s.replace('''    text(s, ML, 6.58, CW, 0.24,
         "一次生成 · 无历史记忆 · 统一视口录屏（去黑边）　｜　十个配置覆盖了国际三强、国内头部与两款“自动档”客户端",
         9.5, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：10 份原始 HTML 录屏（统一视口，去黑边）")''',
'''    footer(s, "来源：10 份原始 HTML 录屏（统一视口，去黑边）；一次生成、无历史记忆")''')

# 2) P21：坐标轴不与结论重叠
s = s.replace('''    # 轴
    for t in (35, 40, 45, 50, 55, 60):
        vr(s, fx(t), 2.14, 3.70, 0.6, C(0xE4, 0xDF, 0xD3))
        text(s, fx(t) - 0.30, 5.92, 0.60, 0.22, str(t), 8.5, FAINT, MONO, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    text(s, x0, 6.16, 5.40, 0.22, "AA 智能指数 v4.3.2（越高越好）", 8.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)''',
'''    # 轴
    for t in (35, 40, 45, 50, 55, 60):
        vr(s, fx(t), 2.20, 3.02, 0.6, C(0xE4, 0xDF, 0xD3))
        text(s, fx(t) - 0.30, 5.28, 0.60, 0.22, str(t), 8.5, FAINT, MONO, False,
             PP_ALIGN.CENTER, 1.0, check=False)
    text(s, x0 - 0.42, 5.28, 1.6, 0.22, "AA 智能指数", 8.5, FAINT, SANS, False,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, x1 - 1.9, 5.52, 1.9, 0.22, "（越高越好）", 8.5, FAINT, SANS, False,
         PP_ALIGN.RIGHT, 1.0, check=False)''')

# 3) P22：六项统一用绿色竖线；第三列加宽避免折行
s = s.replace('vr(s, cx - 0.16, cy + 0.02, 0.62, 1.4, GRN if i < 4 else RED)',
              'vr(s, cx - 0.16, cy + 0.02, 0.62, 1.4, GRN)')
s = s.replace('''        text(s, x0 + 1.66, yy, 2.90, 0.24, b, 9.8, BODY, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x0 + 4.58, yy, 1.10, 0.24, c, 9.2, MUT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)''',
'''        text(s, x0 + 1.62, yy, 2.62, 0.24, b, 9.6, BODY, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x0 + 4.28, yy, 1.62, 0.40, c, 9.2, MUT, SANS, False, PP_ALIGN.LEFT, 1.25, check=False)''')

# 4) P25：模型列不越界；底部两段对齐同一行
s = s.replace('text(s, ML + 4.55, yy + 0.02, 7.0, 0.52, models, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)',
              'text(s, ML + 4.48, yy + 0.02, 6.92, 0.52, models, 10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.45)')
s = s.replace('''    text(s, ML, 6.12, 6.4, 0.50,
         "AA 开源模型前六名全是中国模型。DeepSeek 旧版 V4.1 Flash 只有 39 分，但最便宜、最快。",
         10.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.45)
    text(s, 7.85, 6.08, 4.55, 0.50,
         "最强选 OpenAI / Claude；国产 GLM、Qwen、Kimi、DeepSeek 够日常用。",
         11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.45)''',
'''    hr(s, ML, 6.02, CW, 0.75, RULE)
    text(s, ML, 6.16, 6.35, 0.52,
         "AA 开源模型前六名全是中国模型。DeepSeek 旧版 V4.1 Flash 只有 39 分，但最便宜、最快。",
         10.5, MUT, SANS, False, PP_ALIGN.LEFT, 1.45)
    text(s, ML + 6.85, 6.12, 4.68, 0.54,
         "最强选 OpenAI / Claude；国产 GLM、Qwen、Kimi、DeepSeek 够日常用。",
         11.5, INK, SANS, True, PP_ALIGN.LEFT, 1.42)''')

# 5) P24：表头下移一点
s = s.replace('text(s, xs[0], 1.90, 3.0, 0.22, "榜单"', 'text(s, xs[0], 1.98, 3.0, 0.22, "榜单"')
s = s.replace('text(s, xs[1], 1.90, 2.4, 0.22, "主要测什么"', 'text(s, xs[1], 1.98, 2.4, 0.22, "主要测什么"')
s = s.replace('text(s, xs[2], 1.90, 3.2, 0.22, "第一名"', 'text(s, xs[2], 1.98, 3.2, 0.22, "第一名"')
s = s.replace('text(s, xs[3], 1.90, 3.2, 0.22, "紧随其后"', 'text(s, xs[3], 1.98, 3.2, 0.22, "紧随其后"')
s = s.replace('    hr(s, ML, 2.18, CW, 1.1, RULE2)\n    yy = 2.32',
              '    hr(s, ML, 2.26, CW, 1.1, RULE2)\n    yy = 2.38')
s = s.replace('    text(s, ML, 5.98, 2.4, 0.24, "一句话"', '    text(s, ML, 6.02, 2.4, 0.24, "一句话"')
s = s.replace('    text(s, ML + 1.10, 5.94, 10.30, 0.34,', '    text(s, ML + 1.10, 5.98, 10.30, 0.34,')

if s == o:
    raise SystemExit("NOTHING CHANGED")
open(p, "w", encoding="utf-8").write(s)
print("patched")
