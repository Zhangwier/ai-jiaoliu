# -*- coding: utf-8 -*-
import io
p = r"D:\MG\program\ai交流\_v36\ch2.py"
s = io.open(p, encoding="utf-8").read()

i0 = s.index('def p25(prs, pno):')
i1 = s.index('def p26(prs, pno):')

new = '''def p25(prs, pno):
    s = slide(prs)
    y = head(s, CH, pno, "美国两强领跑，国产头部落后约十分",
             "按 AA 智能指数分层：第一梯队 52 分以上，第二梯队 44–46 分，其后 29 分上下。")
    ty = y + 0.16
    LAY = [("第一梯队", "OpenAI · Anthropic", "Astra 53、Sol 52 ／ Opus 58、Sonnet 56", 1.00,
            C(0x4A, 0x6C, 0x9E)),
           ("紧随其后", "Google · Meta", "Argon 53 ／ Muse Wonder 52", 0.84, TEAL),
           ("第二梯队", "xAI · 国产头部",
            "Grok 4.7 46、MiMo 46、Qwen3.8 Max 45、GLM-5.3 45、Kimi K3 44", 0.66,
            C(0x63, 0x71, 0x88)),
           ("其后", "腾讯 · MiniMax · 字节", "Hy4 未进同口径榜；MiniMax-M3 29；豆包无同口径数据",
            0.40, C(0x9A, 0xA1, 0xAC))]
    BX, BW = 2.46, 10.24
    for i, (nm, who, models, frac, col) in enumerate(LAY):
        yy = ty + i * 0.94
        text(s, ML, yy - 0.02, 2.00, 0.28, nm, 13.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML, yy + 0.28, 2.02, 0.24, who, 9.5, GREY2, SANS, False, PP_ALIGN.LEFT, 1.0,
             check=False)
        box(s, BX, yy + 0.02, BW, 0.28, fill=C(0xF0, 0xEE, 0xE7))
        bw = BW * frac
        box(s, BX, yy + 0.02, bw, 0.28, fill=col)
        inside = BX + 0.14
        room = bw - 0.24
        if est_w(models, 10.5, True) / 72.0 <= room:
            text(s, inside, yy + 0.075, room, 0.24, models, 10.5, WHITE, SANS, True,
                 PP_ALIGN.LEFT, 1.0, check=False)
        else:
            text(s, BX + bw + 0.22, yy + 0.075, BW - bw - 0.30, 0.24, models, 10.5, GREY, SANS,
                 False, PP_ALIGN.LEFT, 1.0, check=False)
    hr(s, ML, 6.04, CW, 1.0, LINE)
    text(s, ML, 6.16, 6.10, 0.28, "另一条值得注意：AA 开源前六名，全部是中国模型", 12.5, NAVY, SANS,
         True, PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML, 6.44, 6.10, 0.26, "闭源看两强，开源看国产 —— 这是当前的格局", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.2, check=False)
    text(s, ML + 6.84, 6.16, 5.28, 0.60,
         "分层比的是“最好那一档”。日常任务里，第二梯队与第一梯队的差距往往小于\\n“外壳是否合适”带来的差距。",
         11, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：AA 智能指数 v4.3.2；各厂商公开模型说明")
    return s


'''
s = s[:i0] + new + s[i1:]
io.open(p, "w", encoding="utf-8").write(s)
print("ok")
