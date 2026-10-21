# -*- coding: utf-8 -*-
"""附录 B 后段 / C / D / E / F / Z"""
from ds import *
from app import sec, term_rows, num_list, b_top, b_matrix, a1, a2, a_index, BRAND, CH, hd
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

BODY = C(0x2B, 0x34, 0x45)


def lab(s, x, y, w, t, col=TEAL, size=10.5, cs=1.4):
    text(s, x, y, w, 0.22, t, size, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=cs, check=False)


def para(s, x, y, w, t, size=12.5, color=BODY, ls=1.42, bold=False, col_w=None):
    ww = col_w or w
    nl = nlines(t, size, ww, bold)
    h = nl * size * LINES * ls / 72 + 0.05
    text(s, x, y, w, h, t, size, color, SANS, bold, PP_ALIGN.LEFT, ls, check=False)
    return y + h


def kv(s, x, y, w, items, ts=12.5, gs=10.5, gap=0.13, lab_w=None):
    yy = y
    for t, b in items:
        lab(s, x, yy, w, t, TEAL, gs)
        yy = para(s, x, yy + 0.19, w, b, ts) + gap
    return yy


# ---------------- B1b OpenAI 数学 ----------------
def b_math(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "OpenAI 的数学：能解真问题了，但仍要挑着解",
           "结论要看有没有可复查的证明。同一家公司在一年里翻过一次车，也交出了附 Lean 证明的成果 —— 两件事都记下来。")
    text(s, ML, y + 0.16, 2.30, 0.24, "一道题的得分率", 10.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, ML, y + 0.34, 1.60, 0.90, "3%", 60, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML, y + 1.14, 2.40, 0.52,
         "68 道未解的 Erdős 题，每题 $300 预算、要求机器可验证：GPT-6 Astra 解出 3%，其余模型 0%（FrontierMath Erdős，2026-09）",
         9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    LX = ML + 2.62
    hr(s, LX - 0.24, y + 0.10, 1.0, 3.90, LINE)
    rows = [("2025-10", "先翻过车", "称“解决 10 道 Erdős 题”，实为检索到已有文献", "公开更正；这一条最该记住"),
            ("2026-05", "冷静看", "内部模型给出 Erdős 单位距离猜想的反例", "Alon 等数学家人工核验"),
            ("2026-08", "十项成果", "Astra 内部版：球堆积、密码、Ramsey 数等", "算力约 $2,000"),
            ("2026-09", "千禧年难题", "称约 1 万个 Agent 跑 88 小时给出纳维–斯托克斯“爆破”证明", "学界评审中，未定论"),
            ("2026-10", "再公开一批", "放上 GitHub，附 Lean 形式化证明", "平均约 3 小时 Pro 算力")]
    yy = y + 0.08
    for d, tag, what, note in rows:
        hr(s, LX, yy, X2 - LX, 1.0, LINE)
        text(s, LX, yy + 0.10, 0.86, 0.24, d, 10.5, FAINT, MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, LX + 0.92, yy + 0.09, 1.10, 0.26, tag, 13, TEAL if tag != "先翻过车" else ORANGE,
             SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, LX + 2.14, yy + 0.10, 4.20, 0.30, what, 12, BODY, SANS, False, PP_ALIGN.LEFT, 1.25,
             check=False)
        text(s, LX + 6.46, yy + 0.11, 2.44, 0.30, note, 10.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.25,
             check=False)
        yy += 0.66
    hr(s, LX, yy, X2 - LX, 1.0, LINE)
    text(s, ML, 6.10, CW, 0.30,
         "读法：先问“有没有能复查的证明”，再问“解了几道”。前者决定结论能不能用，后者只说明它在这一类题上强。",
         12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：openai.com 研究公告（2026-05-20、08-01、09-08、10-06）；The Guardian；arXiv 2609.25050；epoch.ai/frontiermath")
    return s


# ---------------- B1c 降智 ----------------
def b_degrade(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "“降智”：官方承认过 bug，没承认按负载降智",
           "三步走完再决定换不换。重要交付别押在一次生成上 —— 体感会骗人，核过的结果不会。")
    IT = [("①", "先查自己", "对话是不是太长、换了客户端或档位、只试了一次。先自己排除这三样，再怀疑模型。", MINT),
          ("②", "对不对得上指纹", "lm.ikale.io 用“随机数指纹”猜背后是哪个模型；结果只是参考，Key 会经过它的代理。", TEALXL),
          ("③", "再看长期曲线", "Margin Lab 每天跑 50 题：单日波动 ±12.8 个点内都算正常，要看 7 日、30 日。", ORANGEL)]
    yy = y + 0.10
    for no, t, b, fill in IT:
        box(s, ML, yy, CW, 1.06, fill)
        text(s, ML + 0.26, yy + 0.26, 0.40, 0.40, no, 20, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.86, yy + 0.22, 2.10, 0.30, t, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 3.10, yy + 0.24, CW - 3.30, 0.60, b, 12.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.4,
             check=False)
        yy += 1.20
    box(s, ML, yy + 0.02, CW, 0.86, CARD, LINE)
    text(s, ML + 0.26, yy + 0.16, 2.20, 0.24, "官方的一份时间线", 10.5, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    text(s, ML + 0.26, yy + 0.38, CW - 0.52, 0.42,
         "Anthropic：2025 年 8–9 月三个 bug，最严重时 16% 的 Sonnet 4 请求受影响；官方称“从不因需求或负载降低质量”。",
         11.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：Anthropic 事故复盘（2025-09-17）；marginlab.ai；lm.ikale.io/docs")
    return s


# ---------------- B3 其余各家 ----------------
BR2 = [("Google", "Gemini 4 Argon · 3.8 Flash", "53", "知识面最广、中文最顺；单次输出 1M", "智能体榜偏后；不反对你"),
       ("Meta", "Muse Wonder / Wonder Full", "52", "长期记忆与专注（Active Memory）", "只对美加开放；$20 / $100"),
       ("xAI", "Grok 4.7 · Grok Bot", "46", "每个 Bot 一台云端电脑，常驻干活", "常驻＝交出账号权限"),
       ("Mistral", "Mistral Large 4", "—", "开放权重、欧洲合规、可私有化", "生态与教程偏少"),
       ("DeepSeek", "V4.1 Flash 1017", "39", "最便宜最快（约 217 token/s）", "先大段思考；复杂判断要核"),
       ("Qwen", "Qwen3.8 Max · Flash", "45", "尺寸梯队 0.6B→2.4T 最全", "旗舰与顶尖差约 10 分"),
       ("GLM", "GLM-5.3 · 5.3-Flash", "45", "编程较 5.2 提升 50%，开源 SOTA", "通用对话不是强项"),
       ("Kimi", "K3", "44", "长文档；Kimi Delta Attention 省显存", "生态与工具链较新"),
       ("MiniMax", "M3 · 海螺 3.0", "29", "多模态与语音，成本低", "综合分差距明显"),
       ("腾讯混元", "Hy4 preview", "—", "元宝、ima、WorkBuddy 一套外壳", "LMArena 智能体第 20；未开视觉"),
       ("百度文心", "文心 5.0 · 5.1", "—", "千帆生态、单位采购路径清楚", "无同口径第三方分数")]


def b_others(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "其余十一家：各自守住一段，也各有一处硬边界",
           "不是排名，是“什么情况下该想到它”。AA 为智能指数 v4.3.2；写“—”表示没有同口径分数。")
    cw = (CW - 0.56) / 2
    LX, RX = ML, ML + cw + 0.56
    vr(s, RX - 0.28, y + 0.02, 4.42, 1.0, LINE)
    for i, (nm, mdl, aa, good, bad) in enumerate(BR2):
        left = i < 6
        x = LX if left else RX
        yy = y + 0.04 + (i if left else i - 6) * 0.700
        hr(s, x, yy, cw, 1.0, LINE)
        text(s, x, yy + 0.10, 1.22, 0.26, nm, 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        text(s, x + 1.26, yy + 0.11, cw - 1.26 - 0.52, 0.24, mdl, 10, GREY, SANS, False,
             PP_ALIGN.LEFT, 1.1, check=False)
        text(s, x + cw - 0.50, yy + 0.09, 0.50, 0.26, aa, 12.5, TEAL, MONO, True, PP_ALIGN.RIGHT, 1.0,
             check=False)
        text(s, x + 0.02, yy + 0.35, cw - 0.04, 0.21, good, 9.8, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.15, check=False)
        text(s, x + 0.02, yy + 0.55, cw - 0.04, 0.21, "边界：" + bad, 9.6, ORANGE, SANS, False,
             PP_ALIGN.LEFT, 1.15, check=False)
    hr(s, LX, y + 0.04 + 6 * 0.700, cw, 1.0, LINE)
    hr(s, RX, y + 0.04 + 5 * 0.700, cw, 1.0, LINE)
    text(s, ML, 6.50, CW, 0.26,
         "选的时候按任务选，不按总分选：图纸与长任务看 Claude，难题看 GPT，中文与知识面看 Gemini，便宜量大看国产开源。",
         10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：各公司模型页与公告、artificialanalysis.ai；2026-10-18 查询")
    return s
