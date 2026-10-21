# -*- coding: utf-8 -*-
"""附录 A–Z：速查与出处"""
from ds import *
from layouts import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.dml import MSO_LINE_DASH_STYLE as DASH


# ---------------- 通用小部件 ----------------
def sec(s, x, y, w, tag, name, desc=None, size=17):
    """小节头：字母编号 + 名称 + 一句说明"""
    text(s, x, y, 0.62, 0.26, tag, 12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.2, check=False)
    text(s, x + 0.66, y - 0.06, w - 0.66, 0.34, name, size, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
         check=False)
    yy = y + 0.34
    if desc:
        nl = nlines(desc, 12, w)
        text(s, x, yy, w, nl * 12 * LINES * 1.35 / 72 + 0.06, desc, 12, GREY, SANS, False,
             PP_ALIGN.LEFT, 1.35, check=False)
        yy += nl * 12 * LINES * 1.35 / 72 + 0.06
    return yy


def term_rows(s, x, y, w, items, ts=13.5, es=12, ns=10.5, tw=1.42, ew=2.62, gap=0.06,
              rule=True, hpad=0.10):
    """术语三栏：术语 / 一句解释 / 例子与注意；返回底部 y"""
    nw = w - tw - ew
    yy = y
    for i, (t, e, n) in enumerate(items):
        ne = nlines(e, es, ew - hpad)
        nn = nlines(n, ns, nw - hpad)
        h = max(ne * es * LINES * 1.30 / 72, nn * ns * LINES * 1.30 / 72, ts * 1.30 / 72) + 0.15
        if rule:
            hr(s, x, yy - 0.02, w, 1.0, LINE)
        text(s, x, yy + 0.045, tw - 0.08, h, t, ts, NAVY, SANS, True, PP_ALIGN.LEFT, 1.22,
             check=False)
        text(s, x + tw, yy + 0.055, ew - hpad, h, e, es, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.30, check=False)
        text(s, x + tw + ew, yy + 0.075, nw, h, n, ns, TEAL, SANS, False, PP_ALIGN.LEFT, 1.30,
             check=False)
        yy += h
    return yy


def num_list(s, x, y, w, items, no_w=0.46, ts=15, bs=12.5, gap=0.16, col=TEAL, tcol=NAVY):
    """编号条目：01 + 标题 + 正文"""
    yy = y
    for no, t, b in items:
        text(s, x, yy + 0.02, no_w, 0.26, no, 12, col, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        nl = nlines(t, ts, w - no_w, True)
        text(s, x + no_w, yy - 0.02, w - no_w, nl * ts * LINES * 1.2 / 72 + 0.06, t, ts, tcol, SANS,
             True, PP_ALIGN.LEFT, 1.2, check=False)
        yy += nl * ts * LINES * 1.2 / 72 + 0.04
        if b:
            nb = nlines(b, bs, w - no_w)
            text(s, x + no_w, yy + 0.02, w - no_w, nb * bs * LINES * 1.35 / 72 + 0.06, b, bs, GREY,
                 SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
            yy += nb * bs * LINES * 1.35 / 72
        yy += gap
    return yy


def hd(s, pno, ttl, sub, size=28):
    return head(s, CH, pno, ttl, sub, size=size)


CH = "附录  速查与出处"


# ---------------- 01 附录导览 ----------------
def a_index(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "附录：七块速查，一条条查得到",
           "正文里每处判断都能回到这里：术语怎么定义、各家模型什么边界、多少钱、去哪看、原理与出处。页码为本次修订后的新页码。")
    IT = [("A", "常用术语", "36 条：模型本身 / 怎么交代任务 / Agent 与工具 / 工作方式与评测", 55),
          ("B", "公司档案", "13 家：模型线、外壳、强项、边界；含数学进展、降智、地区限制", 57),
          ("C", "订阅与省流", "地区与支付、注册门槛、价格速查、省钱三条路、省流版三档", 64),
          ("D", "资讯渠道", "L 站、GitHub、B 站、X：看什么、别信什么", 68),
          ("E", "分享与工具", "降智四题、中转站与代理池、四道护栏、开源落地", 69),
          ("F", "原理与历史", "发展史、学习四步、token 与向量、注意力、架构、规模定律、多模态", 74),
          ("Z", "参考资料来源", "按主题编号的来源表，与逐条链接清单", 83)]
    yy = y + 0.02
    for tag, nm, desc, pg in IT:
        h = 0.56
        hr(s, ML, yy, CW, 1.0, LINE)
        text(s, ML, yy + 0.07, 0.60, 0.44, tag, 27, TEALL if tag == "Z" else TEAL, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.72, yy + 0.12, 2.20, 0.30, nm, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 3.02, yy + 0.15, 7.30, 0.28, desc, 12, GREY, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, X2 - 0.80, yy + 0.15, 0.80, 0.26, f"P{pg}", 11, FAINT, MONO, False,
             PP_ALIGN.RIGHT, 1.0, check=False)
        yy += h
    hr(s, ML, yy, CW, 1.0, LINE)
    text(s, ML, yy + 0.12, CW, 0.28,
         "读法：正文按“是什么 → 怎么做 → 怎么核”推进；附录按“查得到”组织，查到就停，不必顺序读。",
         12.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：本稿整理；页码为本次修订后的页序")
    return s


# ---------------- A 术语 ----------------
A1 = [("Token（词元）", "模型读写文字的最小单位，计费和上下文都按它算", "汉字与 token 不一一对应，按各家分词器算"),
      ("上下文窗口", "一次对话里模型能“看到”的 token 上限", "1M 约一百万 token；放得下 ≠ 记得住"),
      ("参数", "模型内部可调的数字，决定能力上限与运行成本", "参数多不等于更聪明"),
      ("MoE（混合专家）", "模型拆成多个“专家”，每次只调用一部分", "总参数大、单次计算小，所以又快又便宜"),
      ("推理 / 思考模式", "先在内部推一遍再作答", "想得越久越准，也越贵越慢"),
      ("多模态", "能同时处理文字、图片、音频、视频", "看图 ≠ OCR：前者理解画面，后者只认字"),
      ("幻觉", "编出不存在的内容，还说得很肯定", "引用、数字、条文最容易出，必须核"),
      ("开放权重 / 开源", "公开模型参数可下载 / 连训练代码数据也公开", "DeepSeek、Qwen、GLM 多为开放权重"),
      ("蒸馏 / 量化", "用大模型教小模型 / 降低数字精度让模型变小", "本地部署多用量化版，能力有折损")]

A2 = [("Prompt（提示词）", "交给模型的任务、资料和要求", "不是咒语；说清目标与验收标准更重要"),
      ("系统提示词", "客户端预先写给模型的“岗位说明”", "同一模型换客户端效果不同，原因之一"),
      ("上下文工程", "决定把哪些资料、历史、规则放进上下文", "比反复改提示词更影响结果"),
      ("RAG（检索增强）", "先从资料库检出相关片段，再让模型作答", "检索漏了，回答就漏"),
      ("Embedding（向量）", "把文字变成一串数字，用来算“意思相近”", "知识库检索的底层"),
      ("API / API Key", "程序调用模型的接口；Key 是计费凭证", "Key 泄露＝别人花你的钱"),
      ("按量 vs 订阅", "API 按 token 付费；订阅按月给额度", "积分 / Credits 是客户端自己的计量单位"),
      ("缓存命中", "重复的前文不重复计费或打折", "长对话前文保持不变更省钱"),
      ("限流（429）", "每分钟、每天能调用的次数或 token 上限", "报 429 就是触发限流，等一会儿再试")]

A3 = [("Agent（智能体）", "模型自己拆步骤、调工具、看结果再决定下一步", "不只是自动连续聊天"),
      ("Harness（外壳）", "包住模型的运行环境：工具、上下文、权限、反馈", "同一模型换外壳，表现差很多"),
      ("工具调用", "模型按固定格式请求执行某个函数", "读文件、搜网页、跑代码都靠它"),
      ("MCP", "把外部工具和数据接给模型的开放协议", "像统一接口；不会让模型变聪明"),
      ("Skill（技能）", "打包好的任务说明、脚本和资料，用到才加载", "不等于自动获得工具权限"),
      ("Subagent（子代理）", "主 Agent 把子任务交给独立上下文的小 Agent", "主线上下文更干净，还能并行"),
      ("多 Agent", "多个 Agent 分工协作", "关键是有人负责验收"),
      ("沙箱（Sandbox）", "隔离的运行环境，Agent 在里面跑代码、改文件", "出错不伤本机；权限仍要收紧"),
      ("Computer Use", "模型看屏幕截图，控制鼠标键盘操作电脑", "能用没有接口的软件；慢，要盯着")]

A4 = [("Browser Use", "专门操作网页的 Agent", "填表、比价、下单"),
      ("Plan / Ask 模式", "只出方案 / 只问答，不直接动手改", "先 Plan 再让它执行更稳"),
      ("Vibe Coding", "只描述需求、让 AI 写代码，不逐行看", "做原型很快，交付前要审"),
      ("工作流（Workflow）", "预先固定的步骤链，模型只负责其中几步", "稳定、可复现，适合重复任务"),
      ("人工确认", "关键步骤停下来等人点头（Human-in-the-loop）", "付款、发信、删文件前必须停"),
      ("提示注入", "网页或文件里藏指令，劫持 Agent", "读外部内容的 Agent 都要防"),
      ("Benchmark（榜单）", "用固定题目给模型打分", "只代表那一类题（见 P24）"),
      ("Arena / Elo", "匿名对战、人来投票排名", "测的是偏好，受回答风格影响"),
      ("降智", "用户感觉同一个模型变笨了", "官方承认过 bug（见附录 B）")]


def a1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "术语（一）：模型本身，和把任务交出去",
           "左边九个词说的是“它是什么”，右边九个词说的是“你怎么把活交给它”。查到这里就够用，不必记。")
    cw = (CW - 0.72) / 2
    sec(s, ML, y + 0.10, cw, "A1", "模型本身", "Token、上下文、参数、MoE、推理模式、多模态、幻觉、开放权重、蒸馏与量化")
    term_rows(s, ML, y + 1.06, cw, A1)
    x2 = ML + cw + 0.72
    vr(s, x2 - 0.36, y + 0.10, 4.10, 1.0, LINE)
    sec(s, x2, y + 0.10, cw, "A2", "怎么把任务交给它", "提示词、系统提示词、上下文工程、RAG、向量、API、计费、缓存、限流")
    term_rows(s, x2, y + 1.06, cw, A2)
    footer(s, "来源：各厂商文档；modelcontextprotocol.io；本稿整理")
    return s


def a2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "术语（三）：Agent 与工具，以及怎么评它",
           "左边九个词是这两年被问得最多的；右边九个词是判断“它到底行不行”时会用到的口径。")
    cw = (CW - 0.72) / 2
    sec(s, ML, y + 0.10, cw, "A3", "Agent 与工具", "Agent、Harness、工具调用、MCP、Skill、Subagent、多 Agent、沙箱、Computer Use")
    term_rows(s, ML, y + 1.06, cw, A3)
    x2 = ML + cw + 0.72
    vr(s, x2 - 0.36, y + 0.10, 4.10, 1.0, LINE)
    sec(s, x2, y + 0.10, cw, "A4", "工作方式与评测", "浏览器代理、Plan 模式、Vibe Coding、工作流、人工确认、提示注入、榜单、Arena、降智")
    term_rows(s, x2, y + 1.06, cw, A4)
    footer(s, "来源：各厂商文档；OWASP LLM Top 10；本稿整理")
    return s


# ---------------- B 公司档案 ----------------
BRAND = [("OpenAI", "GPT-6 Astra · 6.1 Sol · 6 Luna", "53", "全闭源", "推理与数学第一梯队；Codex 四种形态", "$10 / $50 每百万 token；档位最多", "s060_2.jpg"),
         ("Anthropic", "Opus 5.5 · Sonnet 5.5 · Haiku 5.5", "58", "全闭源", "长任务不跑偏、会自己查证", "无免费版；地区限制最硬", "s063_2.jpg"),
         ("Google", "Gemini 4 Argon · 3.8 Flash", "53", "全闭源", "知识面、中文、长文写作；单次 1M 输出", "智能体榜偏后；工具链换过一代", "s066_2.jpg"),
         ("Meta", "Muse Wonder · Wonder Full", "52", "闭源", "长期记忆与个人事务（Active Memory）", "只对美加开放；编程不是主场", "s068_2.jpg"),
         ("xAI", "Grok 4.7 · Grok Bot", "46", "不开源", "常驻型 Agent：每个 Bot 一台云端电脑", "地区限制；常驻意味着交出账号权限", "s067_2.jpg"),
         ("Mistral", "Mistral Large 4", "—", "开放权重", "约 1T MoE、1M 上下文；欧洲合规", "生态与教程少于前四家", "s069_2.jpg"),
         ("DeepSeek", "V4.1 Flash", "39", "开放权重", "最便宜、出字最快；开源五个底层库", "先大段思考；复杂判断仍要人核", "s070_2.jpg"),
         ("Qwen", "Qwen3.8 Max · Flash", "45", "开放权重", "尺寸梯队最全（0.6B→2.4T）；衍生模型最多", "旗舰与顶尖仍差约 10 分", "s071_2.jpg"),
         ("GLM", "GLM-5.3 · 5.3-Flash", "45", "开放权重", "编程提升 50%；Terminal-Bench 开源 SOTA", "通用对话不是它的强项", "s072_2.jpg"),
         ("Kimi", "K3", "44", "开放权重", "2.8T 参数、1M 上下文；长文档", "生态与工具链较新", "s073_2.jpg"),
         ("MiniMax", "M3 · 海螺 3.0", "29", "开放权重", "多模态与语音；成本低", "综合分差距明显", "s074_2.jpg"),
         ("腾讯混元", "Hy4 preview", "—", "不开源", "770B / 49B 激活；元宝、ima、WorkBuddy", "LMArena 智能体第 20；视觉默认未启用", "s075_2.jpg"),
         ("百度文心", "文心 5.0 · 5.1", "—", "不开源", "预训练成本约为同规模 6%；千帆生态", "无同口径第三方分数", "s076_2.jpg")]


MX = [("OpenAI", "GPT-6 Astra \u00b7 6.1 Sol \u00b7 6 Luna", "53", "全闭源",
       "推理与数学第一梯队；Codex 四种形态", "$10 / $50 每百万 token"),
      ("Anthropic", "Opus 5.5 \u00b7 Sonnet 5.5 \u00b7 Haiku 5.5", "58", "全闭源",
       "长任务不跑偏、会自己查证", "无免费版；地区限制最硬"),
      ("Google", "Gemini 4 Argon \u00b7 3.8 Flash", "53", "全闭源",
       "知识面、中文、长文；单次 1M 输出", "智能体榜偏后；不反对你"),
      ("Meta", "Muse Wonder \u00b7 Wonder Full", "52", "闭源",
       "长期记忆与个人事务", "只对美加；编程不是主场"),
      ("xAI", "Grok 4.7 \u00b7 Grok Bot", "46", "不开源",
       "常驻型 Agent：每个 Bot 一台云电脑", "常驻＝交出账号权限"),
      ("Mistral", "Mistral Large 4", "\u2014", "开放权重",
       "约 1T MoE、1M 上下文；欧洲合规", "生态与教程偏少"),
      ("DeepSeek", "V4.1 Flash", "39", "开放权重",
       "最便宜、出字最快；五个底层库开源", "先大段思考；复杂判断要核"),
      ("Qwen", "Qwen3.8 Max \u00b7 Flash", "45", "开放权重",
       "尺寸梯队 0.6B\u21922.4T 最全", "旗舰与顶尖差约 10 分"),
      ("GLM", "GLM-5.3 \u00b7 5.3-Flash", "45", "开放权重",
       "编程提升 50%；开源 SOTA", "通用对话不是强项"),
      ("Kimi", "K3", "44", "开放权重",
       "2.8T 参数、1M 上下文；长文档", "生态与工具链较新"),
      ("MiniMax", "M3 \u00b7 海螺 3.0", "29", "开放权重",
       "多模态与语音；成本低", "综合分差距明显"),
      ("\u817e\u8baf\u6df7\u5143", "Hy4 preview", "\u2014", "不开源",
       "770B / 49B 激活；元宝、ima、WorkBuddy", "智能体榜第 20；默认未开视觉"),
      ("\u767e\u5ea6\u6587\u5fc3", "\u6587\u5fc3 5.0 \u00b7 5.1", "\u2014", "不开源",
       "预训练成本约同规模 6%；千帆生态", "无同口径第三方分数")]


def b_matrix(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "十三家模型公司：一张表看完各自的战场与边界",
           "分数取 AA 智能指数 v4.3.2（2026-10-18 查询）；“—”表示没有同口径分数，不作排名依据。")
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
             [{"runs": [(good + "　", {"size": 8.8, "color": C(0x2B, 0x34, 0x45)}),
                        ("边界：" + bad, {"size": 8.8, "color": ORANGE})], "ls": 1.02}])
    hr(s, LX, y + 0.02 + 7 * P, cw, 1.0, LINE)
    hr(s, RX, y + 0.02 + 6 * P, cw, 1.0, LINE)
    text(s, ML, 6.46, CW, 0.24,
         "看法：分数只能排出“谁更聪明”，排不出“什么活该交给谁”。下两页展开 OpenAI 与 Anthropic，其余十一家在第 63 页。",
         10.2, TEAL, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：artificialanalysis.ai；各公司模型页与公告（2026-10-18 查询）；本稿整理")
    return s


def b_top(prs, pno, which):
    D = BRAND[0] if which == 0 else BRAND[1]
    nm, mdl, aa, w, good, bad, img = D
    s = slide(prs)
    if which == 0:
        ttl = "OpenAI：三个档位，各守一段战场"
        sub = "Astra 打最难的题，Sol 干日常的活，Luna 守免费入口 —— 拿免费版的体验评价它，会低估它。"
        body = [("模型线", "GPT-6 Astra（旗舰，AA 53；FrontierMath Erdős 68 题里唯一得分，3%）；GPT-6.1 Sol（AA 52，每道题 $0.72，性价比最高的一档）；GPT-6 Luna（免费档小模型，上下文 27K）。全部闭源。"),
                ("外壳", "ChatGPT 网页 / 桌面 / 手机；Codex 一个名四种形态 —— 终端、IDE 插件、桌面、云端，Astra 版 Codex 把 computer use 提速约 1.9 倍；另有 API 与 Azure。"),
                ("强项", "推理与数学第一梯队；写代码偏“防御式”，兜底与异常处理写得多，改动面也大。本次鹈鹕 G 一次成片，图纸 1/2（命中 K1）。"),
                ("边界", "贵：Astra 输入 / 输出 $10 / $50 每百万 token。档位最多，免费版和付费版常常不是同一个模型。")]
        src = "来源：openai.com 模型页与系统卡、artificialanalysis.ai、epoch.ai/frontiermath；2026-10-18 查询"
    else:
        ttl = "Anthropic：长任务与工程交付最强的一档"
        sub = "贵，但一次做对的概率高。交付物要经得起人核、又不想反复返工时选它。"
        body = [("模型线", "Claude Opus 5.5（2026-09 发布，AA 58 全场第一；$4 / $20 每百万 token，缓存读 $0.20，比 Opus 5 便宜 40%）；Sonnet 5.5（56）、Haiku 5.5（43）；Fable 5.1（53，上一代旗舰，仍在跑）。全部闭源。"),
                ("外壳", "Claude Code（终端 Agent，社区公认与 Opus 5.5 最搭的一套外壳）；Claude 桌面版（原 Cowork，直接读写本机文件）；claude.ai；API 与 AWS / GCP / Azure。"),
                ("强项", "长任务不跑偏、会自己查证。本次图纸 2/2（十家里唯一全中）；修编报告 4 分钟出稿、金额全对，还联网核对了承建单位。"),
                ("边界", "没有免费版，Claude Pro $20 起；对中国大陆用户与中资实体有地区限制（见下一页）；最强档跑一道题约 $6。")]
        src = "来源：anthropic.com、platform.claude.com 定价、artificialanalysis.ai；2026-10-18 查询"
    y = hd(s, pno, ttl, sub)
    text(s, ML, y + 0.06, 1.20, 0.24, "AA 智能指数", 10.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, ML, y + 0.26, 1.60, 0.72, aa, 54, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 0.04, y + 0.92, 1.60, 0.24, "全场第一" if which else "第一梯队", 10.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.0, check=False)
    LX, LW = ML + 1.86, 6.10
    yy = y + 0.02
    for lab, txt in body:
        text(s, LX, yy, LW, 0.24, lab, 10.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, cs=1.4,
             check=False)
        nl = nlines(txt, 12.5, LW)
        hh = nl * 12.5 * LINES * 1.42 / 72 + 0.06
        text(s, LX, yy + 0.24, LW, hh, txt, 12.5, C(0x2B, 0x34, 0x45), SANS, False, PP_ALIGN.LEFT, 1.42,
             check=False)
        yy += 0.24 + hh + 0.14
    RX, RW = X2 - 2.20, 2.20
    frame(s, RX, y + 0.02, RW, 4.16, LINE, 1.0, CARD)
    pic(s, img, RX + 0.10, y + 0.12, RW - 0.20, 3.96, mode="fill")
    footer(s, src)
    return s
