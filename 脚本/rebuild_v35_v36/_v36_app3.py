# -*- coding: utf-8 -*-
"""附录 C / D / E / F / Z"""
from ds import *
from app import sec, term_rows, num_list, BRAND, CH, hd
from app2 import lab, para, kv, BODY
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ---------------- B2b Anthropic 的边界 ----------------
def b_limit(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "Anthropic 的两条线：能干什么，和谁能用",
           "一边是可验证的产出（代码逐帧渲成视频），一边是硬性地区与账号限制。采购前先看右边。")
    # 左：视频实现路径
    LW = 5.60
    sec(s, ML, y + 0.06, LW, "B2b", "它怎么“出片”",
        "写网页动画代码 → 浏览器逐帧截图 → FFmpeg 合成；作品为作者自述，各截 4 秒。")
    IT = [("一句话需求", "15 秒动态图形作品集", "约 37.9 万观看"),
          ("一句话提示词", "中华五千年历史速览", "2 分 38 秒"),
          ("Remotion ＋开源配音", "AI 发展史短片", "3 分钟"),
          ("同一条路径", "轨道中的世界", "约 167.6 万观看")]
    yy = y + 1.10
    for a, b, c in IT:
        hr(s, ML, yy, LW, 1.0, LINE)
        text(s, ML, yy + 0.12, 1.86, 0.26, a, 10, FAINT, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
        text(s, ML + 1.94, yy + 0.09, 2.30, 0.28, b, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, ML + 4.30, yy + 0.11, 1.26, 0.26, c, 10.5, TEAL, MONO, False, PP_ALIGN.RIGHT, 1.0,
             check=False)
        yy += 0.50
    hr(s, ML, yy, LW, 1.0, LINE)
    text(s, ML, yy + 0.14, LW, 0.46,
         "“一句话能出片，但出不了好片”：刷屏作品多半还有长提示词和多轮修改。", 11.5, GREY, SANS,
         False, PP_ALIGN.LEFT, 1.35, check=False)
    # 右：限制时间线
    RX = ML + LW + 0.62
    RW = X2 - RX
    vr(s, RX - 0.31, y + 0.06, 4.10, 1.0, LINE)
    sec(s, RX, y + 0.06, RW, "B2c", "三条红线与封禁规模",
        "按时间顺序；数字来自官方公告与报道。这一页直接决定你能不能长期用。")
    TL = [("2025-09", "中资控股超过 50% 的实体", "不论注册在哪，一律不提供服务"),
          ("2026-02", "公布“蒸馏”报告", "点名 DeepSeek、月之暗面、MiniMax，约 2.4 万个欺诈账号"),
          ("2026-10", "新使用政策（11-12 生效）", "按人所在地、注册地、控股方判断；名单无大陆、港澳"),
          ("2026 上半年", "封禁 1,140 万账号", "申诉 39.8 万，推翻 4.2 万；疑似代理可能要求证件")]
    yy = y + 1.10
    for d, t, b in TL:
        text(s, RX, yy, 1.10, 0.24, d, 10, ORANGE, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX + 1.16, yy - 0.03, RW - 1.16, 0.28, t, 12.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        nl = nlines(b, 10.5, RW - 1.16)
        text(s, RX + 1.16, yy + 0.26, RW - 1.16, nl * 10.5 * LINES * 1.35 / 72 + 0.05, b, 10.5, GREY,
             SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
        yy += 0.28 + nl * 10.5 * LINES * 1.35 / 72 + 0.28
    text(s, RX, 6.02, RW, 0.50,
         "自用与采购都先排除这三条；确认不适用，再谈别的。", 11.5, TEAL, SANS, True, PP_ALIGN.LEFT,
         1.3, check=False)
    footer(s, "来源：anthropic.com 公告与使用政策；technews.tw（2026-10-14）；GitHub opusvideo/awesome-claude-video（作者自述，未独立复现）")
    return s


# ---------------- C1 四件事 + 七步 ----------------
def c1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "动手之前：四件事各查一次，顺序别反",
           "地区不过就别折腾账号，支付不过就别先买会员。四道都过了，再按右边七步跑通一个真实任务。")
    FOUR = [("地区", "先确认：所在地开不开", "定价页搜 supported countries。OpenAI、Anthropic 不支持大陆；Gemini 大陆只开单位账号。不过就别往下走。"),
            ("网络", "再确认：网络通不通", "网页能开、登录不转圈、接口能回结果。只开网页不够：客户端还要连 API 域名。单位网络有管控的先问规定。"),
            ("账号", "然后：能收码的邮箱和手机号", "优先 Google 账号（Gemini 必须用）。手机号要能收国际短信；VoIP、座机多被拒。别买转售账号。"),
            ("支付", "最后：一张能过的卡", "多数只收 Visa / Mastercard，账单地址须在支持地区。国内卡多半过不了：卡在发卡国。国内产品可开票。")]
    LW = 6.90
    yy = y + 0.04
    for i, (t, q, b) in enumerate(FOUR):
        hr(s, ML, yy, LW, 1.0, LINE)
        text(s, ML, yy + 0.13, 0.34, 0.34, f"0{i+1}", 13, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.50, yy + 0.10, 1.00, 0.28, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 1.48, yy + 0.14, 2.10, 0.24, q, 10.5, TEAL, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        nl = nlines(b, 10.5, LW - 1.48)
        text(s, ML + 1.48, yy + 0.38, LW - 1.48, nl * 10.5 * LINES * 1.38 / 72 + 0.05, b, 10.5, BODY,
             SANS, False, PP_ALIGN.LEFT, 1.38, check=False)
        yy += 0.40 + nl * 10.5 * LINES * 1.38 / 72 + 0.16
    hr(s, ML, yy, LW, 1.0, LINE)
    # 右：七步
    RX = ML + LW + 0.56
    RW = X2 - RX
    vr(s, RX - 0.28, y + 0.04, 4.06, 1.0, LINE)
    sec(s, RX, y + 0.04, RW, "C1b", "照着做：从零到跑通第一个任务",
        "做一步，验一步。第 3 步和第 6 步最容易被跳过：不跑真实任务，就测不出它到底行不行。")
    ST = [("选一个能用的产品", "按左边四件事筛"),
          ("注册一个主账号", "一个邮箱一个账号；同一手机号多开会触发风控"),
          ("先用完免费额度", "拿手上真实的小任务跑一遍，别用“你好”测试"),
          ("验证它能读你的文件", "传一份真实 PDF，问某页某行写了什么"),
          ("把任务写成一段说明", "目标、材料、格式、验收标准"),
          ("对照结果，算返工时间", "记下“做多久、错几处”"),
          ("再决定买哪一档", "最低档先跑一周；年付退不了，月付随时停")]
    yy = y + 1.16
    for i, (t, b) in enumerate(ST):
        circ(s, RX + 0.13, yy + 0.13, 0.235, TEAL)
        text(s, RX + 0.055, yy + 0.045, 0.20, 0.20, str(i + 1), 9.5, WHITE, SANS, True,
             PP_ALIGN.CENTER, 1.0, check=False)
        text(s, RX + 0.40, yy + 0.015, RW - 0.40, 0.26, t, 12, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, RX + 0.40, yy + 0.245, RW - 0.40, 0.22, b, 9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.25,
             check=False)
        yy += 0.545
    footer(s, "来源：OpenAI、Anthropic、Google 支持地区页与定价页；各产品帮助中心（2026-10-18 查询）；不含任何绕过地区或身份限制的方法")
    return s


# ---------------- C2/C3 地区与注册 ----------------
def c2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "地区与注册：卡在哪一步，先看清楚",
           "只列事实，不涉及绕过方法。国内发的外币卡为什么刷不了：发卡国是中国，与账单地址要求冲突。")
    rows = [["产品", "大陆能否使用", "收什么付款", "注册 / 登录方式", "对大陆用户的门槛"],
            ["ChatGPT / Codex", "不在支持列表（港澳也不在）", "国际卡；App 内走 App Store / Google Play", "邮箱、Google、Apple；生成 API Key 要验证手机", "大陆不在支持列表"],
            ["Claude / Claude Code", "不支持；11-12 起排除中资控股", "国际卡，发卡国须在支持地区", "Google 账号或邮箱；还要支持地区手机号，一号最多 3 个", "硬门槛：大陆号不行"],
            ["Gemini", "大陆仅单位 Workspace 账号", "Google 账户付款方式", "必须用 Google 账号（Gmail）", "大陆仅单位账号"],
            ["Grok（X 订阅）", "官方未列地区；大陆无法直接访问", "银行卡、Apple / Google Pay", "Google 或 X 账号；X 要验证手机号", "大陆无法直接访问"],
            ["Meta Muse", "人须在开放国家（已确认美、加）", "Meta 账户付款方式", "Meta 账号（邮箱、手机号或 Facebook）", "人须在开放国家"],
            ["Hark", "支持地区官网未写明", "—", "Google、Apple 或邮箱", "—"],
            ["Trae 国际版", "付费不含大陆（国内版用 trae.cn）", "支付宝、微信、银联", "微信、手机号、邮箱、SSO", "无"],
            ["智谱 GLM", "国内产品；海外版 z.ai", "支付宝、微信，可开发票", "手机号＋验证码", "无（依法实名）"],
            ["Kimi / 豆包 / 千问 / DeepSeek", "国内可用", "支付宝、微信", "+86 手机号＋验证码", "无（依法实名）"]]
    table(s, ML, y + 0.10, CW, rows, [2.10, 2.62, 2.72, 2.82, 1.72], size=10.5, head_size=10.5,
          aligns=[PP_ALIGN.LEFT] * 5, minrow=0.335, head_h=0.36, pad=0.10,
          colors=lambda r, c: (NAVY if c == 0 else (ORANGE if c == 4 and r in (2, 4, 6) else BODY)),
          rowlines=True)
    text(s, ML, 6.10, CW, 0.52,
         "建议：海外服务统一用 Google 账号（Gemini 必须用，Claude、ChatGPT 也支持）；国内产品用支付宝 / 微信，能开发票。",
         12, TEAL, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：各产品支持地区与定价页；support.claude.com；help.openai.com；platform.deepseek.com（2026-10-18 查询）")
    return s


# ---------------- C4 价格 ----------------
def c4(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "价格速查：聊天端和任务端，两笔账要分开算",
           "月付，2026-10-18 查询。绿色为人民币，黑色为美元，* 为媒体报道价。买前先看续费价。")
    rows = [["聊天与个人代理", "免费版", "付费（每月）", "任务客户端", "免费额度", "付费（每月）"],
            ["ChatGPT", "GPT-6 Luna 等", "Plus $20 · Pro $100 起", "WorkBuddy", "每月 500 积分", "¥99（连续包月 ¥70）"],
            ["Claude", "Sonnet、Haiku", "Pro $20 · Max $100 起", "Trae 国内版", "500 积分", "Lite ¥49 起"],
            ["Gemini", "Flash-Lite", "Pro $19.99 起", "Qoder", "2 周试用＋300", "$20；国内 ¥59"],
            ["Kimi", "日常对话", "¥49 起，最高 ¥699", "Codex", "免费账号小模型", "Go $8 · Plus $20"],
            ["豆包* / 千问*", "日常对话", "¥68 起 / ¥19 起", "Claude Code", "无", "$20（Claude Pro）"],
            ["Grok", "有限额度", "Plus $8 · $30 起", "Cline", "开源，自带 Key", "ClinePass $9.99"],
            ["Meta Muse", "有用量上限", "$20 · $100", "GLM Coding", "清言网页免费", "Lite ¥118；海外 $18"],
            ["Hark", "Pro 免费有上限", "Pro² $20 · Pro³ $100", "Kiro", "—", "50 credits / $20"]]
    table(s, ML, y + 0.10, CW, rows, [1.90, 2.20, 2.62, 1.94, 2.10, 2.36], size=11, head_size=11,
          aligns=[PP_ALIGN.LEFT] * 6, minrow=0.365, head_h=0.40, pad=0.10,
          colors=lambda r, c: (NAVY if c in (0, 3) else BODY), rowlines=True)
    text(s, ML, 6.22, 6.60, 0.52,
         "口径提醒：积分 / Credits 是客户端自己的计量单位，不等于 token；同样的钱在不同客户端能干的活差很多。",
         11, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    text(s, ML + 7.10, 6.22, 5.02, 0.52,
         "Grok 已并入 X 订阅；GLM Coding 只能在支持的编程工具里用。", 11, GREY, SANS, False,
         PP_ALIGN.LEFT, 1.35, check=False)
    footer(s, "来源：各产品官方定价页（2026-10-18 查询）；豆包、千问为界面新闻报道；meta.com；hark.com")
    return s


# ---------------- C4b/C5 省钱与省流 ----------------
def c5(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "省钱的三条路，代价各不一样",
           "便宜不是捡来的：省下的钱通常换成时间、稳定性或数据风险。判断标准只有一条 —— 这段材料能不能给陌生人看。")
    IT = [("01", "官方活动与首月价", "代价最低，但窗口短",
           "Cline 首月 $1（之后 $9.99）；Codex Go 档 $8/月；opencode Go 档含免费额度；GLM Coding Lite ¥118/月。",
           "窗口只有一个月：先定闹钟，别被自动续费扣住。", MINT),
          ("02", "国内签到与免费积分", "代价是时间，不是钱",
           "WorkBuddy 500 积分＋签到；Trae 每天 100 次（会员 200）；Qoder 2 周试用＋300 次；Codex 免费号只有小模型。",
           "任务不重、能每天手动领一次才划算；赶 deadline 不一定够。", TEALXL),
          ("03", "共享、代充与公益站", "代价最高：数据、稳定、合规",
           "公益站、中转（AnyRouter 一类）；闲鱼共享号 / 代充几元一月；资料要过别人的服务器，维护靠人随时会停。",
           "只跑公开材料；别用于合同、图纸、报告、客户信息。", ORANGEL)]
    yy = y + 0.06
    for no, t, tag, body, warn, fill in IT:
        box(s, ML, yy, 7.20, 1.28, fill)
        text(s, ML + 0.22, yy + 0.19, 0.40, 0.28, no, 13, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.70, yy + 0.14, 3.10, 0.30, t, 15, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1,
             check=False)
        text(s, ML + 0.70, yy + 0.46, 3.10, 0.24, tag, 10, ORANGE if no == "03" else TEAL, SANS, True,
             PP_ALIGN.LEFT, 1.0, check=False)
        text(s, ML + 0.70, yy + 0.72, 6.30, 0.24, body, 10, BODY, SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
        text(s, ML + 0.70, yy + 0.98, 6.30, 0.24, "→ " + warn, 10, GREY, SANS, False, PP_ALIGN.LEFT,
             1.3, check=False)
        yy += 1.40
    RX = ML + 7.66
    RW = X2 - RX
    vr(s, RX - 0.30, y + 0.06, 4.10, 1.0, LINE)
    sec(s, RX, y + 0.06, RW, "C5", "省流版：按你能折腾到哪一步选",
        "从左往右门槛和花费逐级上升；先在左边跑通一个真实任务，再决定要不要往右走。")
    STEP = [("1  最快 · 免费", "WorkBuddy · Trae", "国内手机号直接用，人民币支付；内置 DeepSeek、GLM、混元、豆包 Seed 等；每月 500 积分加每日签到。"),
            ("2  能解决海外网络", "Cline 或 Hark", "Cline：低价套餐额度大，只含开源模型；Hark：Pro 免费档有用量上限，盲测 D 就是 Hark 里的 Opus 5.5。"),
            ("3  想用顶配", "GPT", "推理、数学榜第一是 GPT-6 Astra（P24）；复杂分析、长推理交给它；各档位可用模型以官网为准。")]
    yy = y + 1.14
    for t, who, b in STEP:
        text(s, RX, yy, RW, 0.24, t, 10.5, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
             check=False)
        text(s, RX, yy + 0.22, RW, 0.28, who, 14, TEAL, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        nl = nlines(b, 10.5, RW)
        text(s, RX, yy + 0.52, RW, nl * 10.5 * LINES * 1.4 / 72 + 0.05, b, 10.5, BODY, SANS, False,
             PP_ALIGN.LEFT, 1.4, check=False)
        yy += 0.56 + nl * 10.5 * LINES * 1.4 / 72 + 0.24
    footer(s, "来源：各产品定价与活动页；L 站公益与共享服务的社区反馈；trae.cn/pricing；workbuddy.cn；docs.cline.bot；hark.com/pricing")
    return s


# ---------------- D 资讯渠道 ----------------
def d1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "四个渠道：社区看风向，仓库看版本，视频看演示，X 看一手",
           "遇到问题先去 L 站搜、再去 GitHub 搜报错原文。这四个都不是结论来源，结论仍要回到官方与自己的测试。")
    # 顶部三个数字
    N = [("4.6 万", "30 天新话题"), ("92 万", "30 天新帖子"), ("11.6 万", "30 天活跃用户")]
    for i, (b, c) in enumerate(N):
        x = ML + i * 1.86
        text(s, x, y + 0.02, 1.70, 0.44, b, 30, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, y + 0.50, 1.70, 0.22, c, 9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.1, check=False)
    text(s, ML + 5.70, y + 0.06, 6.56, 0.62,
         "L 站：邀请注册制，邮箱只收白名单域名。新模型、限额变化常比媒体早；报错、封号、“降智”基本都有人问过答过。",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    hr(s, ML, y + 0.90, CW, 1.0, LINE2)
    CH4 = [("D2  GitHub", "工具更新的第一现场", [("更新记录", "每次更新改了什么，以官方写的为准"), ("Issues", "搜报错原文，比问 AI 更快"), ("Trending", "最近几天大家在关注什么项目"), ("早报存档", "imjuya/juya-ai-daily，Markdown 与 RSS")]),
           ("D3  B 站", "两个方向：实测与资讯", [("Token就是词元", "“屎山论剑”拿真实祖传代码让模型修 bug；新模型基本都测"), ("橘鸦Juya", "每日 AI 早报，合集 570 期、累计播放 1700 万以上；daily.juya.uk")]),
           ("D4  X", "一手发布和榜单更新最快", [("厂商官方", "@OpenAI、@AnthropicAI、@GoogleDeepMind、@xai、@deepseek_ai、@Alibaba_Qwen"), ("榜单与评测", "@ArtificialAnlys、@lmarena_ai、@EpochAIResearch、@arcprize")])]
    yy = y + 1.06
    for i, (t, sub, items) in enumerate(CH4):
        text(s, ML, yy, 2.60, 0.30, t, 16, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        text(s, ML, yy + 0.32, 2.60, 0.24, sub, 9.5, TEAL, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        iy = yy
        for a, b in items:
            text(s, ML + 2.86, iy, 1.72, 0.24, a, 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.1,
                 check=False)
            nl = nlines(b, 10.5, 7.10)
            text(s, ML + 4.66, iy - 0.02, 7.10, nl * 10.5 * LINES * 1.35 / 72 + 0.05, b, 10.5, BODY,
                 SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
            iy += 0.22 + nl * 10.5 * LINES * 1.35 / 72 + 0.06
        hr(s, ML, yy + max(1.20, iy - yy) + 0.10, CW, 1.0, LINE)
        yy += max(1.20, iy - yy) + 0.30
    text(s, ML, 6.30, CW, 0.42,
         "三条提醒：帖子是个人体感，结论要自己复核；有账号、邀请码、“羊毛”类内容不要参与；大陆无法直接访问 X，转载常有延迟和误传。",
         11, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：linux.do/about；github.com/trending；bilibili.com 各 UP 主空间与视频页；x.com 各官方账号（2026-10-18 查询）")
    return s
