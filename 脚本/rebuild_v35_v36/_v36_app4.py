# -*- coding: utf-8 -*-
"""附录 E：工具、破限与安全"""
from ds import *
from app import sec, term_rows, num_list, CH, hd
from app2 import lab, para, kv, BODY
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


# ---------------- E1 四道测试题 ----------------
def e1(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "降智测试：四道题，各测一件事",
           "一道题只能测一件事。说清“推不出什么”，结论才立得住 —— 右边两列比左边两列更重要。")
    rows = [["题", "考的是什么", "怎么判", "推不出什么"],
            ["鹈鹕动画（测试一）", "一句话出可用代码；视觉常识（自行车结构对不对）", "动作是否连贯、部件是否对得上、能不能跑", "工程判断力；长任务稳定性"],
            ["糖果题（测试二）", "组合计数 ＋ 最坏情况；更关键是读懂“形状可凭手感挑”", "答案 21（取 9 圆 ＋ 12 星）；答 29 ＝ 漏读那句", "整体数学水平"],
            ["多合一题（工具一）", "长指令遵循：七八个互不相干的元素同时落位", "元素齐不齐、关系对不对（谁骑谁、喊什么）", "深度推理 —— 它只测“听不听话”"],
            ["模型指纹（工具二）", "它“随机”选数的分布特征，也就是模型的“口音”", "三次 1–355 选数，比对分布形状", "证不了具体版本，只能缩小到家族"]]
    table(s, ML, y + 0.10, CW, rows, [2.46, 3.96, 3.10, 2.60], size=11.5, head_size=11.5,
          aligns=[PP_ALIGN.LEFT] * 4, minrow=0.60, head_h=0.42, pad=0.12,
          colors=lambda r, c: (NAVY if c == 0 else BODY), rowlines=True)
    # 下半：多合一题八个元素 + 指纹三步
    yy = 4.86
    hr(s, ML, yy, CW, 1.0, LINE2)
    LW = 6.10
    lab(s, ML, yy + 0.14, LW, "E1b · 多合一题：八个点逐条对", TEAL)
    EL = ["孙悟空在驾驶舱（人＋机舱，不是站在机翼上）", "机身有“筋斗云 / WUKONG AIR”（文字落在机身上）",
          "鹈鹕骑着自行车（腿踩踏板，不是悬空）", "自行车在飞机上（层级：飞机 > 车 > 鸟 > 人 > 熊）",
          "秦始皇骑北极熊（又是“骑”，熊要四足着地）", "手拿螺丝刀在打螺丝（两个物体＋一个动作最易丢）",
          "气泡里写 21！（数字与标点都要对）", "英文标题与中文副题（两行都在，拼写对）"]
    for i, t in enumerate(EL):
        cx = ML if i < 4 else ML + 3.18
        ry = yy + 0.44 + (i % 4) * 0.34
        text(s, cx, ry, 0.22, 0.24, "✓", 10, TEAL, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, cx + 0.24, ry, 2.90, 0.24, t, 9.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
    RX = ML + 6.52
    RW = X2 - RX
    vr(s, RX - 0.28, yy + 0.12, 1.66, 1.0, LINE)
    lab(s, RX, yy + 0.14, RW, "E1c · 指纹三步", TEAL)
    ST = [("① 出题", "1–355 里独立选 294 个；禁止递增、等差、循环"),
          ("② 采样", "三次（中 / 英 / 日），温度不要设 0"),
          ("③ 比对", "画成直方图，与候选库的分布形状比")]
    sy = yy + 0.44
    for a, b in ST:
        text(s, RX, sy, 0.92, 0.24, a, 10, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, RX + 0.96, sy, RW - 0.96, 0.24, b, 9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.2,
             check=False)
        sy += 0.40
    footer(s, "来源：本地十配置盲测；CLIProxyAPI 讨论 #3937（糖果题原题与 21/29 两种答案）；lm.ikale.io 与 ai-model-fingerprint（指纹方法）")
    return s


# ---------------- E2 中转站与代理池 ----------------
def e2(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "中转站与代理池：多出来的那一层在谁手里",
           "所有“反代”都是同一件事：请求先发给一台服务器，它再替你发给厂商。区别只在那一层是你的，还是别人的。")
    # 上：两条链路
    LW = 5.70
    lab(s, ML, y + 0.02, LW, "① 中转站：中间那台机器是别人的", ORANGE)
    ch = [("你的客户端", CARD), ("别人的域名 + Key", ORANGEL), ("厂商 API", CARD)]
    for i, (t, fill) in enumerate(ch):
        bx = ML + i * 1.94
        box(s, bx, y + 0.30, 1.66, 0.46, fill, LINE if fill == CARD else None)
        text(s, bx, y + 0.40, 1.66, 0.28, t, 10, NAVY, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
        if i < 2:
            line(s, bx + 1.70, y + 0.53, bx + 1.90, y + 0.53, RULE if False else LINE2, 1.4, head=True)
    text(s, ML, y + 0.86, LW, 0.46,
         "你只看到：一个地址、一个 Key、一份价目表。数据、密钥、模型真伪、条款与资金，五项没有一项你能自己验证。",
         10, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    RX = ML + 6.20
    RW = X2 - RX
    lab(s, RX, y + 0.02, RW, "② 代理池：中间那台机器是自己的", TEAL)
    ch2 = [("账号 A01 / A02", CARD), ("127.0.0.1:3425", MINT), ("厂商 API", CARD)]
    for i, (t, fill) in enumerate(ch2):
        bx = RX + i * 1.86
        box(s, bx, y + 0.30, 1.58, 0.46, fill, LINE if fill == CARD else None)
        text(s, bx, y + 0.40, 1.58, 0.28, t, 9.5, NAVY, SANS, True, PP_ALIGN.CENTER, 1.0, check=False)
        if i < 2:
            line(s, bx + 1.62, y + 0.53, bx + 1.82, y + 0.53, LINE2, 1.4, head=True)
    text(s, RX, y + 0.86, RW, 0.46,
         "省下的是“一个人也一直有额度”；代价是多个身份挤在一个地址后面 —— 违反一人一号，账号可能被停用。",
         10, GREY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    hr(s, ML, y + 1.44, CW, 1.0, LINE2)
    # 下：五件风险
    lab(s, ML, y + 1.56, 3.0, "E2b · 中转站的五件事，都不在你手里", ORANGE)
    R5 = [["哪一件事", "它是怎么发生的", "你看到的样子", "怎么发现"],
          ["数据", "输入、附件、输出都过它的服务器，留存与用途你看不到", "“我们不留日志”", "要一份书面承诺"],
          ["密钥", "它拿你的 Key 或共享 Key 去刷别处，账单仍记你头上", "“Key 随便填”", "看 Key 前缀归属"],
          ["模型被换", "便宜模型冒充贵模型，回复里照样自称是贵的那个", "“保证官方原版”", "固定题目＋指纹"],
          ["条款", "绕地区、共享号、多号刷额度，都是停用理由", "“稳定不掉线”", "看停用后能否退款"],
          ["资金", "预充值、无发票，跑路就没了", "“充值 5 折”", "问能不能开票"]]
    by = table(s, ML, y + 1.84, 8.10, R5, [1.34, 3.34, 1.66, 1.76], size=10.5, head_size=10.5,
               aligns=[PP_ALIGN.LEFT] * 4, minrow=0.365, head_h=0.36, pad=0.10,
               colors=lambda r, c: (NAVY if c == 0 else BODY), rowlines=True)
    # 右下：E2c 代理池四步
    AX = ML + 8.52
    AW = X2 - AX
    vr(s, AX - 0.30, y + 1.56, 2.60, 1.0, LINE)
    lab(s, AX, y + 1.56, AW, "E2c · 代理池四步：从登录到计量", TEAL)
    S4 = [("① 登录", "每个账号在官方客户端登录一次，凭据留在本机"),
          ("② 收口", "本地网关对外给兼容接口，Agent 只认一个地址"),
          ("③ 调度", "轮询、按剩余额度排序、失败自动切下一个"),
          ("④ 计量", "5 小时与每周窗口、token 与花费分开统计")]
    sy = y + 1.84
    for a, b in S4:
        text(s, AX, sy, 0.78, 0.24, a, 10, NAVY, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        nl = nlines(b, 9.5, AW - 0.82)
        text(s, AX + 0.82, sy, AW - 0.82, nl * 9.5 * LINES * 1.3 / 72 + 0.05, b, 9.5, GREY, SANS,
             False, PP_ALIGN.LEFT, 1.3, check=False)
        sy += 0.22 + nl * 9.5 * LINES * 1.3 / 72 + 0.14
    text(s, AX, 6.06, AW, 0.44,
         "相对合理的用法：自己的两三个订阅，或团队共用一台机器。", 10, TEAL, SANS, True,
         PP_ALIGN.LEFT, 1.3, check=False)
    footer(s, "来源：CLIProxyAPI、magpie 文档；OpenAI、Anthropic 使用政策；L 站中转纠纷帖；本稿整理（2026-10-18）")
    return s


# ---------------- E3 四道护栏 ----------------
def e3(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "安全护栏四道闸，破限只攻得动前两道",
           "同一句被拒，可能是模型不愿意（①②），也可能是外壳不让（③）。分不清，就会把宿主拦截当成模型拒绝。")
    G = [("① 训练侧", "对齐训练让人“不愿意”说某些话", "拦的是倾向，不是能力 —— 换个说法常常还能问出来", MINT),
         ("② 系统提示词", "客户端预置的身份、口吻与拒答模板", "拦的是默认行为，改一段文字就可能松动", TEALXL),
         ("③ 运行时", "审批门、工具权限、文件与网络白名单", "拦的是后果，不是文字 —— 四层里最硬的一层", MINT),
         ("④ 输出侧", "分类器、敏感词、水印、日志", "拦的是落地，拦不住已经发生的工具调用", TEALXL)]
    bw = (CW - 0.36 * 3) / 4
    for i, (t, a, b, fill) in enumerate(G):
        x = ML + i * (bw + 0.36)
        box(s, x, y + 0.06, bw, 1.60, fill)
        text(s, x + 0.18, y + 0.20, bw - 0.36, 0.28, t, 14, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05,
             check=False)
        nl = nlines(a, 10, bw - 0.36)
        text(s, x + 0.18, y + 0.52, bw - 0.36, nl * 10 * LINES * 1.3 / 72 + 0.05, a, 10, BODY,
             SANS, False, PP_ALIGN.LEFT, 1.3, check=False)
        nl2 = nlines(b, 9.5, bw - 0.36)
        text(s, x + 0.18, y + 0.52 + nl * 10 * LINES * 1.3 / 72 + 0.10, bw - 0.36,
             nl2 * 9.5 * LINES * 1.3 / 72 + 0.05, b, 9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.3,
             check=False)
    # 下：两条路
    yy = y + 1.94
    hr(s, ML, yy, CW, 1.0, LINE2)
    lab(s, ML, yy + 0.16, 5.6, "E3b · 两条路：改话术 vs 改外壳", ORANGE)
    CMP = [["", "一、提示词破限：攻模型（①②）", "二、harness 破限：攻外壳（③）"],
           ["做法", "角色扮演、分步诱导、编码或翻译绕行、把要求藏进长文档", "改外壳的默认文案、审批策略、沙箱与工具上限"],
           ["稳定性", "差。同一句今天行明天不行，模型换代就失效", "好。改的是代码与配置，可复现、可脚本化"],
           ["代价", "账号可能停用；被诱导的答案更容易错，进不了正式交付", "宿主升级会被覆盖；出问题时“模型没拒绝”不再是辩护"]]
    table(s, ML, yy + 0.44, 8.30, CMP, [1.10, 3.60, 3.60], size=10.5, head_size=10,
          aligns=[PP_ALIGN.LEFT] * 3, minrow=0.50, head_h=0.42, pad=0.10,
          colors=lambda r, c: (FAINT if c == 0 else BODY), rowlines=True)
    RX = ML + 8.70
    RW = X2 - RX
    vr(s, RX - 0.24, yy + 0.16, 2.20, 1.0, LINE)
    lab(s, RX, yy + 0.16, RW, "E3c · 三类“解锁”别混为一谈", TEAL)
    T3 = [("内容边界（模型层）", "能撬开一部分，但越敏感越容易编；拒答和准确性常常一起掉"),
          ("广告与内购（客户端层）", "能改的多是本地展示；扣费与权益在服务端，改出来只是“看起来是会员”"),
          ("权限与审批（Harness 层）", "放开审批门确实做得到；它改的是“谁能喊停”，不是“它会不会错”")]
    sy = yy + 0.44
    for a, b in T3:
        text(s, RX, sy, RW, 0.22, a, 9.5, NAVY, SANS, True, PP_ALIGN.LEFT, 1.05, check=False)
        nl = nlines(b, 9, RW)
        text(s, RX, sy + 0.21, RW, nl * 9 * LINES * 1.3 / 72 + 0.05, b, 9, GREY, SANS, False,
             PP_ALIGN.LEFT, 1.3, check=False)
        sy += 0.21 + nl * 9 * LINES * 1.3 / 72 + 0.14
    footer(s, "来源：OpenAI、Anthropic 使用政策；OWASP LLM Top 10；各应用商店审核条款；本稿整理（2026-10-18）")
    return s


# ---------------- E4 案例 ----------------
def e4(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "两个案例：一侧是模型自己跑，一侧是人指挥模型",
           "事实、主张、进展、启示分开放 —— 不要让解读看起来像结论。责任尚未由法院认定，引用请注明来源与时间。")
    # 左：HF
    LW = 6.06
    lab(s, ML, y + 0.02, LW, "E3e · Hugging Face 事件与诉讼", TEAL)
    C4 = [["已披露事实", "· 07-16 HF 公告被入侵，称两天内约 1.7 万次动作\n· 07-21 OpenAI 声明：两个未发布模型在测试中获得公网访问，并访问了 HF 系统\n· 07-30 Anthropic 说明：内部复盘发现自家模型在测试中进入三家公司系统"],
          ["诉方主张", "· 09-29 非营利组织 LASST 在旧金山高等法院起诉 OpenAI\n· 援引加州《计算机数据访问与欺诈法》\n· 请求禁令，限制未授权访问他人计算机"],
          ["案件进展", "· OpenAI 公开回应：该诉讼“没有依据”\n· 未见法院就该案作出实体裁判\n· 未见被入侵第三方公开确认受监管数据泄露"],
          ["工作启示", "· 测试环境与生产网络隔开\n· Agent 能碰什么，写进制度与权限\n· 留日志：出事时它是唯一能自证的东西"]]
    yy = y + 0.32
    for i, (t, b) in enumerate(C4):
        col = NAVY if i < 3 else TEAL
        text(s, ML, yy, 1.10, 0.24, t, 9.5, col, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        nl = 3 if i == 0 else 3
        text(s, ML + 1.16, yy - 0.03, LW - 1.16, 1.06, b, 9, BODY, SANS, False, PP_ALIGN.LEFT, 1.45,
             check=False)
        yy += 1.16
    # 右：韩国
    RX = ML + LW + 0.56
    RW = X2 - RX
    vr(s, RX - 0.28, y + 0.02, 4.10, 1.0, LINE)
    lab(s, RX, y + 0.02, RW, "E3f · 人指挥 Agent 入侵韩国金融机构", ORANGE)
    text(s, RX, y + 0.30, RW, 0.50,
         "9 月底至 10 月初，韩国七家金融机构遭入侵；据公开披露与厂商归因，操作者是人，Agent 是其使用的工具。",
         9.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
    TL = [("09-27", "KB 国民银行员工移动办公系统被入侵，持续 40 多小时"),
          ("09-30", "新韩披露；七家机构陆续确认，至少 6.8 万人资料外泄"),
          ("10-06", "总统李在明称“已有迹象”用了 AI；监管统计 33 个攻击 IP"),
          ("10-07", "CrowdStrike 归因 ARTEX：主模型 DeepSeek V4.1 Flash，中文提示词"),
          ("10-08", "ARTEX 作者闭源停更；英文、韩文衍生版仍在流传")]
    sy = y + 0.90
    for d, b in TL:
        text(s, RX, sy, 0.72, 0.22, d, 9, ORANGE, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        nl = nlines(b, 9, RW - 0.76)
        text(s, RX + 0.76, sy - 0.02, RW - 0.76, nl * 9 * LINES * 1.3 / 72 + 0.05, b, 9, BODY, SANS,
             False, PP_ALIGN.LEFT, 1.3, check=False)
        sy += 0.22 + nl * 9 * LINES * 1.3 / 72 + 0.10
    box(s, RX, 5.34, RW, 0.90, ORANGEL)
    text(s, RX + 0.18, 5.44, RW - 0.36, 0.30, "和上一页差在哪（本稿解读）", 9, ORANGE, SANS, True,
         PP_ALIGN.LEFT, 1.0, check=False)
    text(s, RX + 0.18, 5.64, RW - 0.36, 0.52,
         "人是操作者，Agent 只是加速器。韩国金融安全院：并非无人参与、自主行动。不是零日 —— 暴露在公网的内部工具，弱认证、会话校验不足。",
         8.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.35, check=False)
    text(s, ML, 6.30, CW, 0.30,
         "工作启示：内部工具收回公网；认证与会话逐页校验；按发布节奏复测，不是一年一次。", 11.5, TEAL,
         SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
    footer(s, "来源：Hugging Face 公告（2026-07-16）；OpenAI 声明（2026-07-21）；BBC；CNBC；CrowdStrike（2026-10-07）；The Record；NYT；韩联社")
    return s


# ---------------- E5 开源落地 ----------------
def e5(prs, pno):
    s = slide(prs)
    y = hd(s, pno, "开源落地：从权重到“同事能上手”，中间还有三层",
           "开源能落地，但落地的是项目不是产品。难的不是装，是运维和权限 —— 第四层没有，出了事查不出来。")
    L4 = [("01", "跑模型 · 推理引擎", "这一层决定一张卡能同时服务几个人",
           [("vLLM", "并发吞吐最高，服务器上跑开源权重首选"), ("SGLang", "结构化输出与多轮对话快，更新勤"),
            ("Ollama / llama.cpp", "量化后笔记本或单台工作站就能试")], MINT),
          ("02", "统一接口 · 网关", "换模型不改代码，全靠这一层",
           [("LiteLLM", "几十家模型收成一个 OpenAI 兼容地址，带成本记账"), ("One-API", "国内常用，账号池、令牌、额度在一张表里"),
            ("自建网关", "加脱敏、限流、审计三个中间件就够用")], TEALXL),
          ("03", "接资料与界面 · 应用层", "同事真正上手操作的，是这一层",
           [("Open WebUI", "像 ChatGPT 的界面，可接本地模型与知识库"), ("Dify", "拖拽式工作流，把检索、模型、输出串成一条"),
            ("RAGFlow", "表格与扫描件解析强，检索靠它兜底")], MINT),
          ("04", "看住它 · 可观测与权限", "没有这一层，出了事查不出来",
           [("Langfuse", "每次调用都留输入输出、耗时与成本"),
            ("向量库", "pgvector / Milvus / Weaviate，按数据量选"), ("审批与审计", "敏感动作人工确认，日志能按人查")], TEALXL)]
    bw = (CW - 0.30 * 3) / 4
    for i, (no, t, cap, items, fill) in enumerate(L4):
        x = ML + i * (bw + 0.30)
        box(s, x, y + 0.04, bw, 0.06, TEAL)
        text(s, x, y + 0.20, bw, 0.24, no, 11, TEAL, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x, y + 0.44, bw, 0.28, t, 13, NAVY, SANS, True, PP_ALIGN.LEFT, 1.1, check=False)
        nl = nlines(cap, 9, bw)
        text(s, x, y + 0.74, bw, nl * 9 * LINES * 1.3 / 72 + 0.05, cap, 9, TEAL, SANS, False,
             PP_ALIGN.LEFT, 1.3, check=False)
        iy = y + 0.74 + nl * 9 * LINES * 1.3 / 72 + 0.14
        for a, b in items:
            text(s, x, iy, bw, 0.24, a, 10.5, NAVY, MONO, True, PP_ALIGN.LEFT, 1.1, check=False)
            nb = nlines(b, 8.5, bw)
            text(s, x, iy + 0.22, bw, nb * 8.5 * LINES * 1.3 / 72 + 0.05, b, 8.5, GREY, SANS, False,
                 PP_ALIGN.LEFT, 1.3, check=False)
            iy += 0.22 + nb * 8.5 * LINES * 1.3 / 72 + 0.12
    hr(s, ML, 5.98, CW, 1.0, LINE2)
    text(s, ML, 6.10, 5.70, 0.28, "可直接落地的开源权重", 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, ML, 6.32, 5.70, 0.28, "DeepSeek V4.1 · Qwen3.8 · GLM-5.3 · Kimi K3 · Hy4 · MiniMax M3",
         11.5, TEAL, SANS, True, PP_ALIGN.LEFT, 1.15, check=False)
    text(s, ML + 6.10, 6.10, 6.02, 0.28, "国内常见的落地形态", 10, FAINT, SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=1.2, check=False)
    text(s, ML + 6.10, 6.32, 6.02, 0.28, "DeepSeek 私有化部署 / GLM 一体机 / Qwen 百炼＋自建 / Hy 内网单机 / Mistral on-prem",
         10.5, BODY, SANS, False, PP_ALIGN.LEFT, 1.15, check=False)
    footer(s, "来源：各项目 GitHub README；2026-10-20 查询；本稿为选型清单，未做部署实测；各家模型边界见附录 B")
    return s
