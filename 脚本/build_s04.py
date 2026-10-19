# -*- coding: utf-8 -*-
"""重写 04 节的六页：RAG、三类易错、工作流 vs Agent、多 Agent、企业落地。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

from build_v3 import (put, rect, hline, vline, dot, chrome, section_bar,
                      callout, rowband, node, style_run, L, W,
                      C_TITLE, C_BODY, C_MUTED, C_TEAL, C_ORANGE, C_HEAD,
                      C_LINE, C_WHITE, C_WARM, C_COOL, C_ZEBRA)


def clear(slide):
    spTree = slide.shapes._spTree
    for child in list(spTree):
        if child.tag in (qn("p:sp"), qn("p:cxnSp"), qn("p:pic"),
                         qn("p:graphicFrame"), qn("p:grpSp")):
            spTree.remove(child)


# ================================================================ 1. RAG
def page_rag(slide, page):
    chrome(
        slide, page,
        "资料一多，为什么先检索再回答（RAG）",
        "两个理由都很实在：一次省一个量级的 token，一次少一批凭空编出来的答案",
        "来源：Chroma《Context Rot》；Lewis 等 RAG 原论文（2020）；DeepSeek、OpenAI 上下文计费文档",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页回答“为什么要用”，不是“怎么搭”。两个理由按重要性排。\n\n"
            "第一，省 token——而且是省一个量级。一百页中文资料大约五到八万 token，"
            "整包塞进去，问一次付一次；追问十次就付十次。检索只把命中的三到五段（一两千 token）送进去，"
            "差二十到五十倍。这一条在按量计费的 API 上直接是钱。\n\n"
            "第二，少编。模型有个稳定的毛病：材料里没有，它不会说“没有”，会按“最像”往下续。"
            "把原文片段摆到它眼前，它就有东西可抄，还能标出页码；再加上一句"
            "“只依据片段回答，找不到就说找不到”，编造率会明显下降。\n\n"
            "原理五步：入库（切分、向量化、建索引，一次性）、提问（问题也向量化）、"
            "检索（算相似度取前几段）、拼装（片段＋问题＋约束）、输出（结论＋出处）。"
            "关键点：RAG 不训练模型，也不把资料记进模型，它每次现找现给——"
            "所以资料更新了，检索结果立刻跟着变。\n\n"
            "边界要讲清楚：检索没找到，它照样会编。所以“找不到就说找不到”这句必须写进提示词，"
            "而且要在验收时抽查。\n\n"
            "最后那句衔接下一页：这一步只保证“找得到”，找到之后读错、切错、算错，才是真正的坑。"
        ))

    # 左：两个理由
    section_bar(slide, L, 1.90, 6.10, "为什么要用：两个理由", h=0.38, size=13.5)

    put(slide, 0.72, 2.34, 5.90, 0.28, [
        ("理由一　", 13.5, True, C_ORANGE),
        ("省 token：不是省一点，是省一个量级", 13.5, True, C_TITLE)],
        anchor=MSO_ANCHOR.TOP)
    lines1 = [
        ("整包塞", "100 页资料 ≈ 5–8 万 token，问一次付一次"),
        ("只检索", "命中 3–5 段 ≈ 1–2 千 token"),
        ("差距", "20–50 倍；追问十次，差出一整个月的额度"),
    ]
    y = 2.66
    for k, v in lines1:
        put(slide, 0.86, y, 0.95, 0.26, k, size=11.5, bold=True, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 1.86, y, 4.70, 0.26, v, size=12, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        y += 0.27

    put(slide, 0.72, 3.62, 5.90, 0.28, [
        ("理由二　", 13.5, True, C_ORANGE),
        ("少编：把“凭印象”换成“照着原文”", 13.5, True, C_TITLE)],
        anchor=MSO_ANCHOR.TOP)
    lines2 = [
        ("材料不在眼前", "模型只能按“最像”续写，数字、条文、页码最容易编"),
        ("原文进上下文", "有东西可抄，还能标出处，人一眼能核"),
        ("再加一句约束", "“只依据片段回答，找不到就写没找到”——编造率明显下降"),
    ]
    y = 3.94
    for k, v in lines2:
        put(slide, 0.86, y, 1.20, 0.26, k, size=11.5, bold=True, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.10, y, 4.46, 0.46, v, size=11.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.1)
        y += 0.52

    put(slide, 0.72, 5.52, 5.90, 0.40,
        "注意：检索只负责“找得到”。没找到还硬答，它照样会编——所以那句话要写进规则，也要抽查。",
        size=10.5, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # 右：原理五步
    section_bar(slide, 6.95, 1.90, 5.78, "原理：先找，再答", color=C_HEAD,
                h=0.38, size=13.5)
    steps = [
        ("① 入库", "切分 → 向量化 → 建索引（只做一次，资料变了才重做）", C_TEAL),
        ("② 提问", "问题也转成向量", C_TEAL),
        ("③ 检索", "算相似度，取最像的 3–5 段", C_TEAL),
        ("④ 拼装", "片段 ＋ 问题 ＋ “只依据片段回答”", C_ORANGE),
        ("⑤ 输出", "结论 ＋ 文件名 / 页码，能逐条回原文", C_ORANGE),
    ]
    vline(slide, 7.16, 2.42, 3.06, C_LINE, 1.0)
    y = 2.40
    for k, v, col in steps:
        dot(slide, 7.16, y + 0.13, 0.12, col)
        put(slide, 7.42, y, 5.20, 0.26, k, size=12, bold=True, color=col,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 7.42, y + 0.26, 5.20, 0.34, v, size=11, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.1)
        y += 0.62

    rect(slide, 6.95, 5.36, 5.78, 0.52, C_COOL)
    put(slide, 7.12, 5.36, 5.46, 0.52,
        "不训练模型，也不把资料记进模型：每次现找现给，资料改了立刻生效",
        size=11, color=C_BODY, spacing=1.15)

    callout(slide, "一句话", "RAG 买的是“少付 token”和“少编”，不是“让模型更聪明”", kind="cool")


# ================================================================ 2. 三处易错
def page_three_errors(slide, page):
    chrome(
        slide, page,
        "找得到之后，还会错在三处：读、切、算",
        "这三处正好对上前面两个案例——图纸与修编报告，错都不在“模型不聪明”",
        "来源：本地两案例的原始输出与人工复核记录；S31 提示注入；OCR 与表格结构说明",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页的作用是承上启下。上一页说 RAG 只保证“找得到”；这一页说，找到之后还有三关。\n\n"
            "第一关，读取。扫描件漏字、OCR 把 3 认成 8、图里的标注根本没进模型。"
            "图纸那个案例里，有两个配置把“单位 mm/cm”当成了错误——这不是读错，是没读懂图例；"
            "但也有配置确实漏读了 β(°) 那一列。检查动作：核页数、表名、行数，"
            "图纸要先问清是“看图”还是“只 OCR”。\n\n"
            "第二关，结构。跨页表头、合并单元格、页间断裂。修编报告里那个 112 元差额，"
            "就是汇总表口径和分项汇总口径不一致造成的——数字都对，摆错了位置。"
            "检查动作：检查行列归属和页间连续性，别只看合计。\n\n"
            "第三关，计算。把文本数字当数值、单位没换算、把公式抄成结果。"
            "透水管那条 13,723 m → 7,032.785 m，核减 71.8663 万元，是可以独立复算的；"
            "复算不出来就说明有问题。检查动作：打开公式自己重算关键合计。\n\n"
            "收尾要落到“所以下一步是分工”：读取交给 OCR、结构交给解析脚本、计算交给代码，"
            "模型只负责组织和判断。这句话直接接下一天的工作流页。"
        ))

    # 顶部链条：RAG 之后还有三关
    rect(slide, L, 1.90, W, 0.40, C_COOL)
    put(slide, 0.78, 1.90, 11.80, 0.40, [
        ("上一页 RAG 只解决“找得到”　→　", 12, False, C_BODY),
        ("读进来要准　·　摆进去要对　·　算出来要能复算", 12, True, C_TEAL),
        ("　这三关，模型一关都不保证", 12, False, C_BODY)],
        spacing=1.0)

    cols = [
        ("读取", "扫描件漏字、OCR 认错、图没进模型", [
            ("图纸案例", "有配置把 mm/cm 当错误（没读图例）；也有真漏读 β(°) 列"),
        ], "核页数、表名、行数；先确认是“看图”还是“只 OCR”", C_TEAL),
        ("结构", "跨页表头、合并单元格错位、页间断行", [
            ("报告案例", "汇总表口径 vs 分项汇总，差 112 元：数字都对，摆错了位置"),
        ], "查行列归属与页间连续性；合计对不上先看结构", C_ORANGE),
        ("计算", "文本数字、单位没换算、公式抄成结果", [
            ("报告案例", "13,723 m → 7,032.785 m，核减 71.8663 万元，可独立复算"),
        ], "打开公式自己重算关键合计；算不出来的数不进正文", C_HEAD),
    ]
    x = L
    for name, sym, case, check, edge in cols:
        cw = 3.85
        rect(slide, x, 2.44, cw, 3.30, C_WHITE)
        rect(slide, x, 2.44, cw, 0.06, edge)
        put(slide, x + 0.22, 2.58, cw - 0.44, 0.34, name, size=20, bold=True,
            color=C_TITLE)
        put(slide, x + 0.22, 2.96, cw - 0.44, 0.52, sym, size=11.5, color=C_MUTED,
            spacing=1.2, anchor=MSO_ANCHOR.TOP)
        hline(slide, x + 0.22, 3.56, cw - 0.44, C_LINE)
        put(slide, x + 0.22, 3.64, cw - 0.44, 0.24, case[0][0], size=10.5, bold=True,
            color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
        put(slide, x + 0.22, 3.88, cw - 0.44, 0.66, case[0][1], size=10.5,
            color=C_BODY, spacing=1.2, anchor=MSO_ANCHOR.TOP)
        put(slide, x + 0.22, 4.62, cw - 0.44, 0.24, "检查动作", size=10.5, bold=True,
            color=C_TEAL, anchor=MSO_ANCHOR.TOP)
        put(slide, x + 0.22, 4.86, cw - 0.44, 0.72, check, size=10.5, color=C_BODY,
            spacing=1.2, anchor=MSO_ANCHOR.TOP)
        x += cw + 0.26

    rect(slide, L, 5.86, W, 0.40, C_ZEBRA)
    put(slide, 0.78, 5.86, 11.80, 0.40,
        "所以下一步不是“换更强的模型”，而是分工：读取交 OCR、结构交解析脚本、计算交代码，模型只做组织与判断",
        size=11.5, color=C_TITLE, spacing=1.0)

    callout(slide, "结论", "错得最贵的三处，都在“模型之外”：读、切、算")


# ================================================================ 3. 工作流 vs Agent
def page_workflow(slide, page):
    chrome(
        slide, page,
        "一条流水线里，哪几步该写死，哪几步该留给模型",
        "拿修编报告那条真实流程举例：五步里只有两步需要判断，其余三步写成代码更快更准",
        "来源：本稿修编报告案例的实际执行过程；S09–S16 各模型客户端说明",
        notes=(
            "讲解与会后阅读\n\n"
            "上一页说要分工，这一页给分工的具体方法：拿一条真实流程，逐步标“写死”还是“判断”。\n\n"
            "例子就是修编报告那条链：① 读清单（Excel 解析，格式固定）——写死；"
            "② 定位变更条目（要在三个文件之间比对，还要判断哪条算变更）——判断；"
            "③ 套模板填表（模板固定，位置固定）——写死；"
            "④ 复算金额（可独立复算，112 元差额就是这么发现的）——写死；"
            "⑤ 写原因文字（要读函件、组织语言）——判断。\n\n"
            "判断标准三条：输入输出格式固定、有唯一正确答案、出错代价高——三条都满足就写死。"
            "反过来，材料不全要临场判断、输出是自然语言、步骤数事先说不清——就该交给模型。\n\n"
            "混合版的实际效果：脚本三秒算完而且可复算，模型只写两段话；"
            "返工从四十分钟降到五分钟。这不是估计，是那条流程跑下来的实际差别。\n\n"
            "要强调：不是“Agent 更高级所以都用 Agent”。Agent 的价值在探索，"
            "在固定流程上用 Agent，等于每次让它重新摸索一遍，还很难复现。"
        ))

    # 左：流程五步
    section_bar(slide, L, 1.90, 6.55, "以“修编报告”为例：五步，两种做法", h=0.38, size=13.5)
    steps = [
        ("① 读清单", "Excel 解析，行列固定", "写死", C_TEAL),
        ("② 定位变更条目", "跨三个文件比对，判断哪条算变更", "判断", C_ORANGE),
        ("③ 套模板填表", "模板固定，单元格固定", "写死", C_TEAL),
        ("④ 复算金额", "可独立复算，差额 112 元由此发现", "写死", C_TEAL),
        ("⑤ 写原因文字", "要读函件、组织表述", "判断", C_ORANGE),
    ]
    y = 2.38
    for k, v, tag, col in steps:
        rowband(slide, L, y, 6.55, 0.56, 0 if col == C_TEAL else 1)
        put(slide, 0.74, y, 1.55, 0.56, k, size=12.5, bold=True, color=C_TITLE)
        put(slide, 2.34, y, 3.40, 0.56, v, size=10.5, color=C_BODY, spacing=1.1)
        rect(slide, 5.82, y + 0.13, 1.18, 0.30, col)
        put(slide, 5.82, y + 0.13, 1.18, 0.30, tag, size=10.5, bold=True,
            color=C_WHITE, align=PP_ALIGN.CENTER)
        y += 0.56

    put(slide, 0.74, 5.26, 6.40, 0.44,
        "脚本 3 秒算完且可复算；模型只写两段话。返工从 40 分钟降到 5 分钟。",
        size=11, color=C_ORANGE, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # 右：判断标准
    section_bar(slide, 7.35, 1.90, 5.38, "怎么判：三条就够", color=C_ORANGE,
                h=0.38, size=13.5)
    put(slide, 7.55, 2.38, 5.05, 0.26, "写死（脚本 / 工作流）", size=11.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    hard = ["输入输出格式固定", "有唯一正确答案", "出错代价高、需要能复算"]
    y = 2.66
    for t in hard:
        dot(slide, 7.62, y + 0.11, 0.10, C_TEAL)
        put(slide, 7.78, y, 4.80, 0.24, t, size=11, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        y += 0.26

    put(slide, 7.55, 3.56, 5.05, 0.26, "交给模型（Agent）", size=11.5, bold=True,
        color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
    soft = ["材料不全，要临场判断", "输出是自然语言", "步骤数事先说不清"]
    y = 3.84
    for t in soft:
        dot(slide, 7.62, y + 0.11, 0.10, C_ORANGE)
        put(slide, 7.78, y, 4.80, 0.24, t, size=11, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        y += 0.26

    rect(slide, 7.35, 4.78, 5.38, 1.04, C_WARM)
    put(slide, 7.55, 4.84, 5.00, 0.26, "两种都用的混合版长这样", size=11, bold=True,
        color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
    put(slide, 7.55, 5.12, 5.00, 0.64,
        "输入检查 → 计算 → 审批 → 交付：走固定流程；中间的“这算不算变更”“原因怎么写”：交给模型",
        size=10.5, color=C_BODY, spacing=1.2, anchor=MSO_ANCHOR.TOP)

    callout(slide, "结论", "固定流程上用 Agent，等于每次重新摸索一遍，还很难复现", kind="warn")


# ================================================================ 4. 多 Agent
def page_multi_agent(slide, page):
    chrome(
        slide, page,
        "多个 Agent 一起干：什么时候值，什么时候只是更贵",
        "收益来自并行和上下文隔离；代价来自协调、文件冲突，以及“同一个错误前提”",
        "来源：本稿多 Agent 协作记录；Subagent 与并行任务说明；S09–S16 客户端文档",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页补的是“优缺点与适用条件”，不是“怎么搭”。\n\n"
            "先说真实收益，两条：并行——三个 Agent 各查一份清单的不同章节，"
            "原来串行四十分钟的活十几分钟能出初稿；上下文隔离——"
            "每个 Agent 只看自己那一摊，主线上下文不被几十页材料冲淡，"
            "这正好接前面讲的上下文腐烂。\n\n"
            "再说代价，三条，而且都真实发生过：成本翻倍（三个 Agent 就是三份 token）；"
            "写冲突（两个 Agent 改同一个文件，后写的把前面的覆盖掉）；"
            "最要命的是同一个错误前提——三个都用了错的汇总口径，就会一起错，"
            "而且看起来像“三方交叉验证过”，比一个 Agent 错得更可信。\n\n"
            "适用条件四个，缺一就别拆：任务能切成互不依赖的块、每块有独立可查的产出、"
            "写作用范围不重叠、有人负责合并与验收。\n\n"
            "什么时候别拆：步骤前后强依赖（第二步要用第一步的结论）、"
            "总量本来就小（拆完协调成本比干活还多）、没有验收人。\n\n"
            "结论那句要讲透：人多不自动变可靠。复核的人如果只是复述执行者的答案，"
            "那等于没复核——必须回到原材料。"
        ))

    # 左：收益
    section_bar(slide, L, 1.90, 6.05, "什么时候值：两条真实收益", h=0.38, size=13.5)
    gains = [
        ("并行", "三个 Agent 各查一份清单的不同章节：串行 40 分钟的活，十几分钟出初稿"),
        ("上下文干净", "每个只看自己那一摊，主线不被几十页材料冲淡（接上下文腐烂那条）"),
    ]
    y = 2.40
    for k, v in gains:
        rect(slide, L, y, 6.05, 0.80, C_WHITE)
        put(slide, 0.78, y + 0.08, 1.20, 0.26, k, size=12.5, bold=True, color=C_TEAL)
        put(slide, 0.78, y + 0.36, 5.65, 0.40, v, size=10.5, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.88

    # 左：代价
    section_bar(slide, L, 4.18, 6.05, "什么时候只是更贵：三条真实代价", color=C_ORANGE,
                h=0.38, size=13.5)
    costs = [
        ("成本翻倍", "三个 Agent = 三份 token，还得加上协调的往返"),
        ("写冲突", "两个改同一个文件，后写的把前面的覆盖掉"),
        ("错误前提传染", "三个都用了错的汇总口径 → 一起错，还像“交叉验证过”"),
    ]
    y = 4.64
    for k, v in costs:
        put(slide, 0.78, y, 1.55, 0.24, k, size=10.5, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.38, y, 4.20, 0.24, v, size=10.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        y += 0.32

    # 右：适用条件
    section_bar(slide, 6.95, 1.90, 5.78, "拆之前先过这四条（缺一就别拆）", color=C_HEAD,
                h=0.38, size=13.5)
    conds = [
        "任务能切成互不依赖的块",
        "每块有独立、可核对的产出",
        "写文件的范围不重叠",
        "有人负责合并并验收",
    ]
    y = 2.42
    for i, t in enumerate(conds, 1):
        put(slide, 7.12, y + 0.02, 0.40, 0.26, str(i), size=12, bold=True, color=C_TEAL)
        put(slide, 7.52, y, 5.05, 0.26, t, size=11.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        y += 0.34

    rect(slide, 6.95, 3.86, 5.78, 0.86, C_WARM)
    put(slide, 7.14, 3.92, 5.40, 0.24, "这几种就别拆", size=10.5, bold=True,
        color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
    put(slide, 7.14, 4.18, 5.40, 0.50,
        "步骤前后强依赖（第二步要用第一步的结论）· 总量本来就小 · 没有验收人",
        size=10.5, color=C_BODY, spacing=1.2, anchor=MSO_ANCHOR.TOP)

    rect(slide, 6.95, 4.84, 5.78, 0.82, C_COOL)
    put(slide, 7.14, 4.90, 5.40, 0.24, "最常见的一种假复核", size=10.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    put(slide, 7.14, 5.16, 5.40, 0.44,
        "复核的人只复述执行者的答案，等于没复核；必须回到原材料逐条对",
        size=10.5, color=C_BODY, spacing=1.2, anchor=MSO_ANCHOR.TOP)

    callout(slide, "结论", "人多不自动变可靠：拆得开、验得了、有人兜底，才值得拆")


# ================================================================ 5. 企业落地
def page_enterprise(slide, page):
    chrome(
        slide, page,
        "企业里怎么落地：先定一个能验收的场景，再补架构",
        "五层图只是把已有能力摆到该在的位置；真正决定成败的是第一个场景、验收标准和负责人",
        "来源：通用企业 LLM 应用架构示意；S28 网信办备案公告；本地案例的实施记录",
        notes=(
            "讲解与会后阅读\n\n"
            "上一版这页只画了五层，容易变成空谈。这一版把“怎么开始”讲具体。\n\n"
            "先说五层各自的最小可用版本——注意是“最小”，不是“完整”："
            "数据层先做一份整理好的、脱敏过的样本目录；能力层先用云 API，别一上来私有化；"
            "编排层先写死一条流程，不要一上来就 Agent；应用层先做一个场景；"
            "入口层先用企业微信或网页，别做插件。\n\n"
            "再说九十天怎么走，三段：\n"
            "0–2 周：选题。从“有人工答案可对照”的小活开始，比如一页资料摘要、一张表复核。"
            "交付物是一页任务书：目标、材料、边界、交付格式、验收标准。\n"
            "3–6 周：跑通一条。固定十道测试题，每次都跑；记录模型、客户端、设置、原始输出。"
            "交付物是一份能复现的样例包。\n"
            "7–12 周：加一层。要么加权限与审计，要么加第二个场景。"
            "每加一层都先跑一遍那十道题，分数掉了就回退。\n\n"
            "三个常见失败：一上来做“全公司知识库”（范围太大，没人能验收）；"
            "没有验收标准（只能说“感觉还行”）；没有 owner（出事没人认，也没人推）。\n\n"
            "预算上给个量级：一个场景跑三个月，API 费用通常在几百到几千元；"
            "贵的是人，不是 token。这句话能帮听众判断该投入在哪。"
        ))

    # 左：五层 + 最小可用版本
    section_bar(slide, L, 1.90, 6.15, "五层架构：每层的“最小可用版本”", h=0.38, size=13.5)
    layers = [
        ("入口层", "先用企业微信 / 网页，不做插件"),
        ("应用层", "只做一个场景：资料问答 或 表格复核"),
        ("编排层", "先写死一条流程，别上 Agent"),
        ("能力层", "先用云 API；私有部署留到有硬要求时"),
        ("数据层", "先整理一份脱敏样本目录，不求全"),
    ]
    y = 2.40
    for i, (k, v) in enumerate(layers):
        rowband(slide, L, y, 6.15, 0.54, i)
        rect(slide, L, y, 0.07, 0.54, C_TEAL)
        put(slide, 0.80, y, 1.30, 0.54, k, size=11.5, bold=True, color=C_TITLE)
        put(slide, 2.14, y, 4.50, 0.54, v, size=10.5, color=C_BODY, spacing=1.1)
        y += 0.54

    put(slide, 0.80, 5.18, 5.95, 0.46,
        "顺序反了会很难看：先买私有化部署、再想做什么场景，是这类项目最常见的死法。",
        size=10.5, color=C_ORANGE, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # 右：90 天
    section_bar(slide, 7.05, 1.90, 5.68, "90 天：三段，每段都有交付物", color=C_HEAD,
                h=0.38, size=13.5)
    phases = [
        ("0–2 周", "选题", "从“有人工答案可对照”的小活开始：一页摘要、一张表复核", 
         "交付：一页任务书（目标/材料/边界/格式/验收）", C_TEAL),
        ("3–6 周", "跑通一条", "固定 10 道测试题，每次跑；记模型、客户端、设置、原始输出",
         "交付：能复现的样例包", C_TEAL),
        ("7–12 周", "加一层", "加权限与审计，或加第二个场景；每加一层先重跑那 10 道题",
         "交付：验收记录 + 回退方案", C_ORANGE),
    ]
    y = 2.40
    for k, t, v, d, col in phases:
        rect(slide, 7.05, y, 5.68, 0.92, C_WHITE)
        rect(slide, 7.05, y, 0.06, 0.92, col)
        put(slide, 7.20, y + 0.06, 1.05, 0.24, k, size=10.5, bold=True, color=col,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 8.28, y + 0.06, 1.20, 0.24, t, size=11, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 9.50, y + 0.06, 3.10, 0.40, v, size=9.5, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        put(slide, 7.20, y + 0.56, 5.35, 0.30, d, size=9.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        y += 1.00

    # 底部：三个失败
    rect(slide, L, 6.02, W, 0.26, C_ZEBRA)
    put(slide, 0.78, 6.02, W - 0.36, 0.26, [
        ("三个常见失败：", 10.5, True, C_ORANGE),
        ("一上来做“全公司知识库”（没人能验收）　·　没有验收标准（只能说“感觉还行”）　·　没有 owner（出事没人认）",
         10.5, False, C_BODY)], spacing=1.0)

    callout(slide, "结论", "先定一个能验收的小场景，每加一层都先跑一遍那 10 道题")


# ================================================================ assembly
def main(src, dst):
    prs = Presentation(src)

    jobs = [
        (43, page_rag, 45),
        (44, page_three_errors, 46),
        (45, page_workflow, 47),
        (46, page_multi_agent, 48),
        (48, page_enterprise, 50),
    ]
    for idx, fn, page in jobs:
        footer = None
        for sh in prs.slides[idx].shapes:
            if sh.has_text_frame and sh.left is not None:
                if abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.25:
                    footer = sh.text_frame.text.strip()
        clear(prs.slides[idx])
        fn(prs.slides[idx], footer or page)
        print("rebuilt", idx + 1, "->", footer)

    prs.save(dst)
    print("saved", dst, len(prs.slides))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
