# -*- coding: utf-8 -*-
"""按反馈重写 04 节：RAG 加相似度图解、工作流两页合并重做、企业落地改成真实路线。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

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


def dashed_circle(slide, cx, cy, r, color=C_ORANGE):
    sh = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(cx - r), Inches(cy - r), Inches(2 * r), Inches(2 * r))
    sh.fill.background()
    sh.line.color.rgb = __import__("pptx").dml.color.RGBColor.from_string(color)
    sh.line.width = Pt(1.0)
    ln = sh.line._get_or_add_ln()
    ln.append(parse_xml('<a:prstDash %s val="dash"/>' % nsdecls("a")))
    sh.shadow.inherit = False
    return sh


def link(slide, x1, y1, x2, y2, color=C_LINE, pts=0.9):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    ln.line.color.rgb = __import__("pptx").dml.color.RGBColor.from_string(color)
    ln.line.width = Pt(pts)
    return ln


# ================================================================ 43  RAG
def page_rag(slide, page):
    chrome(
        slide, page,
        "资料一多，为什么先检索再回答（RAG）",
        "两个理由：省一个量级的 token；少一批凭空编出来的答案。做法是“先找，再答”",
        "来源：Chroma《Context Rot》；Lewis 等 RAG 原论文（2020）；各厂 embedding 与上下文计费文档",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页回答“为什么用”，并且用一张图说清“靠什么找”。\n\n"
            "两个理由：\n"
            "一，省 token，省一个量级。100 页中文资料约 5–8 万 token，整包塞进去问一次付一次；"
            "检索只把命中的 3–5 段（一两千 token）送进去，差 20–50 倍。\n"
            "二，少编。材料不在眼前，模型不会说“没有”，会按“最像”往下续；"
            "原文片段摆在眼前，它有东西可抄、能标出处，再加一句“只依据片段回答”，编造率明显下降。\n\n"
            "下半页的图是重点，讲三句：\n"
            "① 每段资料和每个问题都被压成一串数字（向量），可以画在平面上；\n"
            "② 意思越近，点越近——问“防撞设施”，也能落到写着“挡块”的那段上，"
            "这是关键词搜索做不到的（字面一个词都不重合）；\n"
            "③ 但“近”不等于“对”。近邻里常有看着像、其实不是同一个东西的段落，"
            "所以工程上还要再做一次精排（rerank），并且必须给页码让人回原文核。\n\n"
            "这就是“搜索”：不是查字典，是找意思最近的那几段。\n\n"
            "边界：检索没找到，它照样会编。所以“找不到就说找不到”要写进规则，也要抽查。"
        ))

    # ---- 上：两个理由（并排两卡）
    section_bar(slide, L, 1.86, W, "为什么要用：两个理由", h=0.34, size=13)

    cards = [
        ("① 省 token", "省一个量级", [
            "100 页整包 ≈ 5–8 万 token，问一次付一次",
            "只检索命中 3–5 段 ≈ 1–2 千 token",
            "差 20–50 倍；追问十次，差出一整个月的额度",
        ], C_ORANGE),
        ("② 少编", "把“凭印象”换成“照着原文”", [
            "材料不在眼前：只能按“最像”续写，数字、条文最容易编",
            "原文摆到眼前：有东西可抄，还能标出处",
            "再加一句“只依据片段回答，找不到就说没找到”",
        ], C_TEAL),
    ]
    x = L
    for tag, head, lines, col in cards:
        rect(slide, x, 2.26, 5.95, 1.30, C_WHITE)
        rect(slide, x, 2.26, 0.06, 1.30, col)
        put(slide, x + 0.18, 2.32, 1.20, 0.24, tag, size=11, bold=True, color=col,
            anchor=MSO_ANCHOR.TOP)
        put(slide, x + 1.40, 2.32, 4.40, 0.24, head, size=10.5, bold=True,
            color=C_TITLE, anchor=MSO_ANCHOR.TOP)
        yy = 2.60
        for ln in lines:
            put(slide, x + 0.18, yy, 5.62, 0.26, "· " + ln, size=9, color=C_BODY,
                anchor=MSO_ANCHOR.TOP)
            yy += 0.29
        x += 5.95 + 0.23

    put(slide, L, 3.64, W, 0.24,
        "怎么做：入库切分建索引（只做一次）→ 问题向量化 → 取最近 3–5 段 → 拼进提示词 → 回答并给出处",
        size=9.5, color=C_MUTED, anchor=MSO_ANCHOR.TOP)

    # ---- 下：相似度示意图
    section_bar(slide, L, 3.98, W, "靠什么找：不是查关键词，是比“意思有多近”", h=0.34, size=13)

    # 坐标系
    px0, py0, px1, py1 = 0.98, 4.50, 5.42, 5.94
    hline(slide, px0, py1, px1 - px0, C_LINE)
    vline(slide, px0, py0, py1 - py0, C_LINE)
    put(slide, 2.55, py1 + 0.05, 2.30, 0.20, "向量维度 1（二维示意）", size=7.5,
        color=C_MUTED, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
    put(slide, 0.62, py0 - 0.04, 0.30, 0.60, ["向", "量", "维", "度", "2"], size=7.5,
        color=C_MUTED, anchor=MSO_ANCHOR.TOP)

    # 检索半径
    dashed_circle(slide, 4.10, 5.02, 0.46, C_ORANGE)
    put(slide, 3.90, 4.52, 0.90, 0.20, "检索半径", size=7.5, color=C_ORANGE,
        anchor=MSO_ANCHOR.TOP)

    dot(slide, 4.24, 5.06, 0.15, C_ORANGE)
    put(slide, 4.38, 4.94, 0.80, 0.20, "问题", size=7.5, bold=True, color=C_ORANGE,
        anchor=MSO_ANCHOR.TOP)

    dot(slide, 3.86, 4.90, 0.11, C_TEAL)
    put(slide, 2.06, 4.78, 1.62, 0.20, "① 挡块（变更令）", size=7.5, color=C_TITLE,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)

    dot(slide, 4.58, 5.26, 0.11, C_TEAL)
    put(slide, 4.70, 5.30, 1.10, 0.20, "② 台背回填", size=7.5, color=C_TITLE,
        anchor=MSO_ANCHOR.TOP)

    dot(slide, 2.62, 5.44, 0.09, C_MUTED)
    put(slide, 2.76, 5.46, 1.40, 0.20, "③ 沥青配合比", size=7.5, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP)

    dot(slide, 1.34, 5.14, 0.09, C_MUTED)
    put(slide, 1.48, 5.16, 1.40, 0.20, "④ 桥面铺装", size=7.5, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP)

    # 右侧三条
    rx = 5.86
    rows = [
        ("字面搜不到", "“防撞设施”与“挡块”一个字都不重合，关键词搜索直接空手", C_ORANGE),
        ("语义能落到", "两段话意思近，向量就靠得近；取最近三五段，正好把“挡块”那段捞上来", C_TEAL),
        ("但“近”≠“对”", "近邻里常有看着像、其实不是一回事的段落：要再精排一次，并给页码回原文核", C_HEAD),
    ]
    y = 4.42
    for k, v, col in rows:
        rect(slide, rx, y, 0.05, 0.42, col)
        put(slide, rx + 0.14, y, 1.05, 0.20, k, size=9, bold=True, color=col,
            anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 1.24, y, 5.60, 0.42, v, size=8.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.48

    put(slide, rx, 5.90, 6.85, 0.34,
        "工程三步：切分按结构（表头、条款、页码一起存）· 检索取近邻 · 回答带出处，能回原文",
        size=8.5, color=C_MUTED, anchor=MSO_ANCHOR.TOP)

    callout(slide, "一句话", "RAG 买的是“少付 token”和“少编”，靠的是向量距离，不是关键词", kind="cool")


# ================================================================ 44  工作流
def page_workflow(slide, page):
    chrome(
        slide, page,
        "把一件系统的事，整理成一条有标准的工作流",
        "审图、做修编报告，骨架其实是同一个：拆步骤 → 每步定“谁做、按什么判、交出什么”",
        "来源：本稿图纸审查与修编报告两个案例的实际执行与复核记录；表格结构说明",
        notes=(
            "讲解与会后阅读\n\n"
            "前面讲了 RAG 只解决“找得到”。这一页回答：找到之后，怎么保证“读得准、摆得对、算得出”。"
            "办法不是换更强的模型，是把一件事拆成步骤，每步写清标准和产出。\n\n"
            "左边是审图的五步，注意第三列——那是这一步“算不算做完”的判据，不是操作说明：\n"
            "① 定范围：几页、哪些表，图纸是“看图”还是“只 OCR”，先问清；\n"
            "② 提取：OCR 与视觉识别，标准是页数、表名、行数对得上，缺页要报出来；\n"
            "③ 逐项核对：拿图纸和图例、数量表互对，标准是每条结论都有位置和原文；\n"
            "④ 复算：关键工程量自己算一遍，标准是能独立复算；\n"
            "⑤ 汇总结论：分“明确不一致 / 待确认 / 未检查”，标准是不许把待确认写成已确认。\n\n"
            "右边做修编报告，同样的骨架：读清单、定位变更、套模板、复算、写原因。"
            "两个案例里踩的坑——单位 mm/cm 被当成错误、112 元差额、透水管核减 71.8663 万元——"
            "全都能归到某一步的“标准没定死”上。\n\n"
            "最下面那条最关键：工作流不是一次画完的。每跑一次，把踩到的坑写成一条标准，"
            "这就是“不断总结”，也是这套东西真正变成资产的地方。"
        ))

    # ---- 左：审图五步表
    section_bar(slide, L, 1.88, 7.45, "例一：审一张图纸（五步）", h=0.36, size=13)
    heads = [(0.72, 1.52, "步骤"), (2.28, 1.05, "谁做"),
             (3.38, 3.05, "按什么判（标准）"), (6.48, 1.42, "交出什么")]
    for x, w, t in heads:
        put(slide, x, 2.28, w, 0.24, t, size=9.5, bold=True, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP)
    hline(slide, L, 2.54, 7.45, C_LINE)

    rows = [
        ("① 定范围", "人", "页数、图号、哪些表；先确认是“看图”还是“只 OCR”", "一页清单", C_TEAL),
        ("② 提取", "OCR＋视觉", "页数、表名、行数对得上；缺页、糊字要点名", "结构化底表", C_TEAL),
        ("③ 逐项核对", "模型", "每条结论要有图中位置＋原文；不许写“疑似”", "问题清单", C_ORANGE),
        ("④ 复算", "脚本", "关键量能独立复算；算不出来不进正文", "计算底稿", C_TEAL),
        ("⑤ 汇总结论", "人", "分“不一致 / 待确认 / 未检查”，不许混写", "审查意见", C_ORANGE),
    ]
    y = 2.60
    for i, (a, b, c, d, col) in enumerate(rows):
        rowband(slide, L, y, 7.45, 0.66, i)
        rect(slide, L, y, 0.05, 0.66, col)
        put(slide, 0.72, y, 1.52, 0.66, a, size=10.5, bold=True, color=C_TITLE,
            spacing=1.05)
        put(slide, 2.28, y, 1.05, 0.66, b, size=9.5, color=C_MUTED)
        put(slide, 3.38, y, 3.05, 0.66, c, size=9.5, color=C_BODY, spacing=1.15)
        put(slide, 6.48, y, 1.42, 0.66, d, size=9.5, color=C_TEAL, spacing=1.1)
        y += 0.66

    # ---- 右：做报告
    section_bar(slide, 8.28, 1.88, 4.45, "例二：做修编报告", color=C_HEAD, h=0.36, size=13)
    rep = [
        ("读清单", "行数、口径、单位对得上"),
        ("定位变更", "每条变更能指到函件序号"),
        ("套模板", "章节、单元格位置不变"),
        ("复算金额", "112 元差额能说清来源"),
        ("写原因", "原因必须有据，不许编"),
    ]
    y = 2.32
    for k, v in rep:
        put(slide, 8.42, y, 1.20, 0.24, k, size=10, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 9.65, y, 3.05, 0.44, v, size=9.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.50

    rect(slide, 8.28, 4.98, 4.45, 1.02, C_WARM)
    put(slide, 8.44, 5.04, 4.15, 0.22, "两个案例的坑，都出在“标准没定死”", size=9.5,
        bold=True, color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
    put(slide, 8.44, 5.28, 4.15, 0.68,
        "把 mm/cm 当错误（没先定“看图例”）· 汇总与分项差 112 元（没定口径）· 核减数算不出（没复算）",
        size=9, color=C_BODY, spacing=1.2, anchor=MSO_ANCHOR.TOP)

    # ---- 底：不断总结
    rect(slide, L, 6.02, W, 0.26, C_ZEBRA)
    put(slide, 0.78, 6.02, W - 0.30, 0.26, [
        ("工作流是长出来的：", 10, True, C_TEAL),
        ("每跑一次，把踩到的坑写成一条标准塞回对应步骤——这一步比画流程图重要得多", 10, False, C_BODY)],
        spacing=1.0)

    callout(slide, "要点", "不是“换更强的模型”，是“把每一步的标准写下来”")


# ================================================================ 45  颗粒度
def page_grain(slide, page):
    chrome(
        slide, page,
        "工作流的颗粒度：写死到哪一步，留白留多少",
        "这页也解释一件事：为什么“整理成工作流”和“交给 Agent”不是对立，而是同一条线上的两档",
        "来源：本稿两案例的执行记录；各模型客户端说明；模型能力变化的观察（2026-10）",
        notes=(
            "讲解与会后阅读\n\n"
            "上一页把工作流拆出来了，这一页回答两个问题：写死到什么程度？会不会限制模型？\n\n"
            "先说判断标准，三条：格式固定、有唯一正确答案、出错代价高——三条都满足就写死。"
            "反过来，材料不全要临场判断、输出是自然语言、步骤事先说不清——就该留给模型。\n\n"
            "再说下面那个提醒，这条要讲重：**过细的流程会变成天花板**。"
            "三种表现：一是模型本来一次能做对的事，被拆成五步反而每步都掉一点；"
            "二是标准写成了“怎么做”（按这个顺序、用这个模板），模型发现更好的路径也不许走；"
            "三是新模型上线后，那套为旧模型打的补丁全成了负担。\n\n"
            "怎么留活口，三条：标准写“要达到什么”而不是“怎么做”；每一步内部留判断位；"
            "定期删规则——半年清一次，删掉那些“当年为了绕开某个毛病”的补丁。\n\n"
            "一句话收尾：工作流是为了让人能验收，不是为了让模型听话。"
        ))

    # 左：写死
    section_bar(slide, L, 1.88, 6.05, "该写死的：三条全中", h=0.36, size=13)
    hard = [
        ("格式固定", "Excel 行列、模板单元格、报告章节都是死的"),
        ("有唯一正确答案", "合计数、差额、核减额，算错就是错"),
        ("出错代价高", "要能复算、能追责，不能“大概对”"),
    ]
    y = 2.32
    for k, v in hard:
        dot(slide, 0.70, y + 0.11, 0.10, C_TEAL)
        put(slide, 0.86, y - 0.01, 1.55, 0.22, k, size=10.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.44, y - 0.01, 4.10, 0.44, v, size=9.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.46

    # 左：留白
    section_bar(slide, L, 3.72, 6.05, "该留白的：中一条就交给模型", color=C_ORANGE,
                h=0.36, size=13)
    soft = [
        ("材料不全", "缺页、口径不清，要临场判断"),
        ("输出是自然语言", "写原因、写信、写说明"),
        ("步骤说不清", "要边做边看结果再决定下一步"),
    ]
    y = 4.16
    for k, v in soft:
        dot(slide, 0.70, y + 0.11, 0.10, C_ORANGE)
        put(slide, 0.86, y - 0.01, 1.55, 0.22, k, size=10.5, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.44, y - 0.01, 4.10, 0.44, v, size=9.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.46

    put(slide, 0.74, 5.44, 5.95, 0.52,
        "两档不是对立的：同一件事，外层（交付、审批、复算）写死，内层（怎么找、怎么写）留给模型。",
        size=9.5, color=C_MUTED, spacing=1.2, anchor=MSO_ANCHOR.TOP)

    # 右：提醒
    rect(slide, 6.95, 1.88, 5.78, 0.44, C_ORANGE)
    put(slide, 7.14, 1.88, 5.45, 0.44, "重要提醒：过细的流程，会变成天花板", size=13,
        bold=True, color=C_WHITE)

    warns = [
        ("拆分反而掉分", "模型本来一次能做对的，被拆成五步，每步掉一点，最后不如一次做完"),
        ("标准写成了操作", "“按这个顺序、用这个模板”——它发现更快的路也不许走"),
        ("给旧模型打的补丁", "新模型一升级，那些绕坑的规则全成了负担，还拖慢速度"),
    ]
    y = 2.44
    for k, v in warns:
        put(slide, 7.14, y, 1.75, 0.24, k, size=10, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 8.94, y, 3.68, 0.56, v, size=9.5, color=C_BODY, spacing=1.15,
            anchor=MSO_ANCHOR.TOP)
        y += 0.64

    rect(slide, 6.95, 4.44, 5.78, 1.30, C_COOL)
    put(slide, 7.14, 4.50, 5.45, 0.22, "怎么留活口：三条", size=10, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    tips = [
        "标准写“要达到什么”，不写“怎么做”",
        "每一步内部留一个判断位，允许它换做法",
        "定期删规则：半年清一次，删当年绕坑的补丁",
    ]
    yy = 4.76
    for t in tips:
        put(slide, 7.14, yy, 5.40, 0.26, "· " + t, size=9.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP)
        yy += 0.30

    callout(slide, "一句话", "工作流是为了让人能验收，不是为了让模型听话", kind="warn")


# ================================================================ 48  企业路线
def page_enterprise(slide, page):
    chrome(
        slide, page,
        "别人到底怎么做的：四条路线，和它们各自的账",
        "同一件事有四种做法，技术含量和回报完全不同；现在的主流是“通用模型 ＋ 工作流 ＋ 检索”",
        "来源：公开的企业落地案例与技术路线文章（含 90% 场景走 RAG 的经验之谈）；各厂微调与上下文计费文档；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页不再讲架构图，讲别人实际怎么干的，以及各自的账。四条路线，按“现在主流程度”排。\n\n"
            "一，只做 Agent（挂几个工具）。看着技术含量最低——确实，多数就是把现成模型接上文件读写、"
            "搜索、代码执行。但它真正的价值不在技术：在于把一件事的步骤、标准、边界写清楚。"
            "这部分是资产，换模型也不会丢。缺点也明显：不稳定，同样的输入两次结果不一样。\n\n"
            "二，工作流 ＋ 通用模型（现在最主流）。把固定步骤写成代码或编排，只在需要判断的地方调模型。"
            "上一页和前一页讲的就是这个。占比最高，因为最划算：稳定、可复算、出问题能定位。\n\n"
            "三，RAG 知识库。九成“企业知识问答”场景走这条。见效快、成本低，缺点是效果上限受检索质量限制，"
            "检索差一点，回答就飘。\n\n"
            "四，自己后训练 / 微调一个专用模型。这条最容易被高估。真要做的条件很苛刻："
            "数据不能出内网、任务极窄且量大、有稳定的标注与评测能力。代价是三个月起步、要人、要卡；"
            "更要命的是通用模型半年一升级，专用小模型很容易被追平甚至反超。"
            "所以现实里更常见的是“通用模型 ＋ 少量示例 / 提示词”先顶上，微调留到最后。\n\n"
            "右下那句话是全页落点：四条路线不是选择题，多数单位最后是二＋三，一和四按需加。"
        ))

    routes = [
        ("① 只做 Agent", "把模型接上文件、搜索、代码执行，让它自己跑", [
            ("技术含量", "低：多数是接线，不是训练"),
            ("真实价值", "把步骤、标准、边界写清楚——这部分换模型也不丢"),
            ("硬伤", "不稳定：同样输入两次结果不同"),
            ("什么时候选", "步骤说不清、要边做边看；接受结果波动"),
        ], C_ORANGE, "看得最像 AI"),
        ("② 工作流＋通用模型", "固定步骤写死，只在判断处调模型", [
            ("技术含量", "中：难在拆解与标准，不在代码"),
            ("真实价值", "稳定、可复算、出问题能定位"),
            ("硬伤", "流程画太细会限制模型（见上一页）"),
            ("什么时候选", "步骤稳定、要能复算、出问题要能定位"),
        ], C_TEAL, "现在最主流"),
        ("③ RAG 知识库", "资料切分建索引，问什么取什么", [
            ("技术含量", "中低：框架现成，难在切分与评测"),
            ("真实价值", "见效快、成本低；资料改了立刻生效"),
            ("硬伤", "上限被检索质量卡住，检索差回答就飘"),
            ("什么时候选", "资料多、更新快、答案要能指到原文"),
        ], C_HEAD, "约九成知识问答"),
        ("④ 自训 / 微调专用模型", "拿自己的数据再训一个", [
            ("技术含量", "高：数据、标注、评测、算力都要"),
            ("真实价值", "数据不出内网；极窄任务上又便宜又快"),
            ("硬伤", "三个月起步；通用模型一升级就被追平"),
            ("什么时候选", "数据不能出内网、任务极窄、有标注与评测"),
        ], C_MUTED, "最容易被高估"),
    ]

    x = L
    cw = 2.92
    for name, sub, lines, edge, tag in routes:
        rect(slide, x, 1.88, cw, 3.94, C_WHITE)
        rect(slide, x, 1.88, cw, 0.06, edge)
        put(slide, x + 0.18, 2.00, cw - 0.36, 0.40, name, size=13, bold=True,
            color=C_TITLE, spacing=1.05)
        rect(slide, x + 0.18, 2.42, cw - 0.36, 0.26, C_ZEBRA)
        put(slide, x + 0.24, 2.42, cw - 0.46, 0.26, tag, size=8.5, color=C_BODY)
        put(slide, x + 0.18, 2.74, cw - 0.36, 0.42, sub, size=8.5, color=C_MUTED,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        yy = 3.22
        for k, v in lines:
            put(slide, x + 0.18, yy, cw - 0.36, 0.18, k, size=8, bold=True,
                color=C_ORANGE if k in ("硬伤", "什么时候选") else C_TEAL, anchor=MSO_ANCHOR.TOP)
            put(slide, x + 0.18, yy + 0.19, cw - 0.36, 0.50, v, size=8.5,
                color=C_BODY, spacing=1.15, anchor=MSO_ANCHOR.TOP)
            yy += 0.74
        x += cw + 0.15

    # 底部：怎么选
    rect(slide, L, 5.94, W, 0.34, C_ZEBRA)
    put(slide, 0.78, 5.94, W - 0.30, 0.34, [
        ("怎么选：", 10, True, C_TEAL),
        ("先确认“数据能不能出内网”→ 再确认“有没有稳定评测”→ 两条都过了才谈微调；否则先 ②＋③",
         10, False, C_BODY)], spacing=1.0)

    callout(slide, "落点", "四条不是选择题：多数单位最后是“工作流＋检索”，Agent 与微调按需加", kind="cool")


# ================================================================ main
def main(src, dst):
    prs = Presentation(src)
    jobs = [(43, page_rag), (44, page_workflow), (45, page_grain), (48, page_enterprise)]
    for idx, fn in jobs:
        footer = None
        for sh in prs.slides[idx].shapes:
            if sh.has_text_frame and sh.left is not None:
                if abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.25:
                    footer = sh.text_frame.text.strip()
        clear(prs.slides[idx])
        fn(prs.slides[idx], footer or (idx + 1))
        print("rebuilt", idx + 1, "footer", footer)

    prs.save(dst)
    print("saved", dst, len(prs.slides))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
