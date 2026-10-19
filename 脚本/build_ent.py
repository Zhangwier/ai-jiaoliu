# -*- coding: utf-8 -*-
"""企业落地部分优化：加「平台必备件」与「工作流/RAG 怎么养」两页，并补 Z 页来源。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

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


# ================================================================ 平台必备件
def page_platform(slide, page):
    chrome(
        slide, page,
        "两个真实开源项目：企业版到底多出了哪些件",
        "MOVO 把 DSH 包成能上生产的平台；AgenTeam 把模型、技能、知识、工具拼成“数字员工”",
        "来源：github.com/himovo/movo（README 与架构图）；github.com/Stonewuu/agenteam（README）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页不讲路线，讲“长什么样”。两个都是能直接部署的开源项目，正好一个偏“平台底座”，"
            "一个偏“协作与岗位”。\n\n"
            "MOVO：一句话是“DSH 负责跑 Agent，MOVO 负责把它搬进企业”。默认 Docker Compose 起 12 个服务："
            "网关、两个 Web（用户台 / 管理台）、三个应用 API（对话、管理、文档）、DSH Runtime Host、"
            "文档 worker、MongoDB、Redis、Weaviate、一次性密钥引导。能力上覆盖企业知识问答、多源研究、"
            "文档解析（PDF/DOCX/XLSX/PPTX/CSV/MD，含图表）、内容生成（报告、PPTX、表格、PDF）、"
            "Skill/Tool/MCP、定时任务、敏感动作审批、执行追踪与产物留存，以及组织/用户/角色/模型/知识/审计的管理台。\n\n"
            "AgenTeam（群策）：把“模型＋技能＋知识库＋工具”组合成一个数字员工，发到能力中心上架、按资源授权；"
            "工作台里按项目组织，历史对话、项目文件、待办都在一处；还有定时任务、企业微信/飞书绑定与业务通知。"
            "技术栈是 Java 21 + Spring Boot + AgentScope 2 + Next.js，MySQL 存业务、Redis 存会话，"
            "另有一个带浏览器与 Git 的沙箱镜像。\n\n"
            "把它们叠在一起看，企业要多出来的就是下面这六件——这才是“落地”真正的工作量：\n"
            "身份与权限、知识与检索、能力（Skill/Tool/MCP）管理、审批与审计、交付物管理、定时与通知。\n\n"
            "收尾一句：这六件里没有一件是“写提示词”。所以企业项目卡住，通常不是模型不行，是这六件没补齐。"
        ))

    # ---- 左：MOVO
    section_bar(slide, L, 1.88, 6.05, "MOVO：DSH 的企业外壳", h=0.36, size=13)
    put(slide, L, 2.30, 6.05, 0.24,
        "12 个容器：网关 · 用户台 · 管理台 · 对话 API · 管理 API · 文档服务 · DSH Runtime · "
        "Mongo / Redis / Weaviate",
        size=9, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    movo = [
        ("知识", "内部文档检索、多轮研究、结论带引用"),
        ("文档", "PDF / DOCX / XLSX / PPTX / CSV / MD，含图表"),
        ("生成", "报告、PPTX、表格、PDF、Markdown 交付"),
        ("能力", "Skill / Tool / MCP，接业务系统"),
        ("治理", "敏感动作审批、执行追踪、产物留存、审计"),
        ("管理", "组织 / 用户 / 角色 / 模型 / 知识 / 健康度"),
    ]
    y = 2.62
    for i, (k, v) in enumerate(movo):
        rowband(slide, L, y, 6.05, 0.50, i)
        put(slide, 0.74, y, 0.95, 0.50, k, size=10.5, bold=True, color=C_TEAL)
        put(slide, 1.74, y, 4.20, 0.50, v, size=9.5, color=C_BODY, spacing=1.1)
        y += 0.50

    put(slide, L, 5.66, 6.05, 0.42,
        "“DSH 负责跑 Agent，MOVO 负责把它搬进企业”——这句是它 README 的原话。", size=9,
        color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # ---- 右：AgenTeam
    section_bar(slide, 6.95, 1.88, 5.78, "AgenTeam：数字员工与协作", color=C_ORANGE,
                h=0.36, size=13)
    put(slide, 6.95, 2.30, 5.78, 0.24,
        "数字员工 ＝ 模型 ＋ 技能 ＋ 知识库 ＋ 工具；上架后按资源授权给团队",
        size=9, color=C_MUTED, anchor=MSO_ANCHOR.TOP)

    agt = [
        ("岗位化", "一个“数字员工”就是一个配置好的智能体，可上架、可版本"),
        ("工作面", "按项目组织：会话、文件、待办都在同一处，能接着做"),
        ("节奏", "一次性 / 每日 / 每周 / 每月定时任务，可暂停、可查历史"),
        ("接入", "企业微信、飞书：绑定账号、发业务通知、逐人回执"),
        ("基建", "Java 21 + AgentScope 2；MySQL 存业务、Redis 存会话"),
        ("沙箱", "带浏览器与 Git 的执行沙箱，跑文件与代码不外溢"),
    ]
    y = 2.62
    for i, (k, v) in enumerate(agt):
        rowband(slide, 6.95, y, 5.78, 0.50, i)
        put(slide, 7.12, y, 0.95, 0.50, k, size=10.5, bold=True, color=C_ORANGE)
        put(slide, 8.12, y, 4.42, 0.50, v, size=9.5, color=C_BODY, spacing=1.1)
        y += 0.50

    put(slide, 6.95, 5.66, 5.78, 0.42,
        "它把“待办、通知、版本、授权”这些组织动作也做进产品里——这部分最容易被自研忽略。",
        size=9, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    callout(slide, "共同点", "企业要多出来的六件，没有一件是“写提示词”——卡住的通常是这六件", kind="cool")


# ================================================================ 工作流与 RAG 怎么养
def page_ops(slide, page):
    chrome(
        slide, page,
        "真正难的不是搭起来，是让它越用越准",
        "工作流和 RAG 都得“养”：把日常使用中的返工和失败变成规则，否则上线三个月就僵住",
        "来源：本稿两案例的返工记录；MOVO / AgenTeam 的产品机制；检索评测常用做法；本稿整理（2026-10-18）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页是全节的重点，也是我自己的判断：企业 AI 项目失败，很少死在“搭不起来”，"
            "多数死在“没人养”。\n\n"
            "先讲工作流这半边。四个动作，构成一个环：\n"
            "① 记录：每一次人工返工都记一条——哪一步、改了什么、为什么。这一条最容易做，也最容易漏；\n"
            "② 提炼：同一类返工出现三次，就升级成一条标准，或者一个小校验脚本（比如“合计必须与分项一致”）；\n"
            "③ 清理：每季度过一遍旧规则，删掉那些为老模型打的补丁——不删就会变成天花板（上一页讲过）；\n"
            "④ 版本化：工作流要有版本号和变更记录，谁改的、为什么改、影响哪些场景。没版本就没法回滚。\n\n"
            "再讲 RAG 这半边。关键是“把失败分类”，不要笼统说“答得不好”：\n"
            "A 没检索到（召回问题）→ 动切分粒度、补别名同义词、加元数据过滤；\n"
            "B 检索到了但没用（排序问题）→ 加 rerank、调 top-k、把片段压短；\n"
            "C 用了但答错（生成问题）→ 改提示词约束、要求引用原文、加二次核对。\n"
            "三类混在一起，就永远不知道该改哪。\n\n"
            "再往下是评测：把真实问题沉淀成 50–100 道固定题，每次改切分、改模型、改提示词都重跑一遍。"
            "没有这套题，所有优化都是“感觉好一点”。\n\n"
            "最后是组织：得有人负责这件事——不是 IT 一个部门，而是业务里那个“最会用”的人，"
            "每周花两小时看返工记录和检索日志。这一条比任何工具都重要。\n\n"
            "收尾那句：把使用变成数据，把数据变成规则。听着朴素，但九成单位没做。"
        ))

    # ---- 左：工作流
    section_bar(slide, L, 1.88, 6.05, "工作流怎么长：四个动作，一个环", h=0.36, size=13)
    loop = [
        ("① 记录", "每次人工返工记一条：哪一步、改了什么、为什么", C_TEAL),
        ("② 提炼", "同类返工出现三次 → 升级成一条标准或一个校验脚本", C_TEAL),
        ("③ 清理", "每季度删旧规则，尤其“当年绕坑”的补丁", C_ORANGE),
        ("④ 版本化", "工作流带版本号与变更记录：谁改、为何改、影响谁", C_ORANGE),
    ]
    vline(slide, 0.72, 2.36, 2.86, C_LINE, 1.0)
    y = 2.32
    for k, v, col in loop:
        dot(slide, 0.72, y + 0.13, 0.12, col)
        put(slide, 0.98, y, 5.45, 0.24, k, size=11.5, bold=True, color=col,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 0.98, y + 0.25, 5.45, 0.46, v, size=9.5, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.72

    put(slide, 0.98, 5.26, 5.60, 0.54,
        "判断标准：一条规则如果半年没人碰过、也没引发过返工，就该删——它只是在拖慢模型。",
        size=9, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # ---- 右：RAG
    section_bar(slide, 6.95, 1.88, 5.78, "RAG 怎么调：先把失败分成三类", color=C_ORANGE,
                h=0.36, size=13)
    fails = [
        ("A 没检索到", "召回问题", "切分粒度、补别名同义词、加元数据过滤", C_ORANGE),
        ("B 检索到没用", "排序问题", "加 rerank、调 top-k、把片段压短", C_TEAL),
        ("C 用了但答错", "生成问题", "改约束、要求引原文、加二次核对", C_HEAD),
    ]
    y = 2.32
    for k, tag, fix, col in fails:
        rect(slide, 6.95, y, 5.78, 0.78, C_WHITE)
        rect(slide, 6.95, y, 0.05, 0.78, col)
        put(slide, 7.10, y + 0.07, 2.05, 0.22, k, size=10.5, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 9.20, y + 0.07, 0.95, 0.22, tag, size=8, color=C_MUTED,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
        put(slide, 7.10, y + 0.31, 5.45, 0.42, "→ " + fix, size=9, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.86

    rect(slide, 6.95, 4.96, 5.78, 0.86, C_COOL)
    put(slide, 7.12, 5.00, 5.48, 0.22, "评测集：把优化从“感觉”变成“数”", size=9.5,
        bold=True, color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    put(slide, 7.12, 5.24, 5.48, 0.52,
        "沉淀 50–100 道真实问题；每次改切分、换模型、动提示词都重跑一遍，分数掉了就回退",
        size=9, color=C_BODY, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # ---- 底部：谁来做
    rect(slide, L, 5.92, W, 0.34, C_WARM)
    put(slide, 0.78, 5.92, W - 0.34, 0.34, [
        ("谁来做：", 9.5, True, C_ORANGE),
        ("得有一个“养”的人——不是 IT 一个部门，而是业务里最会用 AI 的那位，每周两小时看返工记录与检索日志",
         9.5, False, C_BODY)], spacing=1.0)

    callout(slide, "一句话", "把“使用”变成数据，把“数据”变成规则——九成单位没做这一步", kind="cool")


# ================================================================ Z3 增补链接
def rebuild_z3(slide, page):
    chrome(
        slide, page,
        "Z3  链接清单（二）：工具、平台与社区",
        "看仓库判断“它要什么权限、缺哪一层”，看官网判断“厂商怎么说”，看社区判断“大家在踩什么坑”",
        "说明：链接为 2026-10-18 核对；第三方站点内容可能变动，引用前先复核",
        notes=(
            "讲解与会后阅读\n\n"
            "工具与仓库分两组：一组是“壳”（magpie、CLIProxyAPI），一组是“企业平台”（MOVO、AgenTeam）。"
            "看企业平台时重点看两处：它的权限与审计做到什么程度，它的知识库与检索怎么迭代。\n\n"
            "社区与客户端：L 站两个帖子是 D3 那页三个例子的来源；另外补了软件官网、官方多开说明和看雪帖。\n\n"
            "边界：社区帖只证明“有人在这么做”，不提供做法。"
        ))

    def group(gx, gy, gw, title, items, pitch=0.62):
        put(slide, gx, gy, gw, 0.28, title, size=13.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        hline(slide, gx, gy + 0.30, gw, C_TEAL, 1.0)
        yy = gy + 0.40
        for label, url in items:
            put(slide, gx, yy, gw, 0.22, label, size=11, bold=True, color=C_TITLE,
                anchor=MSO_ANCHOR.TOP)
            box = put(slide, gx, yy + 0.22, gw, 0.20, url, size=8, color=C_MUTED,
                      anchor=MSO_ANCHOR.TOP)
            r = box.text_frame.paragraphs[0].runs[0]
            rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
            rPr = r._r.get_or_add_rPr()
            rPr.set("u", "none")
            rPr.append(parse_xml(
                '<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))
            rPr.append(parse_xml(
                '<a:solidFill %s><a:srgbClr val="6B7385"/></a:solidFill>' % nsdecls("a")))
            for tag in ("a:latin", "a:ea", "a:cs"):
                rPr.append(parse_xml(
                    '<%s %s typeface="微软雅黑"/>' % (tag, nsdecls("a"))))
            yy += pitch
        return yy

    shells = [
        ("magpie（本地网关）", "https://github.com/yetone/magpie"),
        ("CLIProxyAPI / CPA", "https://github.com/router-for-me/CLIProxyAPI"),
        ("dsh-purge（破限插件）", "https://github.com/YuJunZhiXue/dsh-purge"),
    ]
    platforms = [
        ("MOVO（DSH 企业版）", "https://github.com/himovo/movo"),
        ("MOVO 中文说明", "https://github.com/himovo/movo/blob/main/README.zh-CN.md"),
        ("AgenTeam 群策（多智能体）", "https://github.com/Stonewuu/agenteam"),
    ]
    community = [
        ("L 站：*果 去广告会员版帖", "https://linux.do/t/topic/2779542"),
        ("L 站：**模拟器 绿色多开帖", "https://linux.do/t/topic/1516872"),
        ("**造价软件 官网", "http://smartcost.com.cn/"),
        ("MuMu 官方多开说明", "https://mumu.163.com/help/20210913/35047_971679.html"),
        ("看雪：加密锁破解经验帖", "https://bbs.kanxue.com/thread-255826.htm"),
        ("GitHub：*果 广告形态适配", "https://github.com/clarithromycine/ad-auto"),
    ]

    put(slide, L, 1.88, 5.05, 0.28, "壳：网关与反代", size=13.5, bold=True, color=C_TEAL,
        anchor=MSO_ANCHOR.TOP)
    hline(slide, L, 2.18, 5.05, C_TEAL, 1.0)
    y = 2.28
    for label, url in shells:
        put(slide, L, y, 5.05, 0.22, label, size=11, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        box = put(slide, L, y + 0.22, 5.05, 0.20, url, size=8, color=C_MUTED,
                  anchor=MSO_ANCHOR.TOP)
        r = box.text_frame.paragraphs[0].runs[0]
        rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
        rPr = r._r.get_or_add_rPr()
        rPr.set("u", "none")
        rPr.append(parse_xml('<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))
        rPr.append(parse_xml(
            '<a:solidFill %s><a:srgbClr val="6B7385"/></a:solidFill>' % nsdecls("a")))
        y += 0.62

    section_bar(slide, L, 4.20, 5.05, "企业平台：可以直接部署的两个", color=C_ORANGE,
                h=0.34, size=12)
    y = 4.64
    for label, url in platforms:
        put(slide, L, y, 5.05, 0.22, label, size=10.5, bold=True, color=C_TITLE,
            anchor=MSO_ANCHOR.TOP)
        box = put(slide, L, y + 0.21, 5.05, 0.20, url, size=8, color=C_MUTED,
                  anchor=MSO_ANCHOR.TOP)
        r = box.text_frame.paragraphs[0].runs[0]
        rId = slide.part.relate_to(url, RT.HYPERLINK, is_external=True)
        rPr = r._r.get_or_add_rPr()
        rPr.set("u", "none")
        rPr.append(parse_xml('<a:hlinkClick %s r:id="%s"/>' % (nsdecls("a", "r"), rId)))
        rPr.append(parse_xml(
            '<a:solidFill %s><a:srgbClr val="6B7385"/></a:solidFill>' % nsdecls("a")))
        y += 0.58

    group(6.72, 1.88, 5.80, "社区与客户端", community)

    callout(slide, "边界", "引社区帖只证明“有人在这么做”，不提供做法本身", kind="warn")


# ================================================================ main
def shift_footers(prs, after_idx, delta):
    for i, s in enumerate(prs.slides):
        if i <= after_idx:
            continue
        for sh in s.shapes:
            if not sh.has_text_frame or sh.left is None:
                continue
            if abs(sh.left / 914400 - 11.73) > 0.05 or sh.width / 914400 > 1.25:
                continue
            p = sh.text_frame.paragraphs[0]
            if p.runs and p.runs[0].text.strip().isdigit():
                p.runs[0].text = str(int(p.runs[0].text.strip()) + delta)


def main(src, dst):
    prs = Presentation(src)
    blank = next(l for l in prs.slide_layouts if l.name == "Blank")

    s1 = prs.slides.add_slide(blank)
    page_platform(s1, 51)
    i1 = len(prs.slides) - 1

    s2 = prs.slides.add_slide(blank)
    page_ops(s2, 52)
    i2 = len(prs.slides) - 1

    lst = prs.slides._sldIdLst
    items = list(lst)
    for el in items:
        lst.remove(el)
    order = items[:49] + [items[i1], items[i2]] + items[49:i1] + items[i1 + 1:i2]
    for el in order:
        lst.append(el)

    # 新页之后页码 +2（原 49.. 起）
    shift_footers(prs, 48, 2)

    # Z3 重做（现在在 index 88）
    z3 = None
    for i, s in enumerate(prs.slides):
        for sh in s.shapes:
            if sh.has_text_frame and sh.name == "TextBox 4" and \
                    sh.text_frame.text.startswith("Z3"):
                z3 = i
    if z3 is None:
        raise SystemExit("Z3 not found")
    footer = None
    for sh in prs.slides[z3].shapes:
        if sh.has_text_frame and sh.left is not None:
            if abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.25:
                footer = sh.text_frame.text.strip()
    clear(prs.slides[z3])
    rebuild_z3(prs.slides[z3], footer or (z3 + 1))
    print("rebuilt Z3 at", z3 + 1, "footer", footer)

    # D 扉页页数 + Z 备注补两条来源
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "共 15 页" in r.text:
                        r.text = r.text.replace("共 15 页", "共 17 页")

    for s in prs.slides:
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.startswith("Z  参考资料与案例来源"):
                notes = s.notes_slide.notes_text_frame
                if "S57" not in notes.text:
                    notes.text = notes.text + (
                        "\n\n成品13 新增来源（2026-10-18 查询）\n"
                        "S57 MOVO（自托管企业 Agent 平台，基于 DSH）："
                        "https://github.com/himovo/movo ；中文说明 "
                        "https://github.com/himovo/movo/blob/main/README.zh-CN.md\n"
                        "S58 AgenTeam 群策（多智能体协作平台）："
                        "https://github.com/Stonewuu/agenteam\n"
                        "口径：两个项目的服务数、技术栈与能力边界均引自各自的 README；"
                        "“六件必备件”与“工作流 / RAG 怎么养”是本稿的归纳，不是项目原文。")

    prs.save(dst)
    print("saved", dst, len(prs.slides))
    for i, s in enumerate(prs.slides, 1):
        if 47 <= i <= 54 or 84 <= i <= 90:
            t = ""
            for sh in s.shapes:
                if sh.has_text_frame and sh.name == "TextBox 4":
                    t = sh.text_frame.text.split("\n")[0]
            print(f"  {i:3d}  {t[:56]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
