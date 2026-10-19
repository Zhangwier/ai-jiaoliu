# -*- coding: utf-8 -*-
"""New Appendix E for the AI 交流分享 deck: one company per page, with AI娘 art.

Base deck: work_e/deck16.pptx (verbatim copy of 成品16).
Output:   a NEW file, never touching 成品16/17/18.
"""
import copy
import os
import shutil

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

FONT = "微软雅黑"
PANOSE = "020B0503020204020204"
C_TITLE = "14213D"
C_BODY = "2B3445"
C_MUTED = "6B7385"
C_TEAL = "0F766E"
C_ORANGE = "D9622B"
C_HEAD = "12203A"
C_LINE = "DAD7CE"
C_WHITE = "FFFFFF"
C_WARM = "FBE9DE"
C_COOL = "E1F0EE"
C_ZEBRA = "ECEAE4"
C_BG = "F6F5F1"
C_HAIR = "E2DFD6"

L = 0.60
W = 12.13
SRC = r"work_e\deck16.pptx"
OUT = "AI交流分享-成品16+附录E-一页一家公司.pptx"
IMGDIR = r"work_e\img"
ART = "AI娘"


# ----------------------------------------------------------------- primitives
def set_bg(slide, color=C_BG):
    cSld = slide._element.find(qn("p:cSld"))
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    cSld.insert(0, parse_xml(
        '<p:bg %s><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        "<a:effectLst/></p:bgPr></p:bg>" % (nsdecls("p", "a"), color)))


def style_run(run, size, bold, color):
    rPr = run._r.get_or_add_rPr()
    rPr.set("sz", str(int(round(size * 100))))
    rPr.set("b", "1" if bold else "0")
    rPr.set("i", "0")
    for tag in ("a:solidFill", "a:latin", "a:ea", "a:cs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
    rPr.append(parse_xml(
        '<a:solidFill %s><a:srgbClr val="%s"/></a:solidFill>' % (nsdecls("a"), color)))
    for tag in ("a:latin", "a:ea", "a:cs"):
        rPr.append(parse_xml(
            '<%s %s typeface="%s" panose="%s" charset="-122"/>'
            % (tag, nsdecls("a"), FONT, PANOSE)))


def textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.MIDDLE):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    bodyPr = tf._txBody.bodyPr
    for tag in ("a:spAutoFit", "a:normAutofit", "a:noAutofit"):
        for e in bodyPr.findall(qn(tag)):
            bodyPr.remove(e)
    bodyPr.append(parse_xml("<a:noAutofit %s/>" % nsdecls("a")))
    return box


def put(slide, x, y, w, h, content, size=15, color=C_BODY, bold=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, spacing=1.15, after=0):
    if isinstance(content, str):
        paras = [[(content, size, bold, color)]]
    elif content and isinstance(content[0], str):
        paras = [[(t, size, bold, color)] for t in content]
    elif content and isinstance(content[0], tuple):
        paras = [list(content)]
    else:
        paras = [list(p) for p in content]

    box = textbox(slide, x, y, w, h, anchor)
    tf = box.text_frame
    for i, runs in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(after)
        for t, sz, bd, cl in runs:
            r = p.add_run()
            r.text = t
            style_run(r, sz, bd, cl)
    return box


def rect(slide, x, y, w, h, fill=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
    sh = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    if fill:
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    else:
        sh.fill.background()
    sh.line.fill.background()
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def hline(slide, x, y, w, color=C_LINE, pts=0.75, dash=None):
    ln = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x), Inches(y), Inches(x + w), Inches(y))
    ln.line.color.rgb = RGBColor.from_string(color)
    ln.line.width = Pt(pts)
    if dash:
        ln.line._get_or_add_ln().append(parse_xml(
            '<a:prstDash %s val="%s"/>' % (nsdecls("a"), dash)))
    return ln


def chrome(slide, page, title, subtitle, source):
    set_bg(slide)
    put(slide, L, 0.32, 6.0, 0.30, "附录  会后速查", size=12, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    put(slide, 11.73, 0.32, 1.00, 0.30, page, size=12, color=C_MUTED,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
    rect(slide, L, 0.66, 0.50, 0.06, C_TEAL)
    put(slide, L, 0.78, W, 0.70, title, size=26, bold=True, color=C_TITLE,
        spacing=1.0)
    put(slide, L, 1.50, W, 0.42, subtitle, size=17, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP, spacing=1.0)
    hline(slide, L, 7.00, W)
    put(slide, L, 7.06, 10.50, 0.30, source, size=10.5, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP)
    put(slide, 10.73, 7.06, 2.00, 0.30, "来源索引 ›", size=10.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)


def callout(slide, label, body, kind="cool"):
    bar = C_WARM if kind == "warn" else C_COOL
    edge = C_ORANGE if kind == "warn" else C_TEAL
    rect(slide, L, 6.35, W, 0.50, bar)
    rect(slide, L, 6.35, 0.09, 0.50, edge)
    put(slide, 0.85, 6.35, 1.05, 0.50, label, size=14, bold=True, color=edge)
    put(slide, 1.78, 6.35, 10.70, 0.50, body, size=17, bold=True,
        color=C_TITLE, spacing=1.0)


# ------------------------------------------------------------------ image prep
def prep(key, src):
    os.makedirs(IMGDIR, exist_ok=True)
    dst = os.path.join(IMGDIR, key + ".jpg")
    im = Image.open(os.path.join(ART, src)).convert("RGB")
    w, h = im.size
    target_h = 1500
    im = im.resize((int(w * target_h / h), target_h), Image.LANCZOS)
    im.save(dst, "JPEG", quality=90, optimize=True)
    return dst


def art_card(slide, path):
    """White card + portrait, left column."""
    card = rect(slide, 0.60, 1.90, 2.76, 4.44, C_WHITE)
    card.line.color.rgb = RGBColor.from_string(C_HAIR)
    card.line.width = Pt(0.75)
    im = Image.open(path)
    w, h = im.size
    box_h = 4.26
    box_w = box_h * w / h
    slide.shapes.add_picture(path, Inches(0.60 + (2.76 - box_w) / 2),
                             Inches(1.99), Inches(box_w), Inches(box_h))


def logo_badge(slide, logo_path, size=0.62):
    """Brand mark in the header row, left of the page number."""
    x = 10.78
    y = 0.28
    chip = rect(slide, x - 0.06, y - 0.05, size + 0.12, size + 0.12, C_WHITE,
                MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
    chip.line.color.rgb = RGBColor.from_string(C_HAIR)
    chip.line.width = Pt(0.75)
    slide.shapes.add_picture(logo_path, Inches(x), Inches(y),
                             Inches(size), Inches(size))


def art_placeholder(slide, word, note):
    card = rect(slide, 0.60, 1.90, 2.76, 4.44, C_WHITE)
    card.line.color.rgb = RGBColor.from_string("B9C6C2")
    card.line.width = Pt(1.0)
    card.line._get_or_add_ln().append(parse_xml(
        '<a:prstDash %s val="dash"/>' % nsdecls("a")))
    put(slide, 0.72, 3.15, 2.52, 0.70, word, size=40, bold=True,
        color="C7D8D4", align=PP_ALIGN.CENTER, spacing=1.0)
    put(slide, 0.78, 3.95, 2.40, 0.80, note, size=10.5, color=C_MUTED,
        align=PP_ALIGN.CENTER, spacing=1.15)


# ------------------------------------------------------------------ page build
def company_page(prs, page, data):
    slide = prs.slides.add_slide(prs.slides[0].slide_layout)
    chrome(slide, str(page), data["title"], data["sub"], data["src"])
    if data.get("logo"):
        logo_badge(slide, data["logo"])
    if data.get("art"):
        art_card(slide, data["art"])
    else:
        art_placeholder(slide, data.get("word", "dots"), data["phnote"])

    rows = [("模型线", data["models"]), ("外壳 Harness", data["harness"]),
            ("强项", data["good"]), ("短板 / 边界", data["bad"]),
            ("在本分享", data["here"])]
    tops = [1.92, 2.82, 3.72, 4.62, 5.52]
    heights = [0.86, 0.86, 0.86, 0.86, 0.62]
    for i, (label, body) in enumerate(rows):
        y, h = tops[i], heights[i]
        if i:
            hline(slide, 3.62, y - 0.06, 9.11, C_HAIR, 0.6)
        put(slide, 3.62, y, 1.10, h, label, size=11.5, bold=True,
            color=C_TEAL, anchor=MSO_ANCHOR.TOP, spacing=1.0)
        put(slide, 4.78, y - 0.02, 7.95, h, body, size=11.5, color=C_BODY,
            anchor=MSO_ANCHOR.TOP, spacing=1.28)
    callout(slide, "选它", data["pick"])
    notes = (
        "讲解与会后阅读\n\n"
        "这一页讲 %s。顺序是：模型线（它到底有哪几个档位）→ 外壳（在哪跑、能干什么）"
        "→ 强项与边界（什么时候别用它）→ 在本分享里出现过的地方。\n\n"
        "模型线：%s\n\n"
        "外壳：%s\n\n"
        "强项：%s\n\n"
        "边界：%s\n\n"
        "在本分享：%s\n\n"
        "收尾一句：%s\n\n"
        "口径提醒：分数、价格、地区政策均为 2026-10-18 查询，活动与版本经常变，"
        "讲之前建议再核一次官网；本页不构成对任何厂商的推荐排序。"
        % (data["title"].split("  ", 1)[-1].strip(), data["models"], data["harness"],
           data["good"], data["bad"], data["here"], data["pick"])
    )
    slide.notes_slide.notes_text_frame.text = notes
    return slide


# ------------------------------------------------------------------ E0 matrix
def page_e0(prs, page):
    slide = prs.slides.add_slide(prs.slides[0].slide_layout)
    chrome(slide, str(page),
           "E0  按任务挑公司：六种活，各找各的人",
           "这十四家不是一条分数线上的名次，而是六种不同的活；先认活，再认公司",
           "来源：本分享第 02、03 部分的实测与公开榜单；各公司官网（2026-10-18）")

    rows = [
        ("改代码、动仓库", "Anthropic Claude", "OpenAI GPT · 智谱 GLM",
         "Opus 5.5 长任务不跑偏，Claude Code 的外壳最成熟"),
        ("看图纸、读图片", "Anthropic · OpenAI", "阿里 Qwen · Google",
         "能不能“真的看见图”决定结果；只靠 OCR 的一个都没找到"),
        ("写长文、查资料", "Google Gemini", "Meta Muse · 月之暗面 Kimi",
         "长输出与知识面；记得回原文核对，长文最容易“编得像”"),
        ("跑量、省钱", "DeepSeek", "智谱 GLM · MiniMax",
         "价格差一到两个数量级；任务边界写清楚，别指望它自己判断"),
        ("私有化、自部署", "智谱 GLM · Mistral", "DeepSeek · Kimi · Qwen",
         "权重能下载，数据不出机房；代价是算力与运维要自己扛"),
        ("常驻办事、盯流程", "xAI Grok Bot", "Meta Muse · Manus 类云端代理",
         "给它一台云端电脑和一个账号；权限给多大，风险就有多大"),
    ]
    # header
    rect(slide, L, 1.90, W, 0.40, C_HEAD)
    for x, w, t in [(0.75, 2.05, "要干的活"), (2.86, 2.40, "首选"),
                    (5.30, 2.40, "备选"), (7.80, 4.90, "为什么")]:
        put(slide, x, 1.90, w - 0.08, 0.40, t, size=12.5, bold=True, color=C_WHITE)

    y = 2.30
    for i, (a, b, c, d) in enumerate(rows):
        rect(slide, L, y, W, 0.60, C_WHITE if i % 2 == 0 else C_ZEBRA)
        put(slide, 0.75, y, 2.05, 0.60, a, size=12.5, bold=True, color=C_TITLE)
        put(slide, 2.86, y, 2.40, 0.60, b, size=11.5, bold=True, color=C_TEAL)
        put(slide, 5.30, y, 2.40, 0.60, c, size=11, color=C_BODY)
        put(slide, 7.80, y, 4.85, 0.60, d, size=11, color=C_MUTED)
        y += 0.60

    callout(slide, "怎么用", "这一节不是排行榜：每一页写的是“它能干什么、代价是什么、别拿它干什么”")
    slide.notes_slide.notes_text_frame.text = (
        "讲解与会后阅读\n\n"
        "这一页是 E 节的地图。不要按分数读，按“活”读：\n\n"
        "① 改代码、动仓库：Claude（终端外壳最成熟）；备选 GPT 与 GLM。\n"
        "② 看图纸、读图片：Claude 与 GPT；Qwen 会自己想办法，Google 知识面广。\n"
        "③ 写长文、查资料：Gemini 长输出稳；Muse 与 Kimi 在长上下文上有优势。\n"
        "④ 跑量、省钱：DeepSeek 最便宜；GLM 与 MiniMax 各有便宜档。\n"
        "⑤ 私有化、自部署：GLM、Mistral、DeepSeek、Kimi、Qwen 都能下权重。\n"
        "⑥ 常驻办事、盯流程：Grok Bot 是“给每个 Bot 一台云端电脑”的形态。\n\n"
        "提醒听众：十四家里有两家（Mistral、百度文心）没有进入本次同口径对比，"
        "页面里已注明，不要拿它们的名字去对 P22–P24 的分数。"
    )
    return slide


# ------------------------------------------------------------ divider (clone)
def divider(prs, page, letter, name, note, items, current):
    src = prs.slides[54]                      # A divider: clean 5-item list
    slide = prs.slides.add_slide(src.slide_layout)
    set_bg(slide, C_HEAD)
    sp_src = src._element.find(qn("p:cSld")).find(qn("p:spTree"))
    sp_new = slide.shapes._spTree
    for child in list(sp_src):
        if child.tag in (qn("p:nvGrpSpPr"), qn("p:grpSpPr")):
            continue
        sp_new.append(copy.deepcopy(child))

    shapes = list(slide.shapes)
    # big letter
    tb = shapes[1]
    tb.text_frame.paragraphs[0].runs[0].text = letter
    # 附录 / 名称
    tb = shapes[2]
    tf = tb.text_frame
    tf.paragraphs[0].runs[0].text = "附录"
    p2 = tf.paragraphs[1]
    if p2.runs:
        p2.runs[0].text = name
    else:
        r = p2.add_run()
        r.text = name
    # note
    tb = shapes[3]
    tb.text_frame.paragraphs[0].runs[0].text = note

    # right-hand list: rebuild 5 boxes -> 6 boxes
    boxes = shapes[6:11]
    proto = copy.deepcopy(boxes[-1]._element)
    lst = slide.shapes._spTree
    lst.append(proto)
    boxes = list(slide.shapes)[6:12]
    for i, b in enumerate(boxes):
        b.top = Inches(2.06 + i * 0.50)
        b.height = Inches(0.42)
        tf = b.text_frame
        p = tf.paragraphs[0]
        for r in list(p.runs)[1:]:
            r._r.getparent().remove(r._r)
        if not p.runs:
            p.add_run()
        p.runs[0].text = items[i]
        on = items[i].startswith(current + " ")
        for r in p.runs:
            r.font.bold = on
            r.font.color.rgb = RGBColor.from_string("FFFFFF" if on else "E6ECF5")
    return slide


# ------------------------------------------------------------------ doc level
def renumber(prs, start_index, delta):
    for s in list(prs.slides)[start_index:]:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            if abs(sh.left - Inches(11.73)) < 20000 and abs(sh.top - Inches(0.32)) < 20000:
                tf = sh.text_frame
                txt = tf.text.strip()
                if txt.isdigit():
                    p = tf.paragraphs[0]
                    p.runs[0].text = str(int(txt) + delta)


def add_e_to_divider_lists(prs):
    """Existing A/B/C/D/Z dividers: 5-item list -> 6 items incl. E."""
    items = ["A  订阅与省流（七页）", "B  常用术语（四页）", "C  资讯渠道（四页）",
             "D  案例、风险与破限（十五页）", "E  一页一家公司（十五页）",
             "Z  参考资料与来源"]
    for slide in list(prs.slides):
        shapes = list(slide.shapes)
        if len(shapes) != 11:
            continue
        letter = shapes[1].text_frame.text.strip() if shapes[1].has_text_frame else ""
        if letter not in ("A", "B", "C", "D", "Z"):
            continue
        if not (shapes[10].has_text_frame and
                shapes[10].text_frame.text.strip().startswith("Z ")):
            continue
        proto = copy.deepcopy(shapes[10]._element)
        slide.shapes._spTree.append(proto)
        boxes = list(slide.shapes)[6:12]
        for i, b in enumerate(boxes):
            b.top = Inches(2.02 + i * 0.50)
            b.height = Inches(0.40)
            p = b.text_frame.paragraphs[0]
            for r in list(p.runs)[1:]:
                r._r.getparent().remove(r._r)
            if not p.runs:
                p.add_run()
            p.runs[0].text = items[i]
            on = items[i].startswith(letter + " ")
            for r in p.runs:
                r.font.bold = on
                r.font.color.rgb = RGBColor.from_string("FFFFFF" if on else "E6ECF5")


def add_e_to_toc(prs):
    slide = prs.slides[2]
    items = ["A 订阅与省流版", "B 常用术语", "C 资讯渠道",
             "D  案例、风险与破限", "E 一页一家公司（十五页）", "Z 参考资料来源"]
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith("A 订阅与省流版"):
            tf = sh.text_frame
            proto = copy.deepcopy(tf.paragraphs[0]._p)
            while len(tf.paragraphs) > 1:
                tf._txBody.remove(tf.paragraphs[-1]._p)
            for _ in range(5):
                tf._txBody.append(copy.deepcopy(proto))
            for i, p in enumerate(tf.paragraphs):
                for r in list(p.runs)[1:]:
                    r._r.getparent().remove(r._r)
                p.runs[0].text = items[i]
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == "跳到附录 ›":
            sh.top = Inches(5.30)
    # 面板高度恢复原状；条目文本块不变，六条在 2.65 起排
    for sh in slide.shapes:
        if abs(sh.left - Inches(9.30)) < 20000 and abs(sh.top - Inches(1.85)) < 20000:
            if sh.height > Inches(3.5) and not sh.has_text_frame:
                sh.height = Inches(3.94)


# ------------------------------------------------------------------------ main
def main():
    prs = Presentation(SRC)
    prs.slides.add_slide  # touch

    imgs = {k: prep(k, v) for k, v in {
        "gpt": "GPT.jpg", "claude": "CLADEU.jpg", "gemini": "GEMINI.jpg",
        "grok": "GROK.jpg", "muse": "MUSE.jpg", "mistral": "Mistral.jpg",
        "deepseek": "DEEPSEEK.jpg", "qwen": "QWEN.jpg", "glm": "GLM.jpg",
        "kimi": "KIMI.jpg", "minimax": "MINIMAX.jpg", "hy": "HY.jpg",
        "baidu": "文心一言.jpg"}.items()}
    logos = {k: os.path.join("work_e", "logo_final", k + ".png") for k in (
        "openai", "anthropic", "gemini", "grok", "meta", "mistral", "deepseek",
        "qwen", "glm", "kimi", "minimax", "hy", "baidu", "dots")}

    pages = []

    # ---------------- divider first (like A/B/C/D: 分隔页在板块最前)
    divider(
        prs, 89, "E", "一页一家公司",
        "E0 是一张“按任务挑公司”的对照表；E1–E14 每家一页：模型线、外壳、强项、边界，以及它在本次分享里出现过的地方。",
        ["A  订阅与省流（七页）", "B  常用术语（四页）", "C  资讯渠道（四页）",
         "D  案例、风险与破限（十五页）", "E  一页一家公司（十五页）",
         "Z  参考资料与来源"], "E")

    page_e0(prs, 90)

    data = []
    data.append(dict(
        logo=logos["openai"], art=imgs["gpt"], title="E1  OpenAI：把“推理”做成了可交付的东西",
        sub="同一家公司三个档位：Astra 打最难的题，Sol 干日常的活，Luna 守免费入口",
        models="GPT-6 Astra（旗舰，AA 53；FrontierMath Erdős 68 题里唯一得分，3%）；GPT-6.1 Sol（AA 52，每道题 $0.72，性价比最高的一档）；GPT-6 Luna（免费档的小模型，上下文 27K）。全部闭源。",
        harness="ChatGPT 网页 / 桌面 / 手机；Codex 一个名四种形态——终端、IDE 插件、桌面、云端，Astra 版的 Codex 外壳把 computer use 提速约 1.9 倍；另有 API 与 Azure。",
        good="推理与数学第一梯队；写代码偏“防御式”，兜底和异常处理写得多，改动面也大。本次鹈鹕 G 一次成片，图纸 2/3。",
        bad="贵：Astra 输入/输出 $10 / $50 每百万 token。档位最多，免费版和付费版常常不是同一个模型——拿免费版的体验评价它会低估它。",
        here="鹈鹕 G（GPT 6.1 Sol high · Codex，P18）；图纸 2/3、假错误 1 条（P31）；榜单与分层（P22–P24）；数学进展（P26）。",
        pick="要解“最难的那道题”选 Astra；日常写代码跑批，Sol 更划算",
        src="来源：openai.com 模型页与系统卡、artificialanalysis.ai、epoch.ai/frontiermath；2026-10-18 查询"))

    data.append(dict(
        logo=logos["anthropic"], art=imgs["claude"], title="E2  Anthropic：把“干活”当成主战场",
        sub="智能体与长任务目前最强的一档；贵，但一次做对的概率高",
        models="Claude Opus 5.5（2026-09 发布，AA 58 全场第一；$4/$20 每百万 token，缓存读 $0.20，比 Opus 5 便宜 40%）；Sonnet 5.5（56）、Haiku 5.5（43）；Fable 5.1（53，上一代旗舰，仍在跑）。全部闭源。",
        harness="Claude Code（终端 Agent，社区公认与 Opus 5.5 最搭的一套外壳）；Claude 桌面版（原 Cowork，直接读写本机文件）；claude.ai；API 与 AWS / GCP / Azure。",
        good="长任务不跑偏、会自己查证。本次图纸 3/3（十家里唯一全中）；修编报告 4 分钟出稿、金额全对，还联网核对了承建单位。",
        bad="没有免费版，Claude Pro $20 起；对中国大陆用户与中资实体有地区限制（见 A5）；最强档跑一道题约 $6。",
        here="鹈鹕 D（Opus 5.5 · Hark，P18）；图纸 3/3（P30）；修编报告 4 分钟（P35）；“脾气”页（P25）。",
        pick="交付物要经得起人核、又不想反复返工时选它",
        src="来源：anthropic.com/claude-opus-5-5、platform.claude.com 定价、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["gemini"], art=imgs["gemini"], title="E3  Google：知识面最广，长输出最稳",
        sub="一个模型三种外壳：App、Antigravity（桌面＋CLI）、云上 Vertex",
        models="Gemini 4 Argon（AA 53；LMArena 文字榜第一 1524；ARC-AGI-3 97.5%；单次输出可达 1M token）；Gemini 3.8 Flash（41，免费档主力）；Flash-Lite（更小更快）。",
        harness="Gemini App；Antigravity——2026-06-18 起 Gemini CLI 并入 Antigravity CLI，桌面版与命令行共用同一套 Agent 外壳；AI Studio / Vertex AI；Workspace 内嵌。",
        good="知识面、中文与长文写作；一次能吐很长，适合整篇报告、长文档改写与资料综述。",
        bad="智能体榜偏后（LMArena Agent 13.73%）；工具链刚换过一代，网上多半旧教程已经对不上；国内访问受限。",
        here="榜单与分层（P22–P24）；免费档 Flash-Lite（P20）；订阅与地区（A2、A4）。",
        pick="写长文、查资料、要一个“什么都懂一点”的默认助手时选它",
        src="来源：blog.google、deepmind.google 模型卡、arena.ai、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["grok"], art=imgs["grok"], title="E4  xAI：模型不追第一，产品先做“常驻同事”",
        sub="Grok 4.7 在第二梯队；真正的差异是 Grok Bot 这套常驻 Agent",
        models="Grok 4.7（AA 46，第二梯队）；Grok 系列主打实时信息与 X 上的语料；权重不开源。模型本身不是它最强的地方。",
        harness="Grok Bot（2026-08-11 公测）：每个 Bot 一台自己的云端电脑，能登录你常用的应用，24 小时跑，只在需要拍板时回来；可多 Bot 并行、互相发消息、拉群协作；用量与 Grok / Cursor 订阅分开算。另有 grok.com、X 内入口、API。",
        good="把“人等结果”变成“结果等人”；适合销售跟进、运营、招聘、修 bug 这类跨应用的多步流程。",
        bad="地区限制，大陆不可直接访问；企业版还在等待名单；常驻意味着要把账号权限交出去很大一块。",
        here="中转与代理池（P77–P82）；支持地区（A2）；Grok 并入 X 订阅（A4、A7）。",
        pick="已经有 X 或 Cursor 订阅、又想试“常驻同事”形态时，它门槛最低",
        src="来源：x.ai/news/introducing-grok-bot、x.ai 订阅页、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["meta"], art=imgs["muse"], title="E5  Meta：把“记忆”当成模型能力来做",
        sub="Muse Wonder 是 Meta 第一款前沿模型；主打个人助理，不是编程",
        models="Muse Wonder 与 Muse Wonder Full（2026-10-12，Meta Superintelligence Labs 出品，AA 52；LMArena 文字榜第二）；Muse Spark（2026-04 首款，现支撑 Meta AI）。主打 Active Memory（跨对话记住你）与 Active Focus（长任务不放弃）。",
        harness="Muse App（muse.ai）；Muse Code（面向订阅者）；Meta Model API；Meta AI 内嵌于 Facebook / Instagram / WhatsApp。",
        good="长期记忆与个人事务：多约束的行程、比价、定制、排座位这类活，它记得住前因；写作体感好。",
        bad="只对美国、加拿大开放（第三方核查口径）；免费档有上限，Power $20、Maximum $100；企业级治理弱于 OpenAI / Anthropic。",
        here="榜单与分层（P22–P24）；“云端 / 个人代理”分类（P21）；价格（A4）。",
        pick="想试“记得住你”的个人助理时；工作交付别把它当主力",
        src="来源：research.meta.ai/blog/muse-wonder、about.fb.com、techcrunch.com、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["mistral"], art=imgs["mistral"], title="E6  Mistral：欧洲的“可自部署”路线",
        sub="唯一把开放权重当主业的头部；模型不最强，但能装进自己机房",
        models="Mistral Large 4（2026-10-06，约 1T 参数 MoE，1M 上下文，多模态，开放权重先 API 预览）；另有 Pixtral（视觉）、Devstral（代码）、Magistral（推理）。",
        harness="Mistral Vibe（原 Le Chat，聊天＋编码 Agent）；Mistral Code（IDE 插件）；La Plateforme API；权重可下载，用 vLLM / TensorRT-LLM 自建。",
        good="数据不出境、可审计、能进内网；同能力下显存与价格友好，是欧盟“主权 AI”的默认选项。",
        bad="绝对能力不在第一梯队，本次榜单没有收录它的旗舰分数；中文语料与国内生态明显弱。",
        here="本次未进入对比（P18–P28）——它的价值在“能自己部署”，不在跑分。",
        pick="单位要求数据不出机房、或要一个可审计的欧洲供应商时",
        src="来源：mistral.ai/news/mistral-large-4、docs.mistral.ai 模型与更新记录；2026-10-18 查询"))

    data.append(dict(
        logo=logos["deepseek"], art=imgs["deepseek"], title="E7  DeepSeek：把价格打到地板的那一家",
        sub="开源、便宜、快；代价是提示词要写细，识图别抱期望",
        models="DeepSeek V4.1 Flash 1017（2026-10-17 上线，1M 上下文，开源，官方自报 DeepSWE 77.0；AA 尚未出分，旧版 39）；V4 Pro 为高配档。",
        harness="网页 / App；API（OpenAI 兼容，改个 base_url 就能接）；第三方客户端（WorkBuddy、Cherry Studio 等）；权重可自建。",
        good="价格最低（缓存命中 0.04 元、未命中 2 元、输出 8 元每百万 token，闲时半价）、输出最快（约 217 token/s）；报告案例里原因写得最具体。",
        bad="提示词写得粗就明显变差；图纸只 1/3、19 条假错误；高峰时段（工作日 9–12、14–18 点）限速。",
        here="鹈鹕 C（DeepSeek V4.1 Flash 1M · WorkBuddy，P18）；图纸 1/3（P31）；修编报告 34 分钟（P35）；价格（P22）。",
        pick="量大、预算紧、任务边界清楚（抽取、改写、批处理）时首选",
        src="来源：api-docs.deepseek.com 定价与 1017 发布说明、huggingface.co/deepseek-ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["qwen"], art=imgs["qwen"], title="E8  阿里 Qwen：会自己想办法的那一家",
        sub="识图链路最“肯折腾”；Qoder 是它自家的 IDE 外壳",
        models="Qwen3.8 Max（2026-08，2.4T 参数 MoE，思考 / 快速双模式，AA 45）；Qwen3.8 Flash（41）与 Flash-Next；开源系列陆续放出权重。",
        harness="通义 App；Qoder（IDE，可接自己的 API Key，有 Quest 与 Editor 两种形态）；阿里云百炼 API；开源权重自建。",
        good="看不见图会自己把 PDF 渲染成图再逐格读——本次因此找到 K2；中文理解与写作稳。",
        bad="慢：图纸 43 分钟；假错误 24 条，报得多不等于对得多；Qoder 的免费额度与国内版价格要分开看。",
        here="鹈鹕 I（Qwen 3.8 Flash · Qoder，P18）；图纸 2/3；读图方式对比（P32）。",
        pick="需要“自己想办法把活干完”、又能等的时候；赶时间别选",
        src="来源：aliyun.com 百炼文档、qoder.com/pricing、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["glm"], art=imgs["glm"], title="E9  智谱 GLM：开源里最会写代码的一档",
        sub="便宜、可自建、编程强；识图和长任务仍然要人盯",
        models="GLM-5.3（开源权重；官方称编程较 GLM-5.2 提升 50%，Terminal-Bench 3.0 开源 SOTA；AA 45）；GLM-5.3-Flash（轻量版，本次鹈鹕用的就是它）。",
        harness="Z.ai 网页；ZCode（官方 Agent 外壳，多智能体协作）；GLM Coding Plan（包月，比按量 API 便宜）；bigmodel.cn API；权重可自建。",
        good="单位里想“私有化 + 便宜 + 会写代码”，它是目前最省事的一档；包月方案对天天写代码的人很划算。",
        bad="本次图纸看了 1 小时 24 分，最后判“无实质错误”，0/3；预览版与正式版能力有差距。",
        here="鹈鹕 A（GLM 5.3 Flash · Trae，P18）；图纸 0/3（P30）；分层（P24）。",
        pick="要私有化部署、或按包月写代码时",
        src="来源：z.ai/blog/glm-5.3、docs.bigmodel.cn、zcode.z.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["kimi"], art=imgs["kimi"], title="E10  月之暗面 Kimi：参数最大，长文最舒服",
        sub="2.8 万亿参数的开源旗舰；长上下文是它的看家本领",
        models="Kimi K3（2026-07，2.8T 参数、1M 上下文、开源，Kimi Delta Attention 混合线性注意力；AA 44）；K2 系列仍在服务。是目前参数最大的开源模型。",
        harness="Kimi 网页 / App；Kimi CLI 与 Kimi Code CLI（TypeScript 重写、单文件二进制）；platform.kimi.com API；权重可自建。",
        good="长文档、长对话不丢线；前端与设计稿还原在社区口碑好，适合读规范、读合同、写长稿。",
        bad="智能体干活不在第一梯队；本次没有进入识图对比，缺同口径数据。",
        here="榜单与分层（P22–P24，AA 44 与 Step 5 并列）。",
        pick="读长材料、写长文、还原前端稿时",
        src="来源：kimi.com/news/kimi-k3、github.com/MoonshotAI/Kimi-K3、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["minimax"], art=imgs["minimax"], title="E11  MiniMax：一边做语言，一边做视频",
        sub="语言模型在追赶，视频与多模态才是它真正的招牌",
        models="MiniMax M3（2026-06，约 428B 参数 / 23B 激活，1M 上下文，原生多模态，开源；AA 29）；H3（即海螺 3.0，2026-07-31，音画同出的视频模型）。",
        harness="MiniMax Agent（桌面端，能读本地文件、操作浏览器）；MiniMax Code；海螺视频；开放平台 API。",
        good="视频与多模态生成是强项；桌面 Agent 直接动本机文件，做素材、配音、短片比纯聊天模型顺手。",
        bad="语言模型分数偏低（AA 29）；本次图纸 0/3、10 条假错误。",
        here="鹈鹕 J（MiniMax M3 · WorkBuddy，P18）；图纸 0/3（P31）；分层（P24）。",
        pick="要做视频、配音、多模态素材时；当主力语言模型还早",
        src="来源：minimax.io/blog/minimax-m3、hailuoai.video、artificialanalysis.ai；2026-10-18 查询"))

    data.append(dict(
        logo=logos["hy"], art=imgs["hy"], title="E12  腾讯混元 Hy：办公场景最“顺手”的国产一档",
        sub="模型开源、客户端齐全；WorkBuddy 与 CodeBuddy 积分共享",
        models="Hy4 preview（770B 总参数 / 49B 激活，1M 上下文，开源；AA 未收录，LMArena 智能体榜第 20）；另有混元图像与视频模型。",
        harness="元宝（聊天）；ima（知识库）；WorkBuddy（桌面办公 Agent，直接读写本机 Excel、PDF）；CodeBuddy（IDE）；腾讯云 API。",
        good="国内可人民币付款，免费积分够跑通一个真实小任务；办公文件处理链路最完整，从读表到出稿一条龙。",
        bad="预览版能力不稳定；本次它自述环境不支持视觉，图纸 0/3，不计入识图能力比较。",
        here="鹈鹕 E（HY4 Preview 1M high · WorkBuddy，P18）；图纸 0/3 且视觉未启用（P31）；免费额度（P28）。",
        pick="单位在国内、要处理本机 Office 文件、又不想折腾海外账号时",
        src="来源：tencent.com 与 hunyuan.tencent.com 的 Hy4 发布页、huggingface.co/tencent/Hy4-preview、workbuddy.cn；2026-10-18 查询"))

    data.append(dict(
        logo=logos["baidu"], art=imgs["baidu"], title="E13  百度文心：中文搜索与生态最深的一档",
        sub="参数很大、成本很低；但公开榜单上很难横向比",
        models="文心 5.0（2.4T 参数，原生全模态，2026-01 正式版）；文心 5.1（2026-05，官方称预训练成本约为同规模模型的 6%）；ERNIE 系列小模型陆续开放。",
        harness="文心一言 App / 网页；文心快码 Comate（IDE 插件，Spec 模式、Figma 转代码）；千帆平台 API；百度搜索内嵌。",
        good="中文时事与搜索时效最好；与百度生态（搜索、地图、网盘）打通；企业侧的合规与备案材料齐全。",
        bad="缺同口径第三方分数，本分享的榜单里没有它，不能与前十三家直接比高低。",
        here="未进入本次对比（P18–P28）——榜单缺同口径数据，不硬凑。",
        pick="要中文搜索时效、或单位已经在用百度云时",
        src="来源：ernie.baidu.com、cloud.baidu.com 千帆文档、comate.baidu.com；2026-10-18 查询"))

    data.append(dict(
        logo=logos["dots"], art=None, word="dots",
        phnote="AI娘 形象待补：\nAI娘 目录里还没有 dots 的图，补一张后替换本框即可",
        title="E14  小红书 dots：一条窄而深的路线",
        sub="不做通用助手，先把数学与“真实笔记”这两件事做到顶",
        models="dots-note-3.0（IMO 2026 官方评卷 42/42，满分金牌）；dots3-note preview（开源：280B 总参数 / 16B 激活，512K 上下文，文本＋视觉＋语音）；dots.vlm1（多模态）。",
        harness="点点 App（AI 搜索与生活助手，2026-09 接替站内“问一问”）；小红书站内 AI 搜索；开源权重可自建。没有 CLI、没有桌面 Agent。",
        good="数学推理的公开成绩最硬——官方评卷满分，不是厂商自测；语料来自真实笔记，生活类问答接地气。",
        bad="通用 Agent 外壳基本空白；长任务、文件处理、代码都不是它的战场；模型更新节奏慢。",
        here="数学页脚注提到 IMO 42/42（P26）；正文没有单独评测它。",
        pick="做数学 / 竞赛题、或做生活类中文问答时；当办公助手不合适",
        src="来源：kimi.com 与 IMO 2026 公开评分、huggingface.co 模型卡、小红书 Dots 团队发布；2026-10-18 查询"))

    for i, d in enumerate(data):
        company_page(prs, 91 + i, d)

    # ---------------- move the 16 new slides in front of the Z divider
    lst = prs.slides._sldIdLst
    ids = list(lst)
    new = ids[-16:]
    for e in new:
        lst.remove(e)
    anchor = ids[88]
    for i, e in enumerate(new):
        lst.insert(88 + i, e)

    # ---------------- page numbers, existing dividers, TOC
    add_e_to_divider_lists(prs)
    add_e_to_toc(prs)
    renumber(prs, 104, 16)

    prs.save(OUT)
    print("saved", OUT, len(prs.slides.__iter__.__self__._sldIdLst), "slides")


if __name__ == "__main__":
    main()
