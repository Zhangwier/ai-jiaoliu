# -*- coding: utf-8 -*-
"""
第四章开头新增两页（插在章节过渡页之后、五种方法之前）：
  新 40：先说我日常：三件具体活，都交给它做
  新 41：拿"找图纸"举例：先建索引，再检索
输出到 AI交流分享-成品26.pptx（在已有 122 页版本上增量插入）
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

SRC = r"D:\MG\program\ai交流\AI交流分享-成品26.pptx"
OUT = r"D:\MG\program\ai交流\AI交流分享-成品26.pptx"
FONT = "微软雅黑"
MONO = "Consolas"

C_TITLE = "14213D"
C_TXT = "2B3445"
C_MUTED = "6B7385"
C_TEAL = "0F766E"
C_ORANGE = "D9622B"
C_BG_SOFT = "E1F0EE"
C_BG_WARN = "FBE9DE"
C_BG_BEIGE = "ECEAE4"
C_DARK = "12203A"
C_WHITE = "FFFFFF"
C_PAGE_BG = "F6F5F1"
C_LINE = "DAD7CE"

NSMAP = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
         'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')


def set_fill(shape, hex_color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(hex_color)


def no_line(shape):
    shape.line.fill.background()


def rect(slide, x, y, w, h, color, border=None, border_pt=0.8):
    sh = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    set_fill(sh, color)
    if border:
        sh.line.color.rgb = RGBColor.from_string(border)
        sh.line.width = Pt(border_pt)
    else:
        no_line(sh)
    return sh


def tb(slide, x, y, w, h, text, size=12, bold=False, color=C_TXT,
       align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.25, mono=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = 0
    tf.margin_right = 0
    tf.margin_top = 0
    tf.margin_bottom = 0
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        r = p.add_run()
        r.text = line
        f = r.font
        f.name = MONO if mono else FONT
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = RGBColor.from_string(color)
        rPr = r._r.get_or_add_rPr()
        if rPr.find(qn("a:ea")) is None:
            rPr.append(parse_xml('<a:ea %s typeface="%s"/>' % (NSMAP, FONT)))
        if rPr.find(qn("a:cs")) is None:
            rPr.append(parse_xml('<a:cs %s typeface="%s"/>' % (NSMAP, FONT)))
    return box


def set_bg(slide, hex_color):
    cSld = slide._element.find(qn("p:cSld"))
    for old in cSld.findall(qn("p:bg")):
        cSld.remove(old)
    cSld.insert(0, parse_xml(
        '<p:bg %s><p:bgPr><a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>' % (NSMAP, hex_color)))


def chrome(slide, page_no, source_text):
    tb(slide, 0.6, 0.32, 6.0, 0.3, "04  怎么用得稳", 12, True, C_TEAL)
    tb(slide, 11.73, 0.32, 1.0, 0.3, str(page_no), 12, False, C_MUTED, align=PP_ALIGN.RIGHT)
    rect(slide, 0.6, 0.66, 0.5, 0.06, C_TEAL)
    ln = slide.shapes.add_connector(1, Inches(0.6), Inches(7.0), Inches(12.73), Inches(7.0))
    ln.line.color.rgb = RGBColor.from_string("E4E7EB")
    ln.line.width = Pt(0.5)
    tb(slide, 0.6, 7.06, 10.5, 0.3, source_text, 8.5, False, "94A3B8")
    tb(slide, 10.73, 7.06, 2.0, 0.3, "来源索引 ›", 8.5, False, "94A3B8")


# ============================================================ 第 40 页
def build_overview(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, C_PAGE_BG)
    chrome(s, 0, "来源：本地日常任务记录（2026-09～10）：修编清单公式、TJ1 文件索引、征地审计与修编报告")

    tb(s, 0.6, 0.74, 12.13, 0.58,
       "先说我日常：三件具体活，都交给它做", 27, True, C_TITLE,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    tb(s, 0.6, 1.36, 12.13, 0.36,
       "不是我让它“表演能力”，是我手上正好有这类活——格式定死、答案唯一、做完能当场验",
       16.5, False, C_MUTED, spacing=1.1)

    cards = [
        ("01", "写 Excel 公式",
         "把口径讲成一句人话：“在 T 列找最后一次出现的相同条件，再把对应行的 D 列加起来”——它给一整行 LET 公式",
         "=LET(当前行,ROW(), … SUMPRODUCT(求和范围,--(条件=T994),--(ROW(求和范围)>=上一个行)))",
         "_TJ1 台账 · 合价重算列",
         "先在一行上和手算对一个数，再往下拖"),
        ("02", "把一堆图纸变成能查的表",
         "让它扫目录、按册和出图编号归类，生成索引；以后找图就是查一行，表里没有的直接开口问",
         "_TJ1文件索引.xlsx · 101 个文件 · 按册 / 出图编号 / 类型 / 相对路径",
         "TJ1（K0+000～K10+870）",
         "随便点几个路径，确认指得到真文件"),
        ("03", "按模板填报告",
         "给模板、给定稿清单、给复算底稿，它按章节往下填；金额和条目数由脚本算，不让它算",
         "29.化北高速…补偿审核意见.docx / TJ5 合同段 0# 变更修编报告.docx",
         "征地审计 · 清单修编",
         "金额复算、原因回原文、旧模板残留"),
    ]
    x = 0.6
    for num, head, how, material, tag, check in cards:
        rect(s, x, 1.94, 3.93, 4.06, C_WHITE, border=C_LINE)
        rect(s, x, 1.94, 3.93, 0.07, C_TEAL)
        tb(s, x + 0.22, 2.1, 0.6, 0.28, num, 11.5, True, C_TEAL)
        tb(s, x + 0.22, 2.4, 3.5, 0.52, head, 16.5, True, C_TITLE, spacing=1.1)

        tb(s, x + 0.22, 2.98, 3.5, 0.24, "我怎么用它", 10, True, C_MUTED)
        tb(s, x + 0.22, 3.22, 3.5, 0.92, how, 11, False, C_TXT, spacing=1.25)

        rect(s, x + 0.22, 4.2, 3.5, 0.62, C_BG_BEIGE)
        is_mono = material.startswith("=")
        tb(s, x + 0.32, 4.24, 3.3, 0.56, material, 8 if is_mono else 8.5, False,
           C_DARK, spacing=1.15, mono=is_mono)
        tb(s, x + 0.22, 4.9, 3.5, 0.22, tag, 9, False, C_MUTED)

        rect(s, x + 0.22, 5.18, 3.5, 0.62, C_BG_SOFT)
        tb(s, x + 0.32, 5.2, 3.3, 0.2, "我核什么", 9, True, C_TEAL)
        tb(s, x + 0.32, 5.42, 3.3, 0.34, check, 10, False, C_TXT, spacing=1.1)
        x += 4.1

    rect(s, 0.6, 6.3, 12.13, 0.55, C_BG_WARN)
    rect(s, 0.6, 6.3, 0.09, 0.55, C_ORANGE)
    tb(s, 0.85, 6.3, 1.05, 0.55, "共同点", 12.5, True, C_ORANGE,
       anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 1.85, 6.3, 10.65, 0.55,
       "都不是“让它发挥”，而是我把话说死，它把活干完——所以做完我一眼能看出对不对",
       17.5, True, C_TITLE, anchor=MSO_ANCHOR.MIDDLE)

    s.notes_slide.notes_text_frame.text = """讲解与会后阅读

先说我日常：三件具体活

这一页的作用是把“AI 能干什么”换成“我拿它干了什么”，避免听众停在抽象层面。三件事都是真实做过、结果可核的。

1. 写 Excel 公式。截图那行是 LET + SUMPRODUCT 组合，用来替代按条件回看的 SUMIF：先定位当前行、再向前找最近一次同条件的行，然后把区间内的单价行加起来。这类任务 AI 很强——语法它比我熟，口径要我说清。我核的方式很简单：先在一行上和手算对一遍，对上了再往下拖。

2. 把一堆图纸变成能查的表。TK1 标段 101 个文件散在十几个分册里，靠翻目录找一个数量表要花很久。做法是让它（或它写的脚本）扫一遍目录，按册、出图编号、类型、相对路径生成索引表；以后找图就是筛选一行。表里没收录的，直接问它“有没有 XX 部位的图”，它会去翻索引和目录文件。我核的是路径能不能点开。

3. 按模板填报告。征地审计意见和 0# 变更修编报告，格式都是死的：章节固定、必填项固定、金额来自清单和复算底稿。这类活推理不复杂，但它不会累、不会漏章节。我核三件事：金额复算、原因能否回到函件或勘误表、旧模板残留（比如承建单位、月份）。

三件事的共同点：格式固定、有唯一正确答案、做完能当场验。这三条同时也是后面“什么样的活该写死、什么样的活该留白”的判断依据。

依据：本地日常任务记录（2026-09～10）。

查询与整理日期：2026-10-19。"""
    return s


# ============================================================ 第 41 页
def build_index_case(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s, C_PAGE_BG)
    chrome(s, 0, "来源：_TJ1文件索引.xlsx（2026-10-14 生成）；make_file_index.py；南湛高速 TJ1 标段施工图")

    tb(s, 0.6, 0.74, 12.13, 0.58,
       "拿“找图纸”举例：先建索引，再检索", 27, True, C_TITLE,
       anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
    tb(s, 0.6, 1.36, 12.13, 0.36,
       "TJ1 标段 101 个文件散在十几个分册里；先让它扫成一张表，之后每次找图只要筛一行",
       16.5, False, C_MUTED, spacing=1.1)

    # ---- 左：索引表长什么样 ----
    tb(s, 0.6, 1.86, 7.3, 0.3, "生成出来的表长什么样", 15, True, C_TEAL)

    rect(s, 0.6, 2.2, 7.3, 0.36, C_DARK)
    heads = [("序号", 0.72, 0.5), ("文件名", 1.24, 2.5), ("类型", 3.78, 1.0),
             ("出图编号", 4.82, 1.0), ("所在册", 5.86, 1.9)]
    for label, hx, hw in heads:
        tb(s, hx, 2.2, hw, 0.36, label, 9.5, True, C_WHITE, anchor=MSO_ANCHOR.MIDDLE)

    rows = [
        ("3", "005_控制成果表.xlsx", "表格", "S1-2-2", "2.0 第二册 路线"),
        ("6", "010_逐桩坐标表.xlsx", "坐标表", "S1-2-7", "2.0 第二册 路线"),
        ("8", "013_砍树挖根数量表.xls", "数量表", "S1-2-10", "2.0 第二册 路线"),
        ("…", "……", "", "", ""),
    ]
    y = 2.56
    for i, (a, b, c, d, e) in enumerate(rows):
        fill = C_WHITE if i % 2 == 0 else C_BG_BEIGE
        rect(s, 0.6, y, 7.3, 0.34, fill)
        tb(s, 0.72, y, 0.5, 0.34, a, 9, False, C_MUTED, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 1.24, y, 2.5, 0.34, b, 9, False, C_TXT, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 3.78, y, 1.0, 0.34, c, 9, False, C_TXT, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 4.82, y, 1.0, 0.34, d, 9, False, C_TXT, anchor=MSO_ANCHOR.MIDDLE)
        tb(s, 5.86, y, 1.9, 0.34, e, 9, False, C_TXT, anchor=MSO_ANCHOR.MIDDLE)
        y += 0.34

    tb(s, 0.6, 4.0, 7.3, 0.5,
       "实际表共 9 列：另含子目录、大小、修改时间、相对路径（可点击直接打开）",
       10, False, C_MUTED, spacing=1.1)

    rect(s, 0.6, 4.58, 7.3, 0.92, C_BG_BEIGE)
    tb(s, 0.78, 4.64, 6.95, 0.24, "另有一张「册汇总」，先看它决定翻哪一册", 10, True, C_TITLE)
    tb(s, 0.78, 4.9, 6.95, 0.54,
       "2.0 第二册 路线 · 26 个文件 · 0.98 MB · 9 个数量表      "
       "3.1 第三册 第一分册 路基 · 18 个 · 4.12 MB · 11 个数量表",
       9, False, C_TXT, spacing=1.15)

    # ---- 右：三种查法 ----
    tb(s, 8.1, 1.86, 4.65, 0.3, "学会用之后，找图变成三件事", 15, True, C_TEAL)

    uses = [
        ("按册筛", "“第三册第一分册有哪些数量表”——筛一列就出来"),
        ("按出图编号定位", "图纸目录上写 S1-2-7，直接查这一列，落到具体文件"),
        ("按类型找同类", "要所有“数量表”，筛类型；要所有坐标表，同理"),
    ]
    y = 2.2
    for head, body in uses:
        rect(s, 8.1, y, 0.06, 0.78, C_TEAL)
        tb(s, 8.26, y - 0.04, 4.4, 0.28, head, 12, True, C_TITLE)
        tb(s, 8.26, y + 0.26, 4.4, 0.54, body, 10.5, False, C_TXT, spacing=1.15)
        y += 0.88

    rect(s, 8.1, 4.80, 4.65, 0.76, C_BG_WARN)
    rect(s, 8.1, 4.80, 0.06, 0.76, C_ORANGE)
    tb(s, 8.26, 4.84, 4.35, 0.24, "表里没有的，直接问它", 10.5, True, C_ORANGE)
    tb(s, 8.26, 5.1, 4.35, 0.44,
       "“有没有 XX 桥台的挡块图”——它会去翻目录和分册表，再告你在哪本",
       10, False, C_TXT, spacing=1.1)

    # ---- 结论 ----
    rect(s, 0.6, 5.88, 12.13, 0.32, C_BG_SOFT)
    rect(s, 0.6, 5.88, 0.06, 0.32, C_TEAL)
    tb(s, 0.78, 5.88, 11.8, 0.32,
       "这就是最土的 RAG：先建索引，再检索。好处是可核对、可重生成——换项目只要改一个路径参数重跑脚本",
       10.5, False, C_TXT, anchor=MSO_ANCHOR.MIDDLE)

    rect(s, 0.6, 6.3, 12.13, 0.55, C_BG_WARN)
    rect(s, 0.6, 6.3, 0.09, 0.55, C_ORANGE)
    tb(s, 0.85, 6.3, 1.05, 0.55, "结论", 14, True, C_ORANGE, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, 1.75, 6.3, 10.73, 0.55,
       "它的检索能力，要一张自己能核对的表来托底",
       17.5, True, C_TITLE, anchor=MSO_ANCHOR.MIDDLE)

    s.notes_slide.notes_text_frame.text = """讲解与会后阅读

拿“找图纸”举例：先建索引，再检索

1. 痛点：TJ1 标段（K0+000～K10+870）101 个文件，散在第二册到第七册的十几个分册里。找一张数量表，要么翻目录 Excel，要么按文件名猜，实际常常要花十几分钟。

2. 做法：先扫一遍根目录，生成 _TJ1文件索引.xlsx。索引表 9 列：序号、文件名、类型、出图编号、所在册/分册、子目录、大小、修改时间、相对路径。噪声文件（CAD 备份 .bak、临时文件）跳过；出图编号来自各册《目录》文件，目录没提供的就留空。

3. 另一张「册汇总」：按册统计文件数、体量、数量表/设计图/文档/PDF 各几个。先看它能判断该翻哪一册，不用一个个点开。

4. 用法：按册筛、按出图编号定位、按类型找同类。表里没收录的（比如某座桥的挡块图还没扫描进来），直接问它——它会去翻目录和分册表，再告诉你大概在哪一本。

5. 这就是最简单的 RAG 思路：先把资料整理成可检索的结构，再让模型在命中的片段上回答，而不是把整个文件夹塞进去。区别在于这里检索的是文件名和出图编号，不是向量。

6. 可复现：脚本 make_file_index.py，换项目只改一个根目录参数。这也是它的价值——索引能重生成、能核对，不是一次性手工整理的死表。

依据：_TJ1文件索引.xlsx（生成时间 2026-10-14 11:29，共 101 个文件）；make_file_index.py；南湛高速广东段 TJ1 标段两阶段施工图。

查询与整理日期：2026-10-19。"""
    return s


def main():
    prs = Presentation(SRC)
    build_overview(prs)
    build_index_case(prs)

    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    new2 = ids[-1]
    new1 = ids[-2]
    sldIdLst.remove(new2)
    sldIdLst.remove(new1)
    # 第 39 页是 04 章节过渡页（0-based 38），插在它之后
    sldIdLst.insert(39, new1)
    sldIdLst.insert(40, new2)

    fixed = 0
    for i, sl in enumerate(prs.slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame:
                continue
            if not sh.text_frame.text.strip().isdigit():
                continue
            if sh.left is None or sh.top is None:
                continue
            if sh.left / 914400 > 11.4 and sh.top / 914400 < 0.6:
                for p in sh.text_frame.paragraphs:
                    for r in p.runs:
                        if r.text.strip().isdigit():
                            r.text = str(i)
                fixed += 1
    print("page numbers:", fixed)
    prs.save(OUT)
    print("saved", OUT)


if __name__ == "__main__":
    main()
