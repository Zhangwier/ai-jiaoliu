# -*- coding: utf-8 -*-
"""ep3 做表：把「同一份图纸、三种交办方式」的实测结果插进第 03 节。
在成品18 基础上生成成品19：新增 6 页（P39–P44），「两个真实案例」改为「三个真实案例」。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from build_v3 import (put, rect, hline, dot, rowband, set_bg, textbox,
                      style_run, link_source_index,
                      L, W, C_TITLE, C_BODY, C_MUTED, C_TEAL, C_ORANGE,
                      C_HEAD, C_LINE, C_WHITE, C_WARM, C_COOL, C_ZEBRA)

SECTION = '03  三个真实案例'


def shell(slide, page, title, subtitle, source, notes):
    set_bg(slide)
    put(slide, L, 0.32, 6.00, 0.30, SECTION, size=12, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    put(slide, 11.73, 0.32, 1.00, 0.30, str(page), size=12, color=C_MUTED,
        align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
    rect(slide, L, 0.66, 0.50, 0.06, C_TEAL)
    put(slide, L, 0.78, W, 0.62, title, size=25, bold=True, color=C_TITLE,
        spacing=1.0)
    if subtitle:
        put(slide, L, 1.50, W, 0.42, subtitle, size=16.5, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP, spacing=1.0)
    hline(slide, L, 7.00, W)
    put(slide, L, 7.06, 10.50, 0.30, source, size=10.5, color=C_MUTED,
        anchor=MSO_ANCHOR.TOP)
    put(slide, 10.73, 7.06, 2.00, 0.30, '来源索引 ›', size=10.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    slide.notes_slide.notes_text_frame.text = notes


def band(slide, y, label, body, kind='warn', h=0.55):
    bar = C_WARM if kind == 'warn' else C_COOL
    edge = C_ORANGE if kind == 'warn' else C_TEAL
    rect(slide, L, y, W, h, bar)
    rect(slide, L, y, 0.09, h, edge)
    put(slide, 0.85, y, 1.05, h, label, size=13.5, bold=True, color=edge)
    put(slide, 1.76, y, 10.70, h, body, size=15.5, bold=True, color=C_TITLE,
        spacing=1.0)


def table(slide, x, y, w, cols, rows, widths, rh=0.52, hdr=0.40,
          fs=12, hfs=12):
    rect(slide, x, y, w, hdr, C_HEAD)
    cx = x
    for label, frac in zip(cols, widths):
        cw = w * frac
        put(slide, cx + 0.14, y, cw - 0.20, hdr, label, size=hfs, bold=True,
            color=C_WHITE)
        cx += cw
    yy = y + hdr
    for i, row in enumerate(rows):
        rowband(slide, x, yy, w, rh, i)
        cx = x
        for j, (cell, frac) in enumerate(zip(row, widths)):
            cw = w * frac
            put(slide, cx + 0.14, yy + 0.02, cw - 0.22, rh - 0.04, cell,
                size=fs, color=C_BODY if j else C_TITLE, bold=(j == 0),
                spacing=1.08, anchor=MSO_ANCHOR.MIDDLE)
            cx += cw
        yy += rh
    return yy


# ============================================================ P39 实验设置
def p1(slide, page):
    shell(
        slide, page,
        '案例三：同一张 326 页图纸，三种把任务交出去的方式',
        '同一模型（Trae · DeepSeek）、新会话、同一份 PDF 与模板；差别只在任务书怎么写、要不要拆阶段',
        '来源：ep3 做表/提示词.txt；三组沙箱落盘产物与时间戳；输入文件哈希核对',
        '第 39 页 · 先说清实验，再揭结果\n\n'
        '这一页只交代实验，不评价谁好。三组用的是同一份 326 页图纸（主线G6-2 新沙高架桥）'
        '和同一份计算稿模板（潭州高架桥），要求都是：按厚度汇总，交付可用的 Excel 底稿。\n\n'
        'A 组只写一件事：按图纸填模板、尽量沿用原有 sheet 和公式。B 组把规则写死：'
        '保留合并单元格与样式、不插删行、单件重与合计必须公式、LibreOffice 重算、'
        '图纸数量与复核数量分开、缺失不许编。C 组把任务拆成七个阶段，每阶段落盘并写 handoff。\n\n'
        '输入一致性我逐字节核过：PDF sha256 4e1a80ff…，模板 34509914…，三个沙箱完全相同，'
        '不存在「换了图纸」这种解释。\n\n'
        '耗时来自各组自己记录的文件时间戳，不是插桩计时，只能看量级：A 64 分钟、B 39 分钟、'
        'C 178 分钟（其中阶段 B 的表格页分离就占 55 分钟）。\n\n'
        '现场可以先问一句：你觉得哪一份最接近能用？先别答，下一页看三份文件到底各做完了什么。')
    cards = [
        ('A', '简短交办', '一句话说目标',
         ['只说明按图纸填模板、尽量沿用原有 sheet 与公式',
          '方法、范围、验收标准都留给模型自己定',
          '不给阶段要求，一次对话做到交文件'],
         '记录耗时 64 分钟', C_TEAL),
        ('B', '详细约束', '把规则写死',
         ['保留 sheet、表头、合并、边框、字体、列宽与公式',
          '不插删行；必须重算并保存缓存值',
          '图纸数量与复核数量分开；缺失不许编'],
         '记录耗时 39 分钟', C_HEAD),
        ('C', '分阶段协作', '拆七个阶段',
         ['模板解析、表格页分离、识别、标准化、填表、验收',
          '每阶段落盘，只传路径与 handoff 摘要',
          '不把 326 页 PDF 读进主上下文'],
         '记录耗时 178 分钟', C_ORANGE),
    ]
    x = L
    for tag, name, sub, items, cost, color in cards:
        rect(slide, x, 2.08, 3.89, 3.92, C_WHITE)
        rect(slide, x, 2.08, 3.89, 0.09, color)
        put(slide, x + 0.24, 2.28, 0.60, 0.60, tag, size=30, bold=True, color=color)
        put(slide, x + 0.86, 2.30, 2.80, 0.36, name, size=17, bold=True, color=C_TITLE)
        put(slide, x + 0.86, 2.64, 2.80, 0.30, sub, size=11.5, color=C_MUTED)
        yy = 3.12
        for it in items:
            dot(slide, x + 0.30, yy + 0.11, 0.08, color)
            put(slide, x + 0.44, yy, 3.20, 0.62, it, size=12, color=C_BODY, spacing=1.12)
            yy += 0.70
        hline(slide, x + 0.24, 5.48, 3.41)
        put(slide, x + 0.24, 5.56, 3.41, 0.30, cost, size=13, bold=True, color=color)
        x += 4.12

    band(slide, 6.28, '前提',
         '输入逐字节相同（PDF 4e1a80ff…／模板 34509914…）；三组都只跑了一次，耗时非插桩计时',
         kind='cool')
    link_source_index(slide)


# ============================================================ P40 结果指纹
def p2(slide, page):
    shell(
        slide, page,
        '三份 final.xlsx 都交了，完成的工作却不是一回事',
        '先看文件里能数出来的东西：公式、合并单元格、明细行、缓存值、图纸数量取自哪个项目',
        '来源：三组 final.xlsx 逐格核对；模板潭州高架桥（计算稿模板）.xlsx 作为基线',
        '第 40 页 · 用结构指纹代替「看起来像」\n\n'
        '这一页不做评分，只列能在文件里数出来的差别。四列分别是：观察项、A、B、C。\n\n'
        '公式数：A 234，B 1,436，C 11,476；模板本身 1,438。B 少 2 个，是被备注文字盖掉的'
        '两个 =F-G 差异公式。\n'
        '合并单元格：模板 735 个（主梁 315、横隔板 362、挡块 24、预埋件 24）。B 和 C 原样保留；'
        'A 只留下 8 个，四个明细表的合并全部消失——A 是重建表格，不是填模板。\n'
        '明细行：A 主梁只有 44 行（其中 38 行是汇总值），B 413 行（模板原样，没换数据），'
        'C 3,647 行（真去识别了节段图）。\n'
        '缓存值：A、B 打开就有数；C 的 11,476 个公式一个缓存都没有，打开是空白格。\n'
        '图纸数量：A 按新沙 16 个厚度档重填；B 只填模板能容纳的 8 档，2,107.7 t 没地方放；'
        'C 的图纸数量列仍是潭州值 2,989,013.3。\n\n'
        '一句话：三份都叫 final.xlsx，但一份是重做、一份是填空、一份是半成品。')
    rows = [
        ('公式数量', '234 个', '1,436 个（模板 1,438）', '11,476 个'),
        ('合并单元格', '8 个（模板 735）', '735 个，逐一保留', '735 个，逐一保留'),
        ('主梁明细行', '44 行，无一行有长宽', '413 行，全部为模板原数据', '3,647 行，来自节段图识别'),
        ('公式缓存值', '234 个全部有值', '1,436 个全部有值', '0 个，打开为空白'),
        ('图纸数量口径', '新沙 16 个厚度档', '新沙 8 档，2,107.7 t 无行可放', '仍是潭州 2,989,013.3'),
        ('记录耗时', '64 分钟', '39 分钟', '178 分钟'),
        ('实际完成到', '重做一张自洽的表', '填了能填的格子', '跑完流程，未完成交付验收'),
    ]
    table(slide, L, 1.96, W,
          ['观察项', 'A · 简短交办', 'B · 详细约束', 'C · 分阶段协作'],
          rows, [0.205, 0.255, 0.275, 0.265], rh=0.545, fs=12.5)
    band(slide, 6.28, '结论',
         '文件名一样，工作量不一样：先数清楚每份做完了什么，再谈谁做得对')
    link_source_index(slide)


# ============================================================ P41 A 组
def p3(slide, page):
    shell(
        slide, page,
        'A 组：总重只差 146 kg，主梁这一段是把汇总值搬回来的',
        '差异小不等于算得细；把总量拆开看，才知道哪一段是真算的',
        '来源：工作沙箱1/final.xlsx；图纸页20 工程数量表；工作沙箱1/validation_report.md',
        '第 41 页 · 最适合现场互动的一页\n\n'
        '先给结论数字：图纸 6,085,603.5 kg，复核 6,085,457.3 kg，差 −146.21 kg，占万分之 0.02%。'
        '这个数字好得不像 AI 干的。\n\n'
        '拆开看。左边这段是 A 组合计数量的最终行。16 个厚度档里有 11 档差异正好是 0，'
        '包括 32、36、40、44、48 这些大吨位档。为什么能正好 0？因为主梁那 44 行是这么填的：'
        '把图纸按材质×板厚的汇总值直接写进单件重，数量一律填 1，合计＝单件重×数量。'
        '44 行里没有一行填了长和宽——它不是算出来的，是把汇总值搬进来的。'
        '公式没错、格式没错，但它是恒等式，不是复核。这 44 行合计 3,986,007.1 kg，'
        '其中 38 行标着图纸汇总值，占 3,454,136.6 kg（86.7%）；余下 6 行是 Z 向钢、焊缝重、'
        '其他板件，也是直接搬的数。\n\n'
        '那 −146.21 是哪来的？全部来自横隔板。横隔板 31 行是真按长×宽×厚×数量算的，'
        '加两行焊缝与其他钢板，合计 2,030,560.1；按厚度汇总 2,000,044.4 kg，'
        '图纸小计 2,000,190.6 kg，差 −146.2 kg。'
        '也就是说：总重差恰好等于横隔板这一段的差，主梁贡献 0——因为它根本没算。\n\n'
        '公平地说：A 组确实自己算了横隔板，还加了面积 A 列处理加劲板按净面积计重的问题，'
        '这不是抄。但主梁这一段，总量接近不能证明构件级复算。\n\n'
        '埋的雷还有：它把四个明细表的合并单元格全清了（735→8），合计数量重建成 16 档，'
        '虽然跨表引用同步改了、结果自洽，但已经不是在填模板。')
    rect(slide, L, 2.02, 4.52, 4.10, C_WHITE)
    rect(slide, L, 2.02, 4.52, 0.42, C_HEAD)
    put(slide, 0.78, 2.02, 4.10, 0.42, '报表最后一行（合计数量 R23）', size=12.5,
        bold=True, color=C_WHITE)
    items = [('图纸数量', '6,085,603.5', 'kg', C_BODY),
             ('复核数量', '6,085,457.3', 'kg', C_TEAL),
             ('差异', '−146.21', 'kg（−0.002%）', C_ORANGE)]
    yy = 2.58
    for label, val, unit, color in items:
        put(slide, 0.82, yy, 1.30, 0.24, label, size=11.5, color=C_MUTED)
        put(slide, 0.82, yy + 0.26, 3.90, 0.44, val, size=25, bold=True, color=color)
        put(slide, 0.82, yy + 0.70, 3.90, 0.20, unit, size=10.5, color=C_MUTED)
        yy += 0.94
    rect(slide, 0.78, 5.46, 4.02, 0.56, C_WARM)
    put(slide, 0.92, 5.46, 3.80, 0.56,
        '16 档里 11 档差 0：主梁按汇总值录入，不是算出来的', size=11, bold=True,
        color=C_ORANGE, spacing=1.05)

    rect(slide, 5.42, 2.02, 7.30, 0.42, C_HEAD)
    put(slide, 5.62, 2.02, 6.90, 0.42, '主梁数量：44 行，合计 3,986,007.1 kg',
        size=12.5, bold=True, color=C_WHITE)
    bar_y = 2.72
    rect(slide, 5.62, bar_y, 5.20, 0.46, C_ORANGE)
    put(slide, 5.72, bar_y, 3.60, 0.46, '44 行全部直接搬数', size=12.5, bold=True,
        color=C_WHITE)
    put(slide, 10.60, bar_y, 1.85, 0.46, '3,986,007.1', size=12, bold=True,
        color=C_ORANGE, align=PP_ALIGN.RIGHT)
    put(slide, 5.62, bar_y + 0.52, 6.95, 0.54,
        '38 行搬图纸汇总值 3,454,136.6（86.7%），6 行搬 Z 向钢 / 焊缝重 / 其他板件 531,870.5；'
        '没有一行填了长宽', size=11, color=C_MUTED, spacing=1.1)

    hdr_y = 3.86
    rect(slide, 5.42, hdr_y, 7.30, 0.34, C_ZEBRA)
    for x, wd, t in [(5.56, 1.30, '板件'), (6.90, 1.05, '材质'),
                     (8.00, 1.20, '单件重'), (9.25, 0.65, '数量'),
                     (9.95, 1.25, '合计重量'), (11.25, 1.35, '备注')]:
        put(slide, x, hdr_y, wd, 0.34, t, size=10.5, bold=True, color=C_MUTED)
    demo = [('主桁', 'Q420D', '8,737.1', '1', '8,737.1', '图纸汇总值'),
            ('主桁', 'Q355D', '56,689.0', '1', '56,689.0', 'Z 向钢'),
            ('主桁拼接', 'Q355D', '1,352.2', '1', '1,352.2', '图纸汇总值')]
    yy = hdr_y + 0.34
    for i, row in enumerate(demo):
        if i % 2 == 0:
            rect(slide, 5.42, yy, 7.30, 0.40, C_WHITE)
        for x, wd, t in zip([5.56, 6.90, 8.00, 9.25, 9.95, 11.25],
                            [1.30, 1.05, 1.20, 0.65, 1.25, 1.35], row):
            put(slide, x, yy, wd, 0.40, t, size=10.5,
                bold=(t == '图纸汇总值'),
                color=C_ORANGE if t == '图纸汇总值' else C_BODY)
        yy += 0.40
    put(slide, 5.56, yy + 0.04, 7.10, 0.30,
        '公式本身没错：合计＝单件重×数量；错的是单件重这一列的来源',
        size=11, color=C_MUTED)

    band(slide, 6.26, '怎么读',
         '总量差 146 kg 不能证明复算：主梁差 0 是恒等式，真算的只有横隔板那一段')
    link_source_index(slide)


# ============================================================ P42 B 组
def p4(slide, page):
    shell(
        slide, page,
        'B 组：模板一格没动，复核列还是潭州的数据',
        '约束写全了，执行也确实照着约束做；问题出在「保留结构」和「装进新数据」本身冲突',
        '来源：工作沙箱2/final.xlsx 与 validation_report.md；模板档位对照',
        '第 42 页 · 最能对上企业实际的一页\n\n'
        'B 组把约束写得最细：保留 sheet、表头、合并、边框、字体、列宽、数字格式；不插删行；'
        '公式必须保留；图纸数量与复核数量分开。它基本都做到了——合并 735 个一个没动，'
        '列宽逐列还原，公式只少了 2 个。\n\n'
        '但也正因为不插删行，模板只有 10/12/14/16/20/24/28/30/40/50 这十个厚度档，'
        '而新沙图纸实际是 8/10/12/14/16/20/24/28/32/36/40/44/48/60/80/100，只有 8 档重合。'
        'δ8、32、36、44、48、60、80、100 加起来 2,107.7 t，占总量的 34.6%，在表里没有行可以放。'
        'B 组把这个事实写进了备注，没有硬塞——这点做得对。\n\n'
        '真正的问题在复核列。G 列（复核数量）是模板的跨表引用，指向主梁/横隔板明细；'
        '而这两张明细表 B 组按「不编造」原则原样没动，仍是潭州的数据。于是出现这一行：'
        '12 mm 档，图纸 193,459.1（新沙），复核 323,303.3（潭州），差异 +129,844.2。'
        '公式算得很对，但两边不是同一个项目。我逐格核过：G4:G13 与模板完全相同，一格没变。\n\n'
        '还有一个细节值得讲：模板里 梁底挡块 I8 和 附属预埋件 I11 是 =F-G 的差异公式，'
        'B 组把备注写进了这两个格子，公式被覆盖。公式数 1,438→1,436 就是这么少的。\n\n'
        '所以这一页的讲法不是「约束太多导致失败」，而是：这次执行优先保住了模板结构，'
        '却没有完成新项目数据与计算关系的替换。')
    rect(slide, L, 2.02, 5.95, 2.30, C_WHITE)
    rect(slide, L, 2.02, 5.95, 0.40, C_HEAD)
    put(slide, 0.78, 2.02, 5.60, 0.40, '合计数量：12 mm 板厚这一行', size=12.5,
        bold=True, color=C_WHITE)
    cols = [(0.78, 1.95, '图纸数量（新沙）', '193,459.1', C_TEAL),
            (2.85, 1.95, '复核数量（潭州）', '323,303.3', C_ORANGE),
            (4.92, 1.60, '差异', '+129,844.2', C_ORANGE)]
    for x, wd, label, val, color in cols:
        put(slide, x, 2.56, wd, 0.28, label, size=10.5, color=C_MUTED)
        put(slide, x, 2.86, wd, 0.50, val, size=20, bold=True, color=color)
    put(slide, 0.78, 3.44, 5.60, 0.62,
        'G 列是模板的跨表引用：主梁数量!M432 + 横隔板!O139。两张明细表没换数据，'
        '复核列就还是旧项目。', size=11, color=C_BODY, spacing=1.14)
    put(slide, 0.78, 4.04, 5.60, 0.24, '逐格核对：G4:G13 与模板完全相同，一格未变',
        size=10.5, bold=True, color=C_ORANGE)

    rect(slide, L, 4.46, 5.95, 1.72, C_COOL)
    put(slide, 0.80, 4.52, 5.55, 0.26, '模板档位 vs 新沙实际板厚', size=11.5,
        bold=True, color=C_TEAL)
    put(slide, 0.80, 4.80, 5.55, 0.26,
        '模板 10/12/14/16/20/24/28/30/40/50', size=10.5, color=C_BODY)
    put(slide, 0.80, 5.08, 5.55, 0.26,
        '新沙 8/10/12/14/16/20/24/28/32/36/40/44/48/60/80/100', size=10.5, color=C_BODY)
    put(slide, 0.80, 5.44, 5.55, 0.62,
        '仅 8 档重合；2,107.7 t（34.6%）在表里没有行可以放', size=11.5, bold=True,
        color=C_ORANGE, spacing=1.1)

    rect(slide, 6.85, 2.02, 5.88, 0.42, C_HEAD)
    put(slide, 7.05, 2.02, 5.50, 0.42, '另外两个能直接查到的事实', size=12.5,
        bold=True, color=C_WHITE)
    facts = [
        ('1', '2 个公式被备注覆盖',
         '梁底挡块 I8、附属预埋件 I11 原本是 =F-G 的差异公式，被备注文字写进了同一格；'
         '公式总数 1,438 → 1,436。'),
        ('2', '模板自带坏引用仍在',
         '横隔板 D140 在模板里就是 =#REF!，B 组原样保留，重算后显示 #NAME?；'
         '这不是本次引入，但交付前该修。'),
    ]
    yy = 2.60
    for num, title, body in facts:
        dot(slide, 7.14, yy + 0.16, 0.30, C_ORANGE)
        put(slide, 7.00, yy + 0.02, 0.28, 0.28, num, size=11, bold=True,
            color=C_WHITE, align=PP_ALIGN.CENTER)
        put(slide, 7.52, yy, 5.00, 0.30, title, size=13, bold=True, color=C_TITLE)
        put(slide, 7.52, yy + 0.32, 4.98, 0.86, body, size=11, color=C_BODY,
            spacing=1.14)
        yy += 1.30

    rect(slide, 6.85, 5.16, 5.88, 0.94, C_WARM)
    put(slide, 7.02, 5.19, 5.50, 0.28, '结论怎么说才准', size=10.5, bold=True,
        color=C_ORANGE)
    put(slide, 7.02, 5.44, 5.52, 0.62,
        '不是「约束太多导致失败」，而是：这次执行优先保住了模板结构，'
        '却没有完成新项目数据与计算关系的替换。', size=10.5, color=C_TITLE, spacing=1.12)

    band(slide, 6.26, '一句话',
         '公式能算，不代表两边是同一个项目的数据')
    link_source_index(slide)


# ============================================================ P43 C 组
def p5(slide, page):
    shell(
        slide, page,
        'C 组：阶段产物最全，交付文件最不能直接用',
        '流程、目录、handoff 都齐；问题都留在了最后那份文件里',
        '来源：工作沙箱3/out/final.xlsx、final_report.json、validation_report.md、handoff_A–G',
        '第 43 页 · 讲流程价值，也讲流程管不到的地方\n\n'
        'C 组是唯一把 326 页真拆开的：模板解析、表格页分离（326 页里挑出 155 页有表格线）、'
        '识别（296 张表、33,889 格）、标准化（3,977 行）、填表（3,758 行）、验收（判 FAIL）、'
        '返修。七份 handoff 都在，每份只写已完成/产物路径/关键字段/未决问题/下一步。\n\n'
        '它也确实修对了一些东西：验收发现汇总行 SUM 范围没随插行扩展、7 处跨表引用没重定向，'
        '返修后重算错误值归零，合并 735 个全保留，样式基本没坏。\n\n'
        '但交付文件还有三个硬伤，都是能打开、能看、不能用：\n'
        '一，11,476 个公式全部没有缓存值。不是公式错，是保存时没重算——Excel 打开一片空白，'
        '看起来像文件坏了。\n'
        '二，按厚度汇总只认 7 个桶（12/14/16/20/24/30/40）。识别出来的厚度值有 20 种，'
        '其中 0.18、4、6 这类明显是误读，但 8、10、28、32、36、44、48、50、60、80、100 是真的；'
        '落在桶外的 2,897,899 kg 不进汇总，覆盖率只有 68.6%。\n'
        '三，图纸数量没换。U 列只在 16 个块首行有值，合计 2,285,699.9，正好等于模板 16 段小计；'
        '合计数量的 F 列也还是 2,989,013.3。于是重算后 G15=7,040,544，对图纸差 +135.5%——'
        '这个百分比没有业务意义，因为一边是新沙明细、一边是潭州基线。\n\n'
        '还有两件事值得说：横隔板、梁底挡块、附属预埋件三张表完全没填（识别结果里没有对应分类），'
        '但保留着模板旧数；明细里有 142 行被自己标了校验 fail，933 行只有数量、没有长宽，'
        '这些行合计重量按 0 计，仍留在表里。\n\n'
        '所以这一页的落点不是「多 Agent 不行」，而是：阶段验收过了，交付验收还没有。')
    flow = ['模板解析', '表格页分离', '识别', '标准化', '填模板', '验收 FAIL', '返修']
    bw = 1.52
    gap = 0.155
    x = L
    for i, label in enumerate(flow):
        fill = C_ORANGE if i == 5 else (C_TEAL if i == 6 else C_WHITE)
        txt = C_WHITE if i in (5, 6) else C_TITLE
        sh = rect(slide, x, 2.00, bw, 0.52, fill, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.14)
        sh.line.color.rgb = RGBColor.from_string(C_ORANGE if i == 5 else C_LINE)
        sh.line.width = Pt(0.9)
        put(slide, x, 2.00, bw, 0.52, label, size=11, bold=True, color=txt,
            align=PP_ALIGN.CENTER)
        if i < len(flow) - 1:
            put(slide, x + bw, 2.00, gap, 0.52, '›', size=13, color=C_MUTED,
                align=PP_ALIGN.CENTER)
        x += bw + gap
    put(slide, L, 2.58, W, 0.28,
        '7 份 handoff：每份只写已完成 / 产物路径 / 关键字段 / 未决问题 / 下一步',
        size=10.5, color=C_MUTED)

    cards = [
        ('11,476', '个公式，0 个缓存',
         '不是算错，是保存时没重算。Excel 打开为空白，像文件坏了；校验时也读不到任何数。', C_ORANGE),
        ('68.6%', '按厚度汇总的覆盖率',
         'SUMIF 只有 7 个桶；识别出的厚度值有 20 种（0.18、4、6 明显是噪声），'
         '桶外 2,897,899 kg 不进汇总。', C_ORANGE),
        ('+135.5%', '重算后的总重差异',
         '图纸数量仍是潭州 2,989,013.3；G15=7,040,544。这个百分比没有业务意义，只说明两边口径不同。',
         C_ORANGE),
    ]
    x = L
    for big, title, body, color in cards:
        rect(slide, x, 3.00, 3.89, 1.62, C_WHITE)
        rect(slide, x, 3.00, 0.07, 1.62, color)
        put(slide, x + 0.22, 3.08, 3.45, 0.50, big, size=22, bold=True, color=color)
        put(slide, x + 0.22, 3.58, 3.50, 0.26, title, size=11, bold=True, color=C_TITLE)
        put(slide, x + 0.22, 3.86, 3.48, 0.70, body, size=10, color=C_BODY, spacing=1.12)
        x += 4.12

    rows = [
        ('横隔板 / 梁底挡块 / 附属预埋件', '三张表完全未填，仍是模板旧数；识别结果里没有对应分类'),
        ('明细里的校验 fail', '142 行被自己标为校验失败；933 行有数量但缺长宽，仍按 0 计入合计'),
        ('图纸数量列', 'U 列只有 16 个块首行有值，合计 2,285,699.9 = 模板 16 段小计'),
    ]
    yy = 4.80
    for i, (a, b) in enumerate(rows):
        if i % 2 == 0:
            rect(slide, L, yy, W, 0.40, C_ZEBRA)
        put(slide, 0.74, yy, 3.55, 0.40, a, size=11, bold=True, color=C_TITLE)
        put(slide, 4.35, yy, 8.20, 0.40, b, size=11, color=C_BODY)
        yy += 0.40

    band(slide, 6.26, '区别',
         '阶段验收看的是每步有没有产物；交付验收看的是打开能不能用、数对不对')
    link_source_index(slide)


# ============================================================ P44 验收清单
def p6(slide, page):
    shell(
        slide, page,
        '算量底稿怎么验收：六个问题，逐份问一遍',
        '这六个问题都能在一分钟内查到答案，不需要重算整册图纸',
        '来源：本稿三组交付物核对；模板档位与跨表引用检查；公式缓存与合并单元格计数',
        '第 44 页 · 把案例接回方法\n\n'
        '前面三页各栽一个坑，这一页把它们变成一个能带走的检查表。六问，每问都写清怎么查。\n\n'
        '一，数据是不是当前项目：看图纸数量列和备注里的出处，别只看文件名。B 组复核列原封不动，'
        '就是这么发现的。\n'
        '二，复核是不是独立：随机抽 5 行，看单件重是长×宽×厚×7850，还是直接等于图纸汇总值、'
        '数量填 1。A 组 38 行都是后者。\n'
        '三，模板档位能不能装下新数据：把图纸的厚度集和模板的厚度行做差集。B 组差出 2,107.7 t。\n'
        '四，计算链是否完整：数公式数、看合并单元格数、查 #REF!/#NAME?。B 的 1,438→1,436、'
        'C 的 11,476 个空缓存，都是这么查出来的。\n'
        '五，最终文件有没有重算并保存：随便点一个公式格看编辑栏有没有值，或用 LibreOffice 无头重算'
        '另存一份对比。\n'
        '六，没做完的是不是隔离开了：待确认行有没有标记、有没有混进合计。C 组 142 行 fail 仍在表里，'
        '但至少有标记；A 组把待设计核实写进了备注，也算隔离。\n\n'
        '下一轮实验怎么改：统一验收标准，放开实现方法。三组的差别不该在谁更听话，'
        '而该在同样的验收口径下，谁能在更短时间内交出能用的东西。资料缺失、模板不适配、'
        '算不完，三种情况要分开记，不能都算成模型的错。')
    rows = [
        ('1  数据是不是当前项目', '图纸数量列、备注出处、复核列来源', 'B：复核列 10 个档全部等于模板旧值'),
        ('2  复核是不是独立', '抽 5 行看单件重：尺寸公式，还是汇总值＋数量 1', 'A：44 行全无长宽，汇总值＋数量 1'),
        ('3  模板档位装不装得下', '图纸厚度集 − 模板厚度行', 'B：缺 8 档，2,107.7 t 无行可放'),
        ('4  计算链完不完整', '公式数、合并数、#REF!/#NAME?、跨表引用指向', 'B：少 2 个公式；C：汇总只剩 7 桶'),
        ('5  最终文件重算并保存', '点公式格看值；或重算另存后对比', 'C：11,476 个公式 0 缓存'),
        ('6  没做完的有没有隔离', '待确认标记、fail 行是否混入合计', 'C：142 行 fail、933 行缺尺寸仍在表内'),
    ]
    table(slide, L, 1.94, W,
          ['验收问题', '一分钟内怎么查', '这次谁栽在上面'],
          rows, [0.255, 0.40, 0.345], rh=0.545, fs=12)
    band(slide, 6.26, '下一轮',
         '统一验收标准，放开方法；资料缺失、模板不适配、算不完，分开记',
         kind='cool')
    link_source_index(slide)


# ============================================================ 装配
RAIL_ITEMS = ['案例一：十个配置审图纸', '看图还是靠 OCR', '假错误长什么样',
              '案例二：四份修编报告', '总额与旧模板残留', '同一条变更的原因',
              '案例三：三组交办做表', '146 kg 的差额怎么来的', '产物齐全 ≠ 能交付']


def retitle_section(prs):
    n = 0
    for s in prs.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if '两个真实案例' in r.text:
                        r.text = r.text.replace('两个真实案例', '三个真实案例')
                        n += 1

    for sh in prs.slides[2].shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() == '图纸审查（看图与 OCR）、修编报告原文对比':
            sh.text_frame.paragraphs[0].runs[0].text = '图纸审查、修编报告、算量底稿三种交办'

    div = prs.slides[27]
    for sh in div.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith('一张挡块图纸'):
            sh.text_frame.paragraphs[0].runs[0].text = (
                '一张挡块图纸、一份修编报告、一份 326 页图纸的算量底稿。'
                '看 AI 找对了什么、漏了什么、编了什么。')

    rail = []
    for sh in div.shapes:
        if sh.has_text_frame and sh.left is not None and sh.top is not None:
            if sh.left / 914400 > 9.0 and 2.00 <= sh.top / 914400 <= 5.60:
                rail.append(sh)
    for sh in rail:
        sh._element.getparent().remove(sh._element)
    for i, label in enumerate(RAIL_ITEMS):
        box = textbox(div, 9.10, 2.04 + 0.47 * i, 3.80, 0.42)
        p = box.text_frame.paragraphs[0]
        p.line_spacing = 1.1
        r = p.add_run()
        r.text = label
        style_run(r, 14.5, False, 'E6ECF5')
    return n


def renumber(prs):
    """页脚数字 = 物理页序（顺带修掉原稿 19–53 页错位的一号）。"""
    for i, s in enumerate(prs.slides, 1):
        for sh in s.shapes:
            if sh.has_text_frame and sh.left is not None and sh.top is not None:
                if abs(sh.left / 914400 - 11.73) < 0.05 and sh.top / 914400 < 0.6:
                    txt = sh.text_frame.text.strip()
                    runs = sh.text_frame.paragraphs[0].runs
                    if txt.isdigit() and runs:
                        runs[0].text = str(i)


def update_z(prs):
    """Z1 编号表：本地案例一行补上第三次实验。"""
    s = prs.slides[96]
    for sh in s.shapes:
        if sh.has_text_frame and sh.text_frame.text.startswith('图纸、清单、四份报告'):
            sh.text_frame.paragraphs[0].runs[0].text = (
                '图纸、清单、四份报告、三组算量底稿与复核口径')
    notes = s.notes_slide.notes_text_frame
    if 'ep3' not in notes.text:
        notes.text += (
            '\n\n2026-10-18 增补：案例三「做表」\n'
            '输入：AI对比/ep3 做表/源文件/主线G6-2 新沙高架桥.pdf（326 页，sha256 4e1a80ff…）；'
            '潭州高架桥（计算稿模板）.xlsx（sha256 34509914…）。三组沙箱内两份输入的哈希完全一致。\n'
            '产物：工作沙箱1/final.xlsx（A）、工作沙箱2/final.xlsx（B）、工作沙箱3/out/final.xlsx（C）。\n'
            '核对方法：openpyxl 逐格读取公式与缓存值；合并单元格、列宽、公式数与模板对比；'
            '主梁明细逐行按 长×宽×厚×7850×数量 独立复算；C 组用 LibreOffice 无头重算复核。\n'
            '口径：只比较「交付与流程」，未对整册图纸做完整工程量审定；三组各跑一次，耗时为文件时间戳推算。')


def main(src, dst):
    prs = Presentation(src)
    n0 = len(prs.slides)
    blank = next(l for l in prs.slide_layouts if l.name == 'Blank')
    pages = [p1, p2, p3, p4, p5, p6]
    new_idx = []
    for k, fn in enumerate(pages):
        s = prs.slides.add_slide(blank)
        fn(s, 39 + k)
        new_idx.append(len(prs.slides) - 1)

    anchor = 38
    order = list(range(anchor)) + new_idx + list(range(anchor, n0))
    lst = prs.slides._sldIdLst
    items = list(lst)
    assert sorted(order) == list(range(len(items)))
    for el in items:
        lst.remove(el)
    for i in order:
        lst.append(items[i])

    n = retitle_section(prs)
    update_z(prs)
    keep = len(sys.argv) > 3 and sys.argv[3] == 'keepnum'
    if keep:
        # 只给原稿中 ≥39 的页脚 +6，保留原有编号习惯
        for i, s in enumerate(prs.slides, 1):
            if i <= 44:
                continue
            for sh in s.shapes:
                if sh.has_text_frame and sh.left is not None and sh.top is not None:
                    if abs(sh.left / 914400 - 11.73) < 0.05 and sh.top / 914400 < 0.6:
                        txt = sh.text_frame.text.strip()
                        if txt.isdigit() and int(txt) >= 39:
                            sh.text_frame.paragraphs[0].runs[0].text = str(int(txt) + 6)
    else:
        renumber(prs)
    for i in range(38, 44):
        link_source_index(prs.slides[i])
    for i in range(44, min(56, len(prs.slides))):
        link_source_index(prs.slides[i])

    prs.save(dst)
    print('saved', dst, 'slides', len(prs.slides), 'retitled', n)
    for i, s in enumerate(prs.slides, 1):
        if 36 <= i <= 47 or i in (54, 55):
            t = next((sh.text_frame.text.split('\n')[0] for sh in s.shapes
                      if sh.has_text_frame and sh.name == 'TextBox 4'), '')
            num = None
            for sh in s.shapes:
                if sh.has_text_frame and sh.left is not None and sh.top is not None:
                    if abs(sh.left / 914400 - 11.73) < 0.05 and sh.top / 914400 < 0.6:
                        num = sh.text_frame.text.strip()
            print('  %3d [%s] %s' % (i, num, t[:56]))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
