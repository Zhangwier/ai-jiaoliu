# -*- coding: utf-8 -*-
"""设计系统与排版基元 —— v35
纸白 / 墨黑 / 朱砂 / 松绿。编辑式排版：不用渐变、不用发光、不用装饰圆。
所有长度单位 = 英寸；字号 = pt。坐标原点在左上角。
"""
import os, math
from pptx import Presentation
from pptx.util import Inches as I, Pt, Emu
from pptx.dml.color import RGBColor as C
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn
import lxml.etree as etree

# ---------- 画布 ----------
W, H = 13.3333, 7.5
ML, MR = 0.90, 0.90
CW = W - ML - MR
X2 = W - MR
MID = W / 2

# ---------- 色 ----------
PAPER = C(0xF6,0xF4,0xEE)
CARD  = C(0xFF,0xFF,0xFF)
INK   = C(0x16,0x1C,0x28)
BODY  = C(0x32,0x39,0x46)
MUT   = C(0x71,0x78,0x84)
FAINT = C(0x9B,0xA1,0xAA)
RULE  = C(0xD9,0xD3,0xC6)
RULE2 = C(0xB4,0xAD,0x9E)
RED   = C(0xA9,0x35,0x20)
REDM  = C(0xC9,0x6A,0x4E)
REDL  = C(0xF1,0xE1,0xDA)
GRN   = C(0x1C,0x58,0x51)
GRNM  = C(0x3F,0x7C,0x72)
GRNL  = C(0xDC,0xE8,0xE5)
BLU   = C(0x2A,0x48,0x6B)
BLUL  = C(0xDF,0xE6,0xEE)
AMB   = C(0x93,0x66,0x18)
AMBL  = C(0xF3,0xE9,0xD4)
WHITE = C(0xFF,0xFF,0xFF)
DARK  = C(0x15,0x1B,0x26)
DARK2 = C(0x23,0x2B,0x39)

SANS  = "Microsoft YaHei"
SERIF = "Noto Serif SC"
MONO  = "Consolas"

MEDIA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "_v34", "media")

WARN = []
LINES = 1.31          # YaHei 行高约 1.32em

def new_deck():
    prs = Presentation()
    prs.slide_width  = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs

# ---------- 基元 ----------
def _nostyle(shp):
    """删掉 p:style，避免主题里的 effectRef 在渲染器里带出投影/描边"""
    sp = shp._element
    for st in sp.findall(qn('p:style')):
        sp.remove(st)

def bgrect(s, color, x=0, y=0, w=W, h=H):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), I(h))
    _nostyle(r)
    r.fill.solid(); r.fill.fore_color.rgb = color
    r.line.fill.background(); r.shadow.inherit = False
    return r

def slide(prs, paper=PAPER):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bgrect(s, paper)
    return s

def est_w(t, size, bold=False, cs=0.0):
    w = 0.0
    for ch in t:
        o = ord(ch)
        w += size * (0.58 if o < 128 else 1.0)
        if o < 128 and ch.isupper():
            w += size * 0.10
    if bold: w *= 1.04
    return w + cs * max(0, len(t) - 1)

def nlines(t, size, w, bold=False, cs=0.0):
    per = w * 72.0
    n = 0
    for seg in str(t).split("\n"):
        if not seg:
            n += 1
            continue
        n += max(1, math.ceil((est_w(seg, size, bold, cs) / per) - 0.015))
    return n

def text(s, x, y, w, h, t, size=13, color=BODY, font=SANS, bold=False,
         align=PP_ALIGN.LEFT, ls=1.35, anchor=MSO_ANCHOR.TOP, check=True,
         italic=False, cs=None):
    tb = s.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0; tf.margin_right = 0
    tf.margin_top = 0;  tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(str(t).split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = ls
        r = p.add_run(); r.text = ln
        f = r.font
        f.size = Pt(size); f.bold = bold; f.name = font; f.italic = italic
        f.color.rgb = color
        if cs is not None:
            f._rPr.set('spc', str(int(cs * 100)))
    if check:
        need = nlines(t, size, w, bold, cs or 0) * size * LINES * ls / 72.0
        if need > h * 1.03 + 0.02:
            WARN.append(f"OVF {need:.2f}>{h:.2f} @{x:.2f},{y:.2f} {str(t)[:30]!r}")
    return tb

def rich(s, x, y, w, h, paras, size=13, ls=1.35, anchor=MSO_ANCHOR.TOP, wrap=True):
    """paras: [[(text,color,bold,font?,size?),...], {runs:[...],align:..,ls:..}]"""
    tb = s.shapes.add_textbox(I(x), I(y), I(w), I(h))
    tf = tb.text_frame
    tf.word_wrap = wrap; tf.auto_size = MSO_AUTO_SIZE.NONE
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, para in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(para, dict):
            p.alignment = para.get("align", PP_ALIGN.LEFT)
            p.line_spacing = para.get("ls", ls)
            if para.get("sp"): p.space_before = Pt(para["sp"])
            items = para["runs"]
        else:
            p.line_spacing = ls
            items = para
        for it in items:
            t_, col, bold = it[0], it[1], it[2]
            fnt = it[3] if len(it) > 3 and it[3] else SANS
            sz  = it[4] if len(it) > 4 and it[4] else size
            r = p.add_run(); r.text = t_
            r.font.size = Pt(sz); r.font.bold = bold
            r.font.name = fnt; r.font.color.rgb = col
    return tb

def hr(s, x, y, w, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), Pt(th))
    _nostyle(r)
    r.fill.solid(); r.fill.fore_color.rgb = color
    r.line.fill.background(); r.shadow.inherit = False
    return r

def vr(s, x, y, h, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), Pt(th), I(h))
    _nostyle(r)
    r.fill.solid(); r.fill.fore_color.rgb = color
    r.line.fill.background(); r.shadow.inherit = False
    return r

def box(s, x, y, w, h, fill=None, line=None, lw=0.75, dash=None, radius=None):
    shp = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    r = s.shapes.add_shape(shp, I(x), I(y), I(w), I(h))
    _nostyle(r)
    if radius:
        try: r.adjustments[0] = radius
        except Exception: pass
    if fill is None: r.fill.background()
    else:
        r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None: r.line.fill.background()
    else:
        r.line.color.rgb = line; r.line.width = Pt(lw)
        if dash: r.line.dash_style = dash
    r.shadow.inherit = False
    return r

def circ(s, cx, cy, d, fill, line=None, lw=0.75):
    r = s.shapes.add_shape(MSO_SHAPE.OVAL, I(cx-d/2), I(cy-d/2), I(d), I(d))
    _nostyle(r)
    if fill is None: r.fill.background()
    else:
        r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None: r.line.fill.background()
    else: r.line.color.rgb = line; r.line.width = Pt(lw)
    r.shadow.inherit = False
    return r

def line(s, x1, y1, x2, y2, color=RULE2, lw=1.0, head=False, dash=None):
    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    _nostyle(c)
    c.line.color.rgb = color; c.line.width = Pt(lw)
    if dash: c.line.dash_style = dash
    if head:
        ln = c.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn('a:tailEnd'))
        tail.set('type', 'triangle'); tail.set('w', 'med'); tail.set('len', 'med')
    return c

def frame(s, x, y, w, h, color=RULE, lw=0.75, fill=None):
    r = box(s, x, y, w, h, fill=fill, line=color, lw=lw)
    return r


def pic(s, fn, x, y, w=None, h=None, mode="fit", path=None, shadow=False, border=None):
    """mode: fit = 等比装进(w,h) 居中；fill = 等比铺满(w,h) 并裁切"""
    from PIL import Image
    p = path or os.path.join(MEDIA, fn)
    iw, ih = Image.open(p).size
    ar = iw / ih
    if w is None: w = h * ar
    if h is None: h = w / ar
    if mode == "fit":
        if w / h > ar:
            w2 = h * ar
            sh = s.shapes.add_picture(p, I(x + (w - w2)/2), I(y), I(w2), I(h))
        else:
            h2 = w / ar
            sh = s.shapes.add_picture(p, I(x), I(y + (h - h2)/2), I(w), I(h2))
    else:
        sh = s.shapes.add_picture(p, I(x), I(y), I(w), I(h))
        if w / h > ar:                       # 图更宽 -> 裁左右
            f = 1 - (w / (h * ar))
            if f > 0:
                sh.crop_left = f/2; sh.crop_right = f/2
        else:                                # 图更高 -> 裁上下
            f = 1 - (h * ar / w)
            if f > 0:
                sh.crop_top = f/2; sh.crop_bottom = f/2
    return sh

# ---------- 页面构件 ----------
def header(s, chapter, pno, dark=False, rule=True):
    col = C(0x8D,0x96,0xA4) if dark else MUT
    if chapter:
        text(s, ML, 0.375, 8.2, 0.22, chapter, 9, col, SANS, False, PP_ALIGN.LEFT, 1.0, cs=1.5, check=False)
    if pno is not None:
        text(s, X2-1.0, 0.375, 1.0, 0.22, f"{pno:02d}", 9, col, MONO, False, PP_ALIGN.RIGHT, 1.0, check=False)
    if rule:
        hr(s, ML, 0.612, CW, 0.75, C(0x38,0x40,0x4E) if dark else RULE)

def head(s, chapter, pno, ttl, sub=None, size=30, dark=False, w=None, sub_w=None, gap=0.17):
    header(s, chapter, pno, dark)
    ww = w if w else CW
    sz = size
    while sz > 17 and est_w(ttl, sz) > ww * 72 * 0.99:
        sz -= 1
    tcol = WHITE if dark else INK
    nl = nlines(ttl, sz, ww)
    th = nl * sz * LINES * 1.12 / 72.0
    text(s, ML, 0.78, ww, th + 0.05, ttl, sz, tcol, SERIF, False, PP_ALIGN.LEFT, 1.12, check=False)
    y = 0.78 + th + gap
    if sub:
        sw = sub_w if sub_w else ww
        sn = nlines(sub, 13.5, sw)
        sh = sn * 13.5 * LINES * 1.45 / 72.0
        text(s, ML, y, sw, sh + 0.04, sub, 13.5,
             C(0xB2,0xBB,0xC7) if dark else MUT, SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
        y += sh + 0.04
    return y

def footer(s, src, dark=False, y=6.94):
    hr(s, ML, y, CW, 0.75, C(0x38,0x40,0x4E) if dark else RULE)
    text(s, ML, y + 0.105, CW - 0.1, 0.2, src, 8.2,
         C(0x8D,0x96,0xA4) if dark else FAINT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)

def chip(s, x, y, t, color=RED, fill=REDL, size=9, padx=0.09, h=0.24, w=None, bold=True, cs=0.0):
    ww = w if w else (est_w(t, size, bold, cs)/72.0 + padx*2 + 0.02)
    box(s, x, y, ww, h, fill=fill)
    text(s, x, y, ww, h, t, size, color, SANS, bold,
         PP_ALIGN.CENTER, 1.0, anchor=MSO_ANCHOR.MIDDLE, check=False, cs=cs)
    return ww

def stat(s, x, y, big, unit, cap, color=INK, big_size=40, w=3.0, cap_size=10,
         cap_color=MUT, cap_ls=1.35, mono=True, bold=True):
    bw = est_w(big, big_size, bold) / 72.0
    text(s, x, y, w, big_size * 1.26 / 72.0, big, big_size, color, MONO if mono else SANS,
         bold, PP_ALIGN.LEFT, 1.0, check=False)
    if unit:
        text(s, x + bw + 0.05, y + big_size * 0.56 / 72.0, 1.6, 0.22, unit, 10, cap_color,
             SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    cy = y + big_size * 1.08 / 72.0
    cn = nlines(cap, cap_size, w)
    text(s, x, cy, w, cn * cap_size * LINES * cap_ls / 72.0 + 0.04, cap, cap_size, cap_color,
         SANS, False, PP_ALIGN.LEFT, cap_ls, check=False)

def bullets(s, x, y, w, items, size=12.5, gap=None, color=BODY, dot=RED, d=0.05, ls=1.35, pad=0.19):
    yy = y
    for it in items:
        circ(s, x + d/2, yy + size*1.18*0.5/72.0, d, dot)
        nl = nlines(it, size, w - pad)
        hh = nl * size * LINES * ls / 72.0
        text(s, x + pad, yy, w - pad, hh + 0.04, it, size, color, SANS, False, PP_ALIGN.LEFT, ls, check=False)
        yy += hh + (gap if gap is not None else 0.13)
    return yy

def takeaway(s, t, y=6.18, color=INK, sub=None, w=None, rule=True, dark=False):
    """页脚上方的结论带：细线 + 一句判断"""
    ww = w if w else CW
    if rule:
        hr(s, ML, y - 0.12, ww, 0.75, C(0x38,0x40,0x4E) if dark else RULE)
    n = nlines(t, 13, ww, True)
    hh = n * 13 * LINES * 1.45 / 72.0
    text(s, ML, y, ww, hh + 0.06, t, 13, color, SANS, True, PP_ALIGN.LEFT, 1.45, check=False)
    return y + hh

# ---------- 表格（自动行高） ----------
def _no_border(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL','a:lnR','a:lnT','a:lnB'):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
        ln = etree.SubElement(tcPr, qn(tag))
        ln.set('w','0'); ln.set('cap','flat'); ln.set('cmpd','sng'); ln.set('algn','ctr')
        etree.SubElement(ln, qn('a:noFill'))

def table(s, x, y, w, rows, colw, size=10.5, head_size=None, head=True, headfill=None,
          aligns=None, rowlines=True, headrule=True, first_bold=False, colors=None,
          line_color=RULE, valign=MSO_ANCHOR.MIDDLE, pad=0.09, minrow=0.30,
          head_color=INK, body_color=BODY, zebra=None, rowh=None, ls=1.24):
    nr, nc = len(rows), len(rows[0])
    total = float(sum(colw))
    ws = [w * c / total for c in colw]
    hs = []
    for ri, row in enumerate(rows):
        sz = (head_size or size) if (head and ri == 0) else size
        mx = 1
        for ci, v in enumerate(row):
            txt = v if isinstance(v, str) else "".join(u[0] for u in v)
            mx = max(mx, nlines(txt, sz, ws[ci] - pad*1.5, first_bold and ci == 0))
        hh = mx * sz * LINES * ls / 72.0 + 0.11
        if head and ri == 0:
            hh = max(hh, 0.29)
        hs.append(max(minrow, hh))
    if rowh:
        hs = [max(rowh, h) for h in hs]
    tot = sum(hs)
    shp = s.shapes.add_table(nr, nc, I(x), I(y), I(w), I(tot))
    tbl = shp.table
    tblPr = tbl._tbl.tblPr
    for a in ('firstRow','bandRow','bandCol','firstCol','lastRow','lastCol'):
        tblPr.set(a, '0')
    _NoStyle = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    sid = tblPr.find(qn('a:tableStyleId'))
    if sid is None:
        sid = etree.SubElement(tblPr, qn('a:tableStyleId'))
    sid.text = _NoStyle
    for i in range(nc):
        tbl.columns[i].width = Emu(int(I(ws[i])))
    for ri in range(nr):
        tbl.rows[ri].height = Emu(int(I(hs[ri])))
    for ri, row in enumerate(rows):
        for ci, v in enumerate(row):
            cell = tbl.cell(ri, ci)
            cell.margin_left = I(pad); cell.margin_right = I(pad*0.55)
            cell.margin_top = I(0.02); cell.margin_bottom = I(0.02)
            cell.vertical_anchor = valign
            _no_border(cell)
            cell.fill.solid()
            cell.fill.fore_color.rgb = CARD
            if head and ri == 0 and headfill is not None:
                cell.fill.fore_color.rgb = headfill
            elif zebra is not None and ri % 2 == 0:
                cell.fill.fore_color.rgb = zebra
            tf = cell.text_frame
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.NONE
            p = tf.paragraphs[0]
            p.alignment = aligns[ci] if aligns else PP_ALIGN.LEFT
            p.line_spacing = ls
            sz = (head_size or size) if (head and ri == 0) else size
            col = head_color if (head and ri == 0) else body_color
            if colors and ri > 0 and ci in colors:
                col = colors[ci]
            if isinstance(v, tuple):
                for run in v:
                    r = p.add_run(); r.text = run[0]
                    r.font.size = Pt(run[2] if len(run) > 2 else sz)
                    r.font.name = SANS
                    r.font.bold = (head and ri == 0) or (first_bold and ci == 0) or (len(run) > 3 and run[3])
                    r.font.color.rgb = run[1] if len(run) > 1 else col
            else:
                r = p.add_run(); r.text = str(v)
                r.font.size = Pt(sz)
                r.font.name = SANS
                r.font.bold = (head and ri == 0) or (first_bold and ci == 0)
                r.font.color.rgb = col
    if rowlines:
        yy = y
        for ri in range(nr-1):
            yy += hs[ri]
            hr(s, x, yy, w, 0.60, line_color)
    if head and headrule:
        hr(s, x, y + hs[0], w, 1.1, RULE2)
    return y + tot
