# -*- coding: utf-8 -*-
"""版式库：把常见页型做成可复用函数，页面只提供数据"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ============================================================ 页型 1：表格页
def pg_table(s, ch, pno, ttl, sub, head, rows, colw, size=15, head_size=15,
             aligns=None, colors=None, take=None, src=None, note=None,
             headfill=TEAL, minrow=0.60, first_bold=True, top=None,
             right=None, tw=None, rowh=None, zebra=None):
    """right: (x_rel, dict) 右侧说明块；tw: 表宽"""
    y = head(s, ch, pno, ttl, sub)
    w = tw if tw else (CW if not right else right[0] - ML - 0.40)
    ty = y + 0.16
    table(s, ML, ty, w, [head] + rows, colw, size=size, head_size=head_size,
          aligns=aligns, colors=colors, minrow=minrow, first_bold=first_bold,
          rowh=rowh, zebra=zebra)
    if note:
        text(s, ML, ty + 0.02, 0.1, 0.1, "", 8, FAINT, check=False)
    if right:
        rx, blk = right
        vr(s, rx - 0.26, ty, 3.9, 1.0, LINE)
        yy = ty
        for it in blk:
            t, b = it[0], it[1]
            text(s, rx, yy, X2 - rx, 0.30, t, 14, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2, check=False)
            nl = nlines(b, 12.5, X2 - rx)
            text(s, rx, yy + 0.30, X2 - rx, nl * 12.5 * 1.32 * 1.4 / 72 + 0.05, b, 12.5, GREY,
                 SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
            yy += 0.38 + nl * 12.5 * 1.32 * 1.4 / 72
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else 6.16, tag=take[2] if len(take) > 2 else "要点")
    if src:
        footer(s, src)
    return s


# ============================================================ 页型 2：三卡对照
def pg_cards(s, ch, pno, ttl, sub, cards, take=None, src=None, cols=3, top=1.94,
             ch_h=3.66, gap=0.29):
    """cards: [(name, meta, [points], color, tag)]"""
    head(s, ch, pno, ttl, sub)
    cw = (CW - gap * (cols - 1)) / cols
    for i, (nm, meta, pts, col, tag) in enumerate(cards):
        x = ML + i * (cw + gap)
        card(s, x, top, cw, ch_h, accent=col)
        text(s, x + 0.26, top + 0.20, cw - 0.52, 0.32, nm, 17, NAVY, SANS, True, PP_ALIGN.LEFT, 1.2,
             check=False)
        text(s, x + 0.26, top + 0.56, cw - 0.52, 0.26, meta, 13, col, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        hr(s, x + 0.26, top + 0.92, cw - 0.52, 1.0, LINE)
        bullets(s, x + 0.26, top + 1.06, cw - 0.52, pts, size=13.5, dot=col, gap=0.20)
        if tag:
            text(s, x + 0.26, top + ch_h - 0.44, cw - 0.52, 0.26, tag, 12, GREY2, SANS, False,
                 PP_ALIGN.LEFT, 1.0, check=False)
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else top + ch_h + 0.30,
                 tag=take[2] if len(take) > 2 else "判断")
    if src:
        footer(s, src)
    return s


# ============================================================ 页型 3：左右分栏
def pg_split(s, ch, pno, ttl, sub, left, right, take=None, src=None, lw=6.10, gapx=0.60,
             top=1.94):
    """left/right: dict(kind=..., ...)
       kind='blocks': items=[(label, body)]；'stats': [(big,unit,cap)]；
       kind='table': head, rows, colw；kind='pic': fn, cap, h"""
    y = head(s, ch, pno, ttl, sub)
    ty = top if top else y + 0.16
    lx = ML
    rx = ML + lw + gapx

    def render(x, w, d):
        yy = ty
        k = d.get("kind", "blocks")
        if k == "blocks":
            for it in d["items"]:
                if len(it) == 3:
                    tag, t, b = it
                else:
                    tag, (t, b) = None, it
                if tag:
                    text(s, x, yy, w, 0.24, tag, 11.5, d.get("tagc", ORANGE), SANS, True,
                         PP_ALIGN.LEFT, 1.0, cs=0.8, check=False)
                    yy += 0.26
                text(s, x, yy, w, 0.32, t, d.get("tsize", 15.5), NAVY, SANS, True, PP_ALIGN.LEFT, 1.25,
                     check=False)
                nl = nlines(b, d.get("bsize", 13), w)
                text(s, x, yy + 0.34, w, nl * 13 * 1.32 * 1.45 / 72 + 0.06, b, d.get("bsize", 13),
                     GREY, SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
                yy += 0.34 + nl * 13 * 1.32 * 1.45 / 72 + d.get("gap", 0.30)
        elif k == "stats":
            yy = ty
            for big, unit, cap in d["stats"]:
                stat(s, x, yy, big, unit, cap, d.get("color", TEAL), d.get("bsize", 38), w, 12)
                yy += d.get("bsize", 38) * 1.24 / 72 + 0.34 + nlines(cap, 12, w) * 12 * 1.32 * 1.35 / 72
        elif k == "table":
            table(s, x, yy, w, d["rows"], d["colw"], size=d.get("size", 13.5),
                  head_size=d.get("head_size", 13.5), aligns=d.get("aligns"),
                  colors=d.get("colors"), minrow=d.get("minrow", 0.48),
                  first_bold=d.get("first_bold", True))
        elif k == "pic":
            hh = d.get("h", 2.6)
            pic(s, d["fn"], x, yy, w, hh, mode=d.get("mode", "fit"))
            text(s, x, yy + hh + 0.14, w, 0.26, d.get("cap", ""), 11, GREY2, SANS, False,
                 PP_ALIGN.LEFT, 1.2, check=False)
        return yy

    render(lx, lw, left)
    if right:
        vr(s, rx - 0.30, ty, 3.90, 1.0, LINE)
        render(rx, ML + CW - rx, right)
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else 6.16,
                 tag=take[2] if len(take) > 2 else "要点")
    if src:
        footer(s, src)
    return s


# ============================================================ 页型 4：横向步骤
def pg_steps(s, ch, pno, ttl, sub, steps, take=None, src=None, top=2.06, box_h=1.46,
             tail=None):
    """steps: [(no, name, body)]；tail: [(t,b)] 底部补充两栏"""
    head(s, ch, pno, ttl, sub)
    n = len(steps)
    gap = 0.26
    w = (CW - gap * (n - 1)) / n
    for i, (no, nm, body) in enumerate(steps):
        x = ML + i * (w + gap)
        box(s, x, top, w, box_h, fill=CARD, line=LINE)
        box(s, x, top, w, 0.055, fill=TEAL if i % 2 == 0 else C(0x4A, 0x6C, 0x9E))
        text(s, x + 0.22, top + 0.18, w - 0.44, 0.26, no, 12, TEAL, SANS, True, PP_ALIGN.LEFT,
             1.0, cs=0.8, check=False)
        text(s, x + 0.22, top + 0.46, w - 0.44, 0.30, nm, 15.5, NAVY, SANS, True, PP_ALIGN.LEFT,
             1.2, check=False)
        nl = nlines(body, 12.5, w - 0.44)
        text(s, x + 0.22, top + 0.80, w - 0.44, nl * 12.5 * 1.32 * 1.42 / 72 + 0.05, body, 12.5,
             GREY, SANS, False, PP_ALIGN.LEFT, 1.42, check=False)
        if i < n - 1:
            line(s, x + w + 0.03, top + box_h / 2, x + w + gap - 0.03, top + box_h / 2, LINE2, 1.4,
                 head=True)
    if tail:
        yy = top + box_h + 0.36
        for i, (t, b) in enumerate(tail):
            x = ML + i * (CW / len(tail))
            text(s, x, yy, 0.36, 0.24, "·", 15, ORANGE, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
            text(s, x + 0.22, yy, CW / len(tail) - 0.40, 0.26, t, 14, NAVY, SANS, True,
                 PP_ALIGN.LEFT, 1.2, check=False)
            text(s, x + 0.22, yy + 0.28, CW / len(tail) - 0.40, 0.52, b, 12.5, GREY, SANS, False,
                 PP_ALIGN.LEFT, 1.4)
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else 6.16,
                 tag=take[2] if len(take) > 2 else "要点")
    if src:
        footer(s, src)
    return s


# ============================================================ 页型 5：编号清单（可两栏）
def pg_list(s, ch, pno, ttl, sub, items, take=None, src=None, cols=1, top=1.98,
            no_color=ORANGE, tag_w=None):
    """items: [(no, name, body)]"""
    head(s, ch, pno, ttl, sub)
    n = len(items)
    per = (n + cols - 1) // cols
    cw = (CW - (0.70 if cols > 1 else 0)) / cols
    for i, (no, nm, body) in enumerate(items):
        c = i // per
        r = i % per
        x = ML + c * (cw + 0.70)
        yy = top + r * 1.02
        text(s, x, yy + 0.02, 0.54, 0.28, no, 15, no_color, SANS, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, x + 0.50, yy - 0.02, cw - 0.50, 0.30, nm, 15.5, NAVY, SANS, True, PP_ALIGN.LEFT,
             1.2, check=False)
        nl = nlines(body, 12.5, cw - 0.50)
        text(s, x + 0.50, yy + 0.32, cw - 0.50, nl * 12.5 * 1.32 * 1.4 / 72 + 0.05, body, 12.5,
             GREY, SANS, False, PP_ALIGN.LEFT, 1.4, check=False)
        hr(s, x, yy + 0.86, cw - 0.10, 1.0, LINE)
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else 6.16,
                 tag=take[2] if len(take) > 2 else "要点")
    if src:
        footer(s, src)
    return s


# ============================================================ 页型 6：大数字
def pg_stats(s, ch, pno, ttl, sub, stats, take=None, src=None, top=2.10, bsize=46,
             note=None):
    """stats: [(big, unit, cap)] 3-4 个"""
    head(s, ch, pno, ttl, sub)
    n = len(stats)
    w = CW / n
    for i, (big, unit, cap) in enumerate(stats):
        x = ML + i * w
        if i:
            vr(s, x - 0.22, top, 2.30, 1.0, LINE)
        text(s, x, top, w - 0.44, bsize * 1.22 / 72, big, bsize, TEAL if i % 2 == 0 else ORANGE,
             SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        if unit:
            bw = est_w(big, bsize, True) / 72
            text(s, x + bw + 0.06, top + bsize * 0.56 / 72, 1.6, 0.26, unit, 13, GREY, SANS, False,
                 PP_ALIGN.LEFT, 1.0, check=False)
        nl = nlines(cap, 12.5, w - 0.50)
        text(s, x, top + bsize * 1.16 / 72, w - 0.50, nl * 12.5 * 1.32 * 1.45 / 72 + 0.05, cap,
             12.5, GREY, SANS, False, PP_ALIGN.LEFT, 1.45, check=False)
    if note:
        yy = top + bsize * 1.16 / 72 + 1.10
        hr(s, ML, yy, CW, 1.0, LINE)
        text(s, ML, yy + 0.14, CW, 0.80, note, 13.5, C(0x2B, 0x34, 0x45), SANS, False,
             PP_ALIGN.LEFT, 1.5)
    if take:
        takeaway(s, take[0], y=take[1] if len(take) > 1 else 6.16,
                 tag=take[2] if len(take) > 2 else "要点")
    if src:
        footer(s, src)
    return s
