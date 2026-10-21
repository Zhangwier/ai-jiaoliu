# -*- coding: utf-8 -*-
"""样张 v2：给页面加视觉重量 —— 满幅图、深色压层、巨号数字栏"""
import sys, os
sys.path.insert(0, r"D:\MG\program\ai交流\_v35")
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Pt
from pptx.oxml.ns import qn


def scrim(s, x, y, w, h, color=(0x10, 0x13, 0x18), alpha=88):
    """半透明压层"""
    sh = box(s, x, y, w, h, fill=C(*color))
    sf = sh.fill._xPr.find(qn('a:solidFill'))
    clr = sf.find(qn('a:srgbClr'))
    a = clr.makeelement(qn('a:alpha'), {'val': str(int(alpha * 1000))})
    clr.append(a)
    return sh


def cover(prs):
    """封面：整幅长卷铺满，上缘深色压层"""
    s = slide(prs, DARK)
    pic(s, "s001_1.png", 0, 0, W, H, mode="fill")
    scrim(s, 0, 0, W, 3.55, (0x0E, 0x11, 0x16), 90)
    scrim(s, 0, 3.55, W, 3.95, (0x0E, 0x11, 0x16), 42)
    text(s, ML, 0.96, 9.0, 0.28, "工作常用 AI 的使用观察  ·  2026 年 10 月", 10.5, C(0xC8, 0xCF, 0xD9),
         SANS, False, PP_ALIGN.LEFT, 1.0, cs=1.6, check=False)
    text(s, ML, 1.38, 9.4, 1.30, "AI 交流分享", 68, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.06, check=False)
    text(s, ML, 2.72, 8.2, 0.44, "从原理到实践", 20, C(0xDC, 0xE2, 0xEA), SANS, False, PP_ALIGN.LEFT, 1.0,
         check=False)
    hr(s, ML, 4.24, 4.60, 1.2, C(0xE8, 0xDD, 0xCE))
    text(s, ML, 4.44, 7.4, 0.30, "三个真实测试", 12, C(0xF0, 0xE6, 0xD8), SANS, True, PP_ALIGN.LEFT, 1.0,
         cs=2.0, check=False)
    text(s, ML, 4.80, 8.0, 0.66,
         "AI 能做到哪一步 · 哪里会出错 · 哪些地方必须人来核", 15, C(0xD5, 0xDC, 0xE4), SANS, False,
         PP_ALIGN.LEFT, 1.4)
    return s


def photo_split(prs):
    """左深色文字栏 + 右满幅图纸"""
    s = slide(prs, DARK)
    pic(s, "s096_1.jpg", 5.10, 0, W - 5.10, H, mode="fill")
    vr(s, 5.10, 0, H, 0.9, C(0x3A, 0x42, 0x50))
    header(s, "03   两个真实案例", 29, dark=True)
    text(s, ML, 1.30, 3.9, 0.26, "案例一 · 挡块图纸审查", 9.5, C(0x8D, 0x96, 0xA4), SANS, False,
         PP_ALIGN.LEFT, 1.0, cs=1.4, check=False)
    text(s, ML, 1.66, 3.95, 1.92, "一张图纸，\n两个检查点。", 38, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.24,
         check=False)
    hr(s, ML, 3.76, 3.72, 1.0, C(0x4A, 0x54, 0x63))
    items = [("K1", "分段与总长对不上", "1a 大样各段加起来，比图上印的总长公式多 45 cm"),
             ("K2", "符号前后不一致", "3a 图上写 sinα，引线公式写 sinβ")]
    yy = 3.98
    for k, nm, body in items:
        text(s, ML, yy + 0.02, 0.52, 0.24, k, 10.5, C(0xD9, 0x62, 0x2B), MONO, True, PP_ALIGN.LEFT, 1.0,
             check=False)
        text(s, ML + 0.56, yy - 0.04, 3.20, 0.28, nm, 13.5, WHITE, SANS, True, PP_ALIGN.LEFT, 1.15,
             check=False)
        text(s, ML + 0.56, yy + 0.28, 3.14, 0.56, body, 9.4, C(0x9E, 0xA8, 0xB6), SANS, False,
             PP_ALIGN.LEFT, 1.4)
        yy += 0.94
    text(s, ML, 6.62, 3.80, 0.40,
         "图上的斜梯大样：青色阴影处即 K2 所在位置。", 8.6, C(0x7E, 0x88, 0x96), SANS, False,
         PP_ALIGN.LEFT, 1.35, check=False)
    return s


def dark_bar(prs):
    """顶部深色通栏（标题 + 三个巨号数字）+ 下方浅色内容"""
    s = slide(prs)
    bgrect(s, DARK, 0, 0, W, 2.98)
    header(s, "03   两个真实案例", 30, dark=True)
    text(s, ML, 0.80, 9.6, 0.50, "三个配置命中，假错误累计三十条", 27, WHITE, SERIF, False, PP_ALIGN.LEFT,
         1.12, check=False)
    text(s, ML, 1.34, 8.4, 0.28,
         "十个配置逐条复核：命中看的是那两个检查点，不是条目多少。", 11.2, C(0xA8, 0xB2, 0xC0), SANS,
         False, PP_ALIGN.LEFT, 1.35, check=False)
    K = [("3", "个配置命中 K1 或 K2", C(0xE0, 0x8A, 0x6E)),
         ("30", "条假错误（与原图对不上）", C(0xE0, 0x8A, 0x6E)),
         ("21", "条待核未计", C(0x9E, 0xA8, 0xB6))]
    for i, (v, cap, col) in enumerate(K):
        x = ML + i * 3.90
        text(s, x, 1.86, 1.6, 0.72, v, 46, col, MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
        text(s, x + (0.62 if len(v) == 1 else 1.12), 2.10, 2.6, 0.24, cap, 9.2, C(0x9E, 0xA8, 0xB6),
             SANS, False, PP_ALIGN.LEFT, 1.0, check=False)

    rows = [["配置", "耗时", "K1", "K2", "其他", "假错误", "条数"],
            ["Opus 5.5 · Hark", "3:03", "●", "●", "2", "2", "11"],
            ["GPT 6.1 Sol · Codex", "14:34", "●", "○", "4", "1", "8"],
            ["Qwen 3.8 Flash · Qoder", "43:09", "○", "●", "2", "5", "24"],
            ["Auto fast · WorkBuddy", "2:38", "○", "○", "1", "2", "7"],
            ["DeepSeek Flash", "32:05", "○", "○", "1", "2", "19"],
            ["Step 5 Preview", "29:40", "○", "○", "1", "8", "26"]]
    table(s, ML, 3.32, 7.10, rows, [2.48, 0.70, 0.44, 0.44, 0.62, 0.84, 0.58], size=9.6, head_size=8.6,
          aligns=[PP_ALIGN.LEFT, PP_ALIGN.RIGHT, PP_ALIGN.CENTER, PP_ALIGN.CENTER, PP_ALIGN.CENTER,
                  PP_ALIGN.RIGHT, PP_ALIGN.RIGHT], minrow=0.36, first_bold=True,
          line_color=C(0xD5, 0xCF, 0xC0))

    # 右侧：一个真实的对比条
    x = 8.28
    text(s, x, 3.32, 4.10, 0.22, "条目多 ≠ 找得准", 8.8, MUT, SANS, True, PP_ALIGN.LEFT, 1.0, cs=0.8,
         check=False)
    B = [("Opus", 11, 2, "两个检查点全中"), ("Step 5", 26, 8, "26 条里 8 条是假的"),
         ("Qwen", 24, 5, "只有 1 个检查点")]
    yy = 3.62
    for nm, tot, fk, note in B:
        text(s, x, yy, 0.86, 0.22, nm, 9.6, INK, SANS, True, PP_ALIGN.LEFT, 1.0, check=False)
        box(s, x + 0.92, yy + 0.06, (tot - fk) / 26 * 2.34, 0.13, fill=INK)
        box(s, x + 0.92 + (tot - fk) / 26 * 2.34, yy + 0.06, fk / 26 * 2.34, 0.13, fill=RED)
        text(s, x + 3.10, yy, 0.50, 0.22, str(tot), 9.4, INK, MONO, True, PP_ALIGN.RIGHT, 1.0, check=False)
        text(s, x, yy + 0.26, 3.60, 0.20, note, 8.4, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
        yy += 0.62

    footer(s, "来源：本地图纸审查记录（十份输出逐条复核）")
    return s


def grid_of_evidence(prs):
    """图配文：四张真实截图，一张大图 + 三张小图"""
    s = slide(prs)
    header(s, "03   两个真实案例", 34)
    text(s, ML, 0.80, 8.6, 0.46, "四份报告，四条证据", 27, INK, SERIF, False, PP_ALIGN.LEFT, 1.12,
         check=False)
    text(s, ML, 1.32, 9.4, 0.28,
         "不看结论，先看它们各自写下的那一段原文 —— 差别都在措辞里。", 11.2, MUT, SANS, False,
         PP_ALIGN.LEFT, 1.35, check=False)

    frames = [("s034_1.png", ML, 1.92, 6.02, 1.82, "Opus · Hark（4 分 · 43 条）"),
              ("s034_3.png", ML, 4.08, 6.02, 1.82, "WorkBuddy Auto（14 分 · 101 条）"),
              ("s034_4.png", 6.92, 1.92, 5.44, 1.82, "DeepSeek Flash（34 分 · 99 条）"),
              ("s034_2.png", 6.92, 4.08, 5.44, 1.82, "Trae Auto（7 分 · 24 行）")]
    for fn, x, y, w, h, cap in frames:
        box(s, x - 0.10, y - 0.10, w + 0.20, h + 0.20, fill=CARD, line=C(0xD5, 0xCF, 0xC0))
        pic(s, fn, x, y, w, h, mode="fill")
        text(s, x, y + h + 0.16, w, 0.22, cap, 8.6, MUT, SANS, False, PP_ALIGN.LEFT, 1.0, check=False)
    footer(s, "来源：四份修编报告原文（截图为各份对应段落）")
    return s


def main():
    prs = new_deck()
    cover(prs)
    photo_split(prs)
    dark_bar(prs)
    grid_of_evidence(prs)
    prs.save(r"D:\MG\program\ai交流\_v35\probe_v2.pptx")
    print("ok", len(prs.slides))
    for w in WARN:
        print(w)


main()
