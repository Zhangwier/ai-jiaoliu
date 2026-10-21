# -*- coding: utf-8 -*-
"""章节分隔页等共用构件"""
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR


def divider(prs, no, name, sub, items=None, quote=None, pno=None, variant="list", img=None):
    s = slide(prs, DARK)
    header(s, "章节", pno, dark=True)
    text(s, ML, 1.62, 2.2, 1.30, no, 96, C(0x2E,0x37,0x47), MONO, True, PP_ALIGN.LEFT, 1.0, check=False)
    text(s, ML + 1.62, 1.90, 6.2, 0.30, "PART", 9.5, C(0x7E,0x88,0x96), SANS, False,
         PP_ALIGN.LEFT, 1.0, cs=2.4, check=False)
    text(s, ML + 1.62, 2.14, 8.4, 0.72, name, 40, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.1, check=False)
    text(s, ML + 1.62, 2.98, 7.6, 0.34, sub, 13, C(0xB2,0xBB,0xC7), SANS, False, PP_ALIGN.LEFT, 1.35)
    hr(s, ML, 3.62, CW, 0.75, C(0x38,0x40,0x4E))

    if variant == "list" and items:
        yy = 3.92
        for i, (k, v) in enumerate(items):
            text(s, ML, yy, 1.5, 0.26, k, 11, C(0x7E,0x88,0x96), MONO, False, PP_ALIGN.LEFT, 1.0, check=False)
            text(s, ML + 1.62, yy - 0.03, 7.2, 0.28, v, 12.5, C(0xDD,0xE2,0xE8), SANS, False, PP_ALIGN.LEFT, 1.2, check=False)
            yy += 0.44
    if variant == "img" and img:
        pic(s, img[0], img[1], img[2], img[3], img[4], mode=img[5] if len(img) > 5 else "fill")
    if quote:
        text(s, ML, 6.30, CW - 1.0, 0.34, quote, 11.5, C(0x8D,0x96,0xA4), SANS, False,
             PP_ALIGN.LEFT, 1.3, check=False)
    return s
