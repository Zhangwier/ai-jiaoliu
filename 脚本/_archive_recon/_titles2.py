# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *

prs = openprs()


def alltext(slide):
    out = []
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text.strip():
            out.append(sh.text_frame.text)
    return "\n".join(out)


def big(slide):
    """最大字号文本"""
    best = (0, "")
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if r.font.size and r.font.size.pt > best[0]:
                    best = (r.font.size.pt, r.text)
    return best


for i, s in enumerate(prs.slides):
    if i >= 119:
        break
    sz, t = big(s)
    tmp = [sh for sh in s.shapes if sh.has_text_frame and 1.5 < sh.top / CM < 4.5 and sh.text_frame.text.strip()]
    head = tmp[0].text_frame.text.split("\n")[0][:6] if tmp else ""
    print("%3d | %-8s | %s" % (i + 1, head, t[:56]))
