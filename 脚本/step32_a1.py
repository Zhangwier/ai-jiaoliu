# -*- coding: utf-8 -*-
"""A1/A2：总览页前移 + 神经网络移入附录 F"""

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


def idx(key):
    for i, s in enumerate(prs.slides):
        if key in alltext(s):
            return i
    return -1


def dump(tag):
    print("=== " + tag)
    for i, s in enumerate(prs.slides):
        if i < 22 or 104 <= i < 120:
            t = alltext(s).split("\n")[0][:40]
            print(i + 1, t)


def move_after(slide_obj, key_anchor):
    a = idx(key_anchor)
    m = next(i for i, s in enumerate(prs.slides) if s.slide_id == slide_obj.slide_id)
    after = a - 1 if a > m else a
    move_to(prs, slide_obj, after + 1)
    print("moved: [%d] -> after [%d] (anchor)" % (m + 1, a + 1))


dump("BEFORE")

# A1 总览页 → 01 分隔页之后（分隔页 alltext 含“认识 AI”）
target = prs.slides[idx("四个问题，都能在原理里找到原因")]
move_after(target, "它在一个字一个字地猜")

# A2 神经网络 → 附录 F 的 F2（神经网络怎么学）之后
target = prs.slides[idx("神经网络：很多层")]
move_after(target, "F2  神经网络怎么学")

dump("AFTER")
save(prs)
print("saved")
