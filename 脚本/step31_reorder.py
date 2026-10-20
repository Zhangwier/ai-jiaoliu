# -*- coding: utf-8 -*-
"""阶段 A/B：页序重排（不改文字）
A1 总览页（四问题汇总）前移到 01 分隔页之后
A2 神经网络页 移入附录 F（F2 之后）
B1 C6 降智 移到 B1b（OpenAI 数学）之后
B2 C5 Anthropic 限制 移到 B2b（Opus 生成视频）之后
"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *

prs = openprs()

def dump(tag, ranges=((0, 20), (55, 82), (100, 125))):
    print("=== " + tag)
    for i, s in enumerate(prs.slides):
        if any(a <= i < b for a, b in ranges):
            print(i + 1, title_of(s)[:46].replace("\n", " / "))

def move_after(prs, key_moving, key_anchor):
    """把 key_moving 页移到 key_anchor 页之后"""
    a = index_of(prs, key_anchor)
    m = index_of(prs, key_moving)
    if a < 0 or m < 0:
        print("!! 未找到", key_moving, a, m); return False
    after = a - 1 if a > m else a
    move_to(prs, prs.slides[m], after + 1)
    print("moved  [%d]%s -> after [%d]%s" % (m + 1, key_moving, a + 1, key_anchor))
    return True

dump("BEFORE")

# A1
move_after(prs, "四个问题，都能在原理里找到原因", "认识 AI")
# A2 神经网络进附录 F
move_after(prs, "神经网络：很多层", "F2 神经网络")
# B1 降智 → OpenAI 数学 之后
move_after(prs, "官方承认过 bug", "OpenAI 数学")
# B2 Anthropic 限制 → Opus 生成视频 之后
move_after(prs, "Anthropic 对中国用户", "Opus 生成视频")

dump("AFTER")
save(prs)
print("saved")
