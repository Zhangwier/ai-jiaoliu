# -*- coding: utf-8 -*-
"""收尾修正：
1) 第 102 页 "A01–C10" 应为 "A01–A10"（账号打码编号）
2) 第 57 页备注里的“会后查阅”导航仍是旧编号，改成本版结构
"""
import re
from pptx import Presentation

F = r"D:\MG\program\ai交流\AI交流分享-成品26.pptx"
prs = Presentation(F)

# 1) 账号编号
n = 0
for idx in (101, 102):
    s = prs.slides[idx - 1]
    shapes = [sh for sh in s.shapes if sh.has_text_frame]
    if s.has_notes_slide:
        shapes.append(None)
    for sh in shapes:
        tf = sh.text_frame if sh is not None else s.notes_slide.notes_text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                if "C10" in r.text:
                    r.text = r.text.replace("A01–C10", "A01–A10").replace("A01-C10", "A01-A10")
                    n += 1
print("account range fixed:", n)

# 2) 第 57 页会后查阅导航
s57 = prs.slides[56]
NEW_NAV = ("会后查阅  常用术语 A1–A4 · 一页一家公司 B1–B13 · 订阅与省流 C1–C7 · "
           "资讯渠道 D1–D4 · 案例与风险 E1–E4 · 来源 Z1–Z3  ›")
hit = 0
for sh in s57.shapes:
    if not sh.has_text_frame:
        continue
    if sh.text_frame.text.strip().startswith("会后查阅"):
        tf = sh.text_frame
        p = tf.paragraphs[0]
        runs = p.runs
        if runs:
            runs[0].text = NEW_NAV
            for r in runs[1:]:
                r._r.getparent().remove(r._r)
            hit += 1
print("slide57 nav updated:", hit)

prs.save(F)
