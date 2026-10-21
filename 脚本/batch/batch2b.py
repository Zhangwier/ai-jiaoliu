# -*- coding: utf-8 -*-
"""收紧第二章新页（s20）"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
EMU=360000
prs = openprs(); s = prs.slides[19]

# 卡片 4.72 → 高 9.55；脚注条移到 13.15
for sh in s.shapes:
    t = sh.top/EMU; h = sh.height/EMU
    if abs(t-4.72) < 0.05 and h > 9.5:
        sh.height = Cm(9.55)
    if abs(t-14.10) < 0.05:
        sh.top = Cm(13.02)
    if abs(t-14.10) < 0.06 and sh.has_text_frame:
        sh.top = Cm(13.02)

# 正文加大
bodies = [sh for sh in s.shapes if sh.has_text_frame and abs(sh.top/EMU-7.05) < 0.06]
for b in bodies:
    for p in b.text_frame.paragraphs:
        for r in p.runs:
            r.font.size = Pt(13.5)

extras = [
 "\n\n自测：换个贵一档的模型再跑同一条，若明显变好，就是档位问题。",
 "\n\n自测：把同一个任务分别给两家，看谁更少返工。",
 "\n\n自测：换一个能读原文件的客户端，任务不变再跑一次。",
]
for b, ex in zip(bodies, extras):
    tf = b.text_frame
    p = tf.add_paragraph()
    r = p.add_run(); r.text = ex.strip()
    r.font.size = Pt(12); r.font.bold = True; r.font.color.rgb = TEAL; r.font.name="微软雅黑"
    p.line_spacing = 1.2
save(prs); print("ok", len(bodies))
