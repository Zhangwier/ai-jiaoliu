# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
EMU=360000
prs = openprs(); s = prs.slides[108]

# 删除底部两条说明（14.2 以下、16.0 以上的文本框）
for sh in list(s.shapes):
    if sh.has_text_frame and sh.top/EMU > 14.0 and sh.top/EMU < 16.05:
        sh._element.getparent().remove(sh._element)

def tb(l,t,w,h,lines,size=11,bold=True,color=None):
    box=s.shapes.add_textbox(Cm(l),Cm(t),Cm(w),Cm(h))
    tf=box.text_frame; tf.word_wrap=True
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i,ln in enumerate(lines):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=PP_ALIGN.LEFT; p.line_spacing=1.15
        r=p.add_run(); r.text=ln; r.font.size=Pt(size); r.font.bold=bold; r.font.name="微软雅黑"
        r.font.color.rgb = color or TEAL
    return box

tb(1.52, 14.96, 30.83, 0.60, ["可直接落地的开源权重：DeepSeek V4.1 · Qwen3.8 · GLM-5.3 · Kimi K3 · Hy4 · MiniMax M3（各家边界见附录 B）"], size=11)
tb(1.52, 15.52, 30.83, 0.60, ["国内常见的落地形态：DeepSeek 私有化部署 / GLM 一体机 / Qwen 百炼 + 自建 / Hy 内网单机 / Mistral on-prem"], size=11, color=ORANGE)
save(prs); print("ok")
