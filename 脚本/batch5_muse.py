# -*- coding: utf-8 -*-
"""新增 B5b Muse 专页（插在 s68 Meta 之后）"""
import sys, os
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

prs = openprs()
NEW_AFTER = 68
lay = prs.slides[NEW_AFTER-1].slide_layout
s = prs.slides.add_slide(lay)
for sh in list(s.shapes): sh._element.getparent().remove(sh._element)

def tb(x,y,w,h,text,size,color=NAVY,bold=False,align=None,anchor=None,space=1.28):
    b=s.shapes.add_textbox(Cm(x),Cm(y),Cm(w),Cm(h))
    tf=b.text_frame; tf.word_wrap=True
    if anchor is not None: tf.vertical_anchor=anchor
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    for i,ln in enumerate(text.split("\n")):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        if align is not None: p.alignment=align
        p.line_spacing=space
        r=p.add_run(); r.text=ln
        r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name="微软雅黑"
    return b

def rect(x,y,w,h,fill=None,line=None,name=None):
    r=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(x),Cm(y),Cm(w),Cm(h))
    if fill is None: r.fill.background()
    else: r.fill.solid(); r.fill.fore_color.rgb=fill
    if line is None: r.line.fill.background()
    else: r.line.color.rgb=line; r.line.width=Pt(0.75)
    r.shadow.inherit=False
    if name: r.name=name
    return r

tb(1.52,0.81,15.24,0.76,"附录  速查",12,GREY)
tb(29.80,0.81,2.54,0.76,"66",12,RGBColor(0x6B,0x73,0x85),align=PP_ALIGN.RIGHT)
rect(1.52,1.68,1.27,0.15,fill=TEAL)
tb(1.52,1.98,30.81,1.55,"B5b  Muse：把“记得住你”做成产品",26,NAVY,True)
tb(1.52,3.55,30.81,0.95,"Meta 的第一款前沿模型，主打不是写代码，是当你的长期私人助理",16.5,GREY)
conn=s.shapes.add_connector(1,Cm(1.52),Cm(17.78),Cm(32.33),Cm(17.78))
conn.line.color.rgb=RGBColor(0xE2,0xDF,0xD6); conn.line.width=Pt(0.75)
tb(1.52,17.93,26.67,0.76,"来源：research.meta.ai/blog/muse-wonder；muse.ai；about.fb.com；artificialanalysis.ai（2026-10-18 查询）",8.5,GREY)
tb(27.26,17.93,5.08,0.76,"来源索引 ›",8.5,GREY)

# 右上 logo 徽标
badge_p = r"D:\MG\program\ai交流\成品\_work29\logos3\meta.png"
if not os.path.exists(badge_p):
    badge_p = r"D:\MG\program\ai交流\成品\_work29\logos2\Meta2.png"
ov = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Cm(27.23), Cm(0.58), Cm(1.88), Cm(1.88))
ov.fill.solid(); ov.fill.fore_color.rgb=WHITE
ov.line.color.rgb=RGBColor(0xE2,0xDF,0xD6); ov.line.width=Pt(0.75); ov.shadow.inherit=False
try: ov.adjustments[0]=0.22
except: pass
s.shapes.add_picture(badge_p, Cm(27.38), Cm(0.71), Cm(1.57), Cm(1.57))

# 左：形象
rect(1.52,4.83,7.01,11.28,fill=WHITE,line=RGBColor(0xE8,0xE5,0xDC))
s.shapes.add_picture(r"D:\MG\program\ai交流\素材\AI娘\MUSE.jpg", Cm(2.01), Cm(5.05), Cm(6.04), Cm(10.82))

# 右：五行
rows = [
 ("它是什么", "Meta Superintelligence Labs 的第一款前沿模型（Muse Wonder / Wonder Full，2026-10-12）。AA 智能指数 52；LMArena 文字榜一度第二。前身 Muse Spark（2026-04）现在仍在支撑 Meta AI。"),
 ("两个招牌能力", "Active Memory：跨对话记住你的偏好、行程、家人信息，下次直接接着用。\nActive Focus：长任务不跑偏——订多程行程、反复改方案时不丢前面的条件。"),
 ("入口与形态", "Muse App（muse.ai，网页＋手机）；Muse Code（订阅者的编程入口）；Meta Model API；Facebook / Instagram / WhatsApp 里的 Meta AI。"),
 ("短板与边界", "只对美国、加拿大开放（第三方核查口径）；免费档有用量上限，Power $20、Maximum $100；\n企业级权限与审计弱于 OpenAI、Anthropic；编程不是它的主场。"),
 ("什么时候值得试", "要一个“记得住你”的私人助理：行程、比价、定制、排座位、反复修改的方案。\n要交付工程文件、要接公司系统，还是用别的。"),
]
y = 4.83
RH = 2.26
for i,(k,v) in enumerate(rows):
    if i % 2 == 0:
        rect(9.19,y,23.14,RH,fill=WHITE) if i>0 else None
    tb(9.19,y+0.05,2.79,1.30,k,11.5,TEAL,True)
    tb(12.14,y,20.19,RH-0.10,v,11.5,NAVY)
    if i>0:
        c=s.shapes.add_connector(1,Cm(9.19),Cm(y),Cm(32.33),Cm(y))
        c.line.color.rgb=RGBColor(0xE2,0xDF,0xD6); c.line.width=Pt(0.75)
    y += RH

rect(1.52,16.13,30.81,1.27,fill=MINT)
rect(1.52,16.13,0.23,1.27,fill=TEAL)
tb(2.16,16.13,2.67,1.27,"记住它",14,NAVY,anchor=MSO_ANCHOR.MIDDLE)
tb(4.52,16.13,27.18,1.27,"“记得住你”是它唯一真正的差异点；别拿它去比写代码和跑长任务",17,NAVY,anchor=MSO_ANCHOR.MIDDLE)
move_to(prs, s, NEW_AFTER)
save(prs); print("ok")
