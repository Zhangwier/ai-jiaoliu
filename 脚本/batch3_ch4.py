# -*- coding: utf-8 -*-
"""batch3: 第四章改三类场景"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
EMU=360000
prs = openprs()

def tb(sl,x,y,w,h,text,size,color=NAVY,bold=False,align=None,anchor=None,space=1.2,name=None):
    b = sl.shapes.add_textbox(Cm(x),Cm(y),Cm(w),Cm(h))
    tf=b.text_frame; tf.word_wrap=True
    if anchor is not None: tf.vertical_anchor=anchor
    tf.margin_left=0; tf.margin_right=0; tf.margin_top=0; tf.margin_bottom=0
    for i,ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i==0 else tf.add_paragraph()
        if align is not None: p.alignment=align
        p.line_spacing=space
        r=p.add_run(); r.text=ln
        r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color; r.font.name="微软雅黑"
    if name: b.name=name
    return b

def rect(sl,x,y,w,h,fill=None,line=None,name=None):
    r=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,Cm(x),Cm(y),Cm(w),Cm(h))
    if fill is None: r.fill.background()
    else: r.fill.solid(); r.fill.fore_color.rgb=fill
    if line is None: r.line.fill.background()
    else: r.line.color.rgb=line; r.line.width=Pt(0.75)
    r.shadow.inherit=False
    if name: r.name=name
    return r

# ---------- s40 重做 ----------
s = prs.slides[39]
for sh in list(s.shapes):
    # 保留页眉/页脚/标题/副标
    if not sh.has_text_frame: 
        if sh.top/EMU > 4.3: sh._element.getparent().remove(sh._element)
        continue
    t = sh.top/EMU
    if t > 4.3: sh._element.getparent().remove(sh._element)

settext(find(s,name="TextBox 7"), "三类活，按顺序判断", size=27)
settext(find(s,name="TextBox 8"), "不是“AI 能不能干”，是“这件事属于哪一类”：第一类最容易见效，第三类最需要人盯", size=16.5)

cards = [
 ("01  打破信息茧房", TEAL,
  "你不知道有这个东西，所以不知道可以要",
  "找工具、找产品、找公式、找规范、找先例、找一条以前没人帮你查的知识。",
  "典型：一百多个文件里定位一份图纸；某个专业问题该查哪本规范；Excel 里这个效果用什么函数。",
  "它的价值是“省掉搜索和试错”，不是“替你判断”。",
  "先问它“有哪些路子”，再让它把路子做成表"),
 ("02  量大、门槛低、目标明确", NAVY,
  "一堆低级活，人做会烦、会漏",
  "批量图片转表格、批量改名归档、批量抽取字段、按固定模板灌数。",
  "关键前提：不涉及复核——转完就是结果，不需要你再判断对错。",
  "这类它比人稳：会自己反复检查格式和条数。",
  "一次跑一百个，先抽 3 个验格式"),
 ("03  综合任务", ORANGE,
  "要判断、要取舍、要跨材料对照",
  "写修编报告、审图纸、合同条文改写、多步工作流、要调用工具的长任务。",
  "这类最容易出彩，也最容易出事：幻觉、迎合、上下文腐烂全在这里出现。",
  "必须留人复核：给依据、给中间结果、给可回退的版本。",
  "拆成小步，每步留一个能查的结果"),
]
x = 1.52; W = 9.86; GAP = 0.62
for i,(tag,color,head,body,note,foot,howto) in enumerate(cards):
    xx = x + i*(W+GAP)
    rect(s, xx, 4.55, W, 10.45, fill=WHITE)
    rect(s, xx, 4.55, W, 1.05, fill=color)
    tb(s, xx+0.28, 4.55, W-0.56, 1.05, tag, 15, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
    tb(s, xx+0.28, 5.75, W-0.56, 0.85, head, 15.5, NAVY, True, space=1.15)
    tb(s, xx+0.28, 6.62, W-0.56, 1.35, body, 12.5, NAVY, space=1.22)
    tb(s, xx+0.28, 7.98, W-0.56, 1.55, "典型：" + note, 12, GREY, space=1.22)
    tb(s, xx+0.28, 9.55, W-0.56, 1.60, "注意：" + foot, 12, ORANGE, space=1.22)
    rect(s, xx+0.28, 11.45, W-0.56, 1.05, fill=CREAM)
    tb(s, xx+0.45, 11.45, W-0.90, 1.05, "做法：" + howto, 11.5, GREY, anchor=MSO_ANCHOR.MIDDLE, space=1.1)

# 一句过渡
tb(s, 1.52, 15.15, 30.81, 0.72, "顺序很重要：先用第一类建立信任，再把第二类交给它跑量，最后才把第三类整段交出去。", 14, TEAL, True)
rect(s, 1.52, 16.05, 30.81, 1.30, fill=PEACH)
rect(s, 1.52, 16.05, 0.23, 1.30, fill=ORANGE)
tb(s, 2.16, 16.05, 2.54, 1.30, "结论", 14, NAVY, anchor=MSO_ANCHOR.MIDDLE)
tb(s, 4.45, 16.05, 27.26, 1.30, "先按类分，再按类给不同的检查强度：第三类永远要人签字", 18, NAVY, anchor=MSO_ANCHOR.MIDDLE)

# ---------- s41 找图纸：标注为第一类例子 ----------
s = prs.slides[40]
settext(find(s,name="TextBox 7"), "第一类的例子：一百多个文件里，找一张图", size=26)
settext(find(s,name="TextBox 8"), "这就是“打破信息茧房”落到地上：不是它比你聪明，是它肯把一百多个文件翻一遍", size=16)
save(prs); print("ok")
