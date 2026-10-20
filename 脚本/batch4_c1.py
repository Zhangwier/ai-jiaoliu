# -*- coding: utf-8 -*-
"""batch4: C1 改成可执行 + 新增 C1b 照做清单"""
import sys, copy
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
EMU=360000
prs = openprs()
def S(n): return prs.slides[n-1]
def g(sl,n): return find(sl,name=n)

# ---------------- C1 (s80) ----------------
s = S(80)
settext(g(s,"TextBox 4"), "C1  动手之前：四件事，各查一次", size=26)
data = [
 ("TextBox 9","TextBox 10","TextBox 11","地区",
  "先确认：这项服务在你所在地开不开",
  "查法：官网定价页或帮助中心搜 supported countries，看你所在国家在不在列表里。\n"
  "OpenAI、Anthropic 不支持中国大陆；Gemini 大陆只开放单位 Workspace 账号。\n"
  "不合格就别往下走：注册到最后一步也会被拦。"),
 ("TextBox 13","TextBox 14","TextBox 15","网络",
  "再确认：官网和接口能不能稳定打开",
  "查法：官网能打开、登录不转圈、接口能返回结果，三样都过才算通。\n"
  "只开网页不够：多数客户端还要连 API 域名。\n"
  "单位网络有管控的，先问清规定再动手。"),
 ("TextBox 17","TextBox 18","TextBox 19","账号",
  "然后：一个能收验证码的邮箱和手机号",
  "优先用 Google 账号：Gemini 必须用，Claude、ChatGPT 也认。\n"
  "手机号要能收国际短信；VoIP、座机多被拒。\n"
  "别买转售账号：找不回、也随时可能被停用。"),
 ("TextBox 21","TextBox 22","TextBox 23","支付",
  "最后：一张能过的卡，或一个能收的渠道",
  "海外产品多数只收 Visa/Mastercard，且要求账单地址在支持地区。\n"
  "国内卡多半过不了，原因在发卡国，不是额度问题。\n"
  "国内产品用支付宝、微信，可开发票。"),
]
for i,(a,b,c,big,head,body) in enumerate(data):
    settext(g(s,a), big, size=32)
    settext(g(s,b), head, size=15.5)
    settext(g(s,c), body, size=12.5)
    for p in g(s,c).text_frame.paragraphs: p.line_spacing = 1.3
settext(g(s,"TextBox 27"), "顺序别反：地区不过就别折腾账号，支付不过就别先买会员", size=18)

# ---------------- 新增 C1b ----------------
NEW_AFTER = 80
lay = prs.slides[NEW_AFTER-1].slide_layout
s = prs.slides.add_slide(lay)
for sh in list(s.shapes): sh._element.getparent().remove(sh._element)

def tb(x,y,w,h,text,size,color=NAVY,bold=False,align=None,anchor=None,space=1.25,num=None):
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
    return r

def oval(x,y,d,fill=WHITE,line=TEAL):
    r=s.shapes.add_shape(MSO_SHAPE.OVAL,Cm(x),Cm(y),Cm(d),Cm(d))
    r.fill.solid(); r.fill.fore_color.rgb=fill
    r.line.color.rgb=line; r.line.width=Pt(1.0); r.shadow.inherit=False
    return r

tb(1.52,0.81,15.24,0.76,"附录  速查",12,GREY)
tb(29.80,0.81,2.54,0.76,"78",12,RGBColor(0x6B,0x73,0x85),align=PP_ALIGN.RIGHT)
rect(1.52,1.68,1.27,0.15,fill=TEAL)
tb(1.52,1.98,30.82,1.55,"C1b  照着做：从零到跑通第一个任务",26,NAVY,True)
tb(1.52,3.40,30.82,0.85,"七步，做一步验一步；任一步不过就停下，别硬往下走",16.5,GREY)

conn = s.shapes.add_connector(1,Cm(1.52),Cm(17.78),Cm(32.34),Cm(17.78)); conn.line.color.rgb=RGBColor(0xE2,0xDF,0xD6); conn.line.width=Pt(0.75)
tb(1.52,17.93,26.67,0.76,"来源：各产品官方帮助中心与定价页（2026-10-18 查询）；本文为操作顺序建议，不含任何绕过地区或身份限制的方法",8.5,GREY)
tb(27.26,17.93,5.08,0.76,"来源索引 ›",8.5,GREY)

steps = [
 ("01","选一个能用的产品", "先按 C1 的四件事筛一遍：地区、网络、账号、支付，四个都能过再选它。",
  "过不了就别试：换国内产品，成本最低。"),
 ("02","注册，只用一个主账号", "优先 Google 账号登录；一个邮箱只注册一个账号，别开小号。",
  "同一手机号多开会被判风控。"),
 ("03","先不付费，用完免费额度", "把免费额度当成试用：找一个你手上真实的小任务跑一遍。",
  "别用“你好”测试，测不出能力。"),
 ("04","验证它能不能读你的文件", "上传一份真实的 PDF 或表格，让它指出某一页某一行写了什么。",
  "这一步决定它能不能干你的活。"),
 ("05","把任务写成一段说明", "目标、材料、格式、验收标准四句话；写完自己读一遍。",
  "写不清就先别怪它。"),
 ("06","对照结果，算一次返工时间", "它出结果、你核对，记下“做多久、错几处”。",
  "这组数字决定值不值得付费。"),
 ("07","再决定买哪一档", "先用最低档跑一周；不够用再加，别一上来买年付。",
  "年付退不了，月付随时停。"),
]
y = 4.55
RH = 1.78
for i,(no,head,body,note) in enumerate(steps):
    if i % 2 == 0:
        rect(1.52,y,30.82,RH,fill=WHITE)
    else:
        rect(1.52,y,30.82,RH,fill=CREAM)
    oval(1.90,y+0.55,0.66,fill=TEAL,line=TEAL)
    tb(1.90,y+0.55,0.66,0.66,no,12,WHITE,True,align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE,space=1.0)
    tb(2.90,y+0.20,6.60,1.40,head,15,NAVY,True,anchor=MSO_ANCHOR.MIDDLE,space=1.2)
    tb(9.75,y+0.20,14.20,1.40,body,12.5,RGBColor(0x2B,0x34,0x45),anchor=MSO_ANCHOR.MIDDLE,space=1.28)
    tb(24.30,y+0.20,7.70,1.40,note,11.5,ORANGE,anchor=MSO_ANCHOR.MIDDLE,space=1.28)
    y += RH

rect(1.52,16.05,30.82,1.30,fill=MINT)
rect(1.52,16.05,0.23,1.30,fill=TEAL)
tb(2.16,16.05,2.54,1.30,"提醒",14,NAVY,anchor=MSO_ANCHOR.MIDDLE)
tb(4.45,16.05,27.26,1.30,"第 3 步和第 6 步最容易被跳过：不跑真实任务，就测不出它到底行不行",17,NAVY,anchor=MSO_ANCHOR.MIDDLE)
move_to(prs, s, NEW_AFTER)
save(prs); print("ok")
