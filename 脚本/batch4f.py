# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
prs = openprs(); s = prs.slides[79]
def g(n): return find(s, name=n)

def rewrite(nm, lines, size=11.5):
    sh = g(nm); tf = sh.text_frame
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs)[1:]: r._r.getparent().remove(r._r)
    r0 = p0.runs[0]
    r0.text = lines[0]; r0.font.size = Pt(size); r0.font.bold=False
    r0.font.color.rgb = RGBColor(0x2B,0x34,0x45); r0.font.name="微软雅黑"
    p0.line_spacing = 1.32
    for ln in lines[1:]:
        p = tf.add_paragraph(); p.line_spacing = 1.32
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold=False
        r.font.color.rgb = RGBColor(0x2B,0x34,0x45); r.font.name="微软雅黑"

settext(g("TextBox 22"), "最后：一张能过的卡", size=15.5)
settext(g("TextBox 10"), "先确认：所在地开不开", size=15.5)
settext(g("TextBox 14"), "再确认：网络通不通", size=15.5)
settext(g("TextBox 18"), "然后：能收码的邮箱和手机号", size=15.5)

rewrite("TextBox 11", [
 "查法：定价页搜 supported countries。",
 "OpenAI、Anthropic 不支持大陆；Gemini 大陆只开单位账号。",
 "不过就别往下走。",
])
rewrite("TextBox 15", [
 "查法：网页能开、登录不转圈、接口能回结果。",
 "只开网页不够：客户端还要连 API 域名。",
 "单位网络有管控的先问规定。",
])
rewrite("TextBox 19", [
 "优先 Google 账号：Gemini 必须用。",
 "手机号要能收国际短信；VoIP、座机多被拒。",
 "别买转售账号：随时可能被停。",
])
rewrite("TextBox 23", [
 "多数只收 Visa/Mastercard，账单地址须在支持地区。",
 "国内卡多半过不了：卡在发卡国。",
 "国内产品用支付宝、微信，可开票。",
])
for nm in ("TextBox 11","TextBox 15","TextBox 19","TextBox 23"):
    g(nm).height = Cm(2.70)
save(prs); print("ok")
