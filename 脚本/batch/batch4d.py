# -*- coding: utf-8 -*-
"""C1 四卡正文重写：统一 12pt 深色正文，两段结构"""
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.text import MSO_ANCHOR
prs = openprs(); s = prs.slides[79]
def g(n): return find(s, name=n)

def rewrite(name, lines, size=12.0):
    sh = g(name)
    tf = sh.text_frame
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs)[1:]:
        r._r.getparent().remove(r._r)
    r0 = p0.runs[0] if p0.runs else p0.add_run()
    r0.text = lines[0]
    r0.font.size = Pt(size); r0.font.bold = False
    r0.font.color.rgb = RGBColor(0x2B,0x34,0x45); r0.font.name = "微软雅黑"
    p0.line_spacing = 1.30
    for ln in lines[1:]:
        p = tf.add_paragraph(); p.line_spacing = 1.30
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = False
        r.font.color.rgb = RGBColor(0x2B,0x34,0x45); r.font.name = "微软雅黑"

rewrite("TextBox 11", [
 "查法：定价页或帮助中心搜 supported countries。",
 "OpenAI、Anthropic 不支持中国大陆；Gemini 大陆只开放单位 Workspace 账号。",
 "不过就别往下走——注册到最后一步也会被拦。",
])
rewrite("TextBox 15", [
 "查法：官网能打开、登录不转圈、接口能返回结果。",
 "只开网页不够：多数客户端还要连 API 域名。",
 "单位网络有管控的，先问清规定。",
])
rewrite("TextBox 19", [
 "优先 Google 账号：Gemini 必须用，Claude、ChatGPT 也认。",
 "手机号要能收国际短信；VoIP、座机多被拒。",
 "别买转售账号：找不回，也随时可能被停用。",
])
rewrite("TextBox 23", [
 "多数海外产品只收 Visa/Mastercard，账单地址还须在支持地区。",
 "国内卡多半过不了：原因在发卡国，不是额度。",
 "国内产品用支付宝、微信，可开发票。",
])
# 四个正文框高度统一 2.55，起点 7.15
for nm in ("TextBox 11","TextBox 15","TextBox 19","TextBox 23"):
    g(nm).top = Cm(7.12); g(nm).height = Cm(2.55)
for nm in ("TextBox 10","TextBox 14","TextBox 18","TextBox 22"):
    g(nm).top = Cm(5.30); g(nm).height = Cm(1.70)
for nm in ("TextBox 9","TextBox 13","TextBox 17","TextBox 21"):
    g(nm).top = Cm(5.30); g(nm).height = Cm(6.80)
    g(nm).text_frame.vertical_anchor = MSO_ANCHOR.TOP
save(prs); print("ok")
