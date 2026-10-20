# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR
EMU=360000
prs = openprs(); s = prs.slides[85]
def g(n): return find(s, name=n)

# 副标（代价…）位置下移保障不压正文；正文从 6.45 起
for nm in ("TextBox 12","TextBox 19","TextBox 26"):
    g(nm).top = Cm(5.66); g(nm).height = Cm(0.62)
    for p in g(nm).text_frame.paragraphs: p.line_spacing = 1.1
for nm in ("TextBox 13","TextBox 20","TextBox 27"):
    g(nm).top = Cm(6.40); g(nm).height = Cm(2.95)
    for p in g(nm).text_frame.paragraphs: p.line_spacing = 1.42
    for p in g(nm).text_frame.paragraphs:
        for r in p.runs: r.font.size = Pt(12.5)
for nm in ("Rectangle 14","Rectangle 21","Rectangle 28"):
    g(nm).top = Cm(9.52)
for nm in ("TextBox 15","TextBox 22","TextBox 29"):
    g(nm).top = Cm(9.75); g(nm).height = Cm(2.10)
    for p in g(nm).text_frame.paragraphs:
        p.line_spacing = 1.40
        for r in p.runs: r.font.size = Pt(12)

# 底部要点条
quotes = ["窗口只有一个月：先定闹钟", "省的是钱，付的是每天的一次点击", "省下来的钱，可能赔上的是数据"]
for i, (nm, q) in enumerate(zip(("Rectangle 9","Rectangle 16","Rectangle 23"), quotes)):
    pass
xs = [1.52, 12.00, 22.48]
for i, q in enumerate(quotes):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Cm(xs[i]+0.30), Cm(12.10), Cm(9.26), Cm(1.85))
    r.fill.solid(); r.fill.fore_color.rgb = CREAM; r.line.fill.background(); r.shadow.inherit=False
    b = s.shapes.add_textbox(Cm(xs[i]+0.50), Cm(12.10), Cm(8.86), Cm(1.85))
    tf=b.text_frame; tf.word_wrap=True; tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    tf.margin_left=0;tf.margin_right=0;tf.margin_top=0;tf.margin_bottom=0
    p=tf.paragraphs[0]; p.line_spacing=1.30
    rr=p.add_run(); rr.text = q
    rr.font.size=Pt(12.5); rr.font.bold=True; rr.font.color.rgb=NAVY; rr.font.name="微软雅黑"

for nm in ("Rectangle 9","Rectangle 16","Rectangle 23"):
    g(nm).height = Cm(9.95)   # 底 14.50
for nm in ("Rectangle 30",):
    g(nm).top = Cm(14.72); g(nm).height = Cm(0.80)
g("TextBox 31").top = Cm(14.72); g("TextBox 31").height = Cm(0.80)
for nm in ("Rectangle 32",):
    g(nm).top = Cm(15.68); g(nm).height = Cm(1.20)
g("Rectangle 33").top = Cm(15.68); g("Rectangle 33").height = Cm(1.20)
g("TextBox 34").top = Cm(15.68); g("TextBox 34").height = Cm(1.20)
g("TextBox 35").top = Cm(15.68); g("TextBox 35").height = Cm(1.20)
save(prs); print("ok")
