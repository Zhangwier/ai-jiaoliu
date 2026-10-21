# -*- coding: utf-8 -*-
import re
p = "ds.py"
s = open(p, encoding="utf-8").read()

s = s.replace("""def hr(s, x, y, w, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), Pt(th))
    r.fill.solid()""", """def hr(s, x, y, w, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), I(w), Pt(th))
    _nostyle(r)
    r.fill.solid()""")

s = s.replace("""def vr(s, x, y, h, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), Pt(th), I(h))
    r.fill.solid()""", """def vr(s, x, y, h, th=0.75, color=RULE):
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, I(x), I(y), Pt(th), I(h))
    _nostyle(r)
    r.fill.solid()""")

s = s.replace("""    r = s.shapes.add_shape(shp, I(x), I(y), I(w), I(h))
    if radius:""", """    r = s.shapes.add_shape(shp, I(x), I(y), I(w), I(h))
    _nostyle(r)
    if radius:""")

s = s.replace("""    r = s.shapes.add_shape(MSO_SHAPE.OVAL, I(cx-d/2), I(cy-d/2), I(d), I(d))
    if fill is None:""", """    r = s.shapes.add_shape(MSO_SHAPE.OVAL, I(cx-d/2), I(cy-d/2), I(d), I(d))
    _nostyle(r)
    if fill is None:""")

s = s.replace("""    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    c.line.color.rgb = color""", """    c = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, I(x1), I(y1), I(x2), I(y2))
    _nostyle(c)
    c.line.color.rgb = color""")

open(p, "w", encoding="utf-8").write(s)
print("patched", s.count("_nostyle("))
