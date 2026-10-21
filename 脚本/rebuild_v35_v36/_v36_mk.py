# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\_v36")
import ds
from ds import new_deck
import open as mopen
import ch1, ch2, ch3, ch4
import app, app2, app3, app4, app5

prs = new_deck()
for i, f in enumerate([mopen.p01, mopen.p02, mopen.p03, mopen.p04, mopen.p05], 1):
    f(prs, i)
ch1.build(prs, 6)
ch2.build(prs, 18)
ch3.build(prs, 28)
ch4.build(prs, 39)

app.a_index(prs, 54)
app.a1(prs, 55)
app.a2(prs, 56)
app.b_matrix(prs, 57)
app.b_top(prs, 58, 0)
app.b_top(prs, 59, 1)
app2.b_math(prs, 60)
app2.b_degrade(prs, 61)
app3.b_limit(prs, 62)
app2.b_others(prs, 63)
app3.c1(prs, 64)
app3.c2(prs, 65)
app3.c4(prs, 66)
app3.c5(prs, 67)
app3.d1(prs, 68)
app4.e1(prs, 69)
app4.e2(prs, 70)
app4.e3(prs, 71)
app4.e4(prs, 72)
app4.e5(prs, 73)
app5.f1(prs, 74)
app5.f2(prs, 75)
app5.f3(prs, 76)
app5.f4(prs, 77)
app5.f5(prs, 78)
app5.f6(prs, 79)
app5.f7(prs, 80)
app5.f8(prs, 81)
app5.f9(prs, 82)
app5.z1(prs, 83)
app5.z2(prs, 84)

prs.save(r"D:\MG\program\ai交流\_v36\build.pptx")
print("slides", len(prs.slides))
for w in ds.WARN:
    print(w)
