# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU = 360000
prs = openprs()
s = prs.slides[6]
def g(name): return find(s, name=name)
trows = {}
for sh in s.shapes:
    if not sh.has_text_frame: continue
    t = round(sh.top/EMU, 2)
    if abs(t - 13.82) < 0.2: trows[round(sh.left/EMU,1)] = sh
extra = [trows[k] for k in sorted(trows)]
print([e.text_frame.text[:12] for e in extra])
rows = [
 ("抽卡","同一句话问两次，结果时好时坏","按概率抽 token","多跑几次，关键结论复核",17.0,15.5,15.5,15.5),
 ("幻觉","编出不存在的条款、判例、数字","只挑“最像”的下一个字","要出处，回原文核",17.0,15.5,15.5,15.5),
 ("迎合","你说有错它就找错；你给的前提它照用","偏好训练奖励“让人满意”","开放提问，先让它复述口径",17.0,15.5,15.5,15.5),
 ("上下文腐烂","聊久了忘要求、漏信息","注意力被更多内容摊薄","新开对话，要求放开头",17.0,15.5,15.5,15.5),
 ("奖励黑客","绕规则、改测试、走捷径；护住自己、多要权限","按结果给奖励，只对指标负责","写清红线，过程也要查",17.0,15.0,15.0,15.0),
]
names = ["TextBox 12","TextBox 18","TextBox 23","TextBox 29", None]
cells = [["TextBox 13","TextBox 19","TextBox 24","TextBox 30"],["TextBox 14","TextBox 20","TextBox 25","TextBox 31"],
         ["TextBox 15","TextBox 21","TextBox 26","TextBox 32"]]
for i,(a,b,c,d,sa,sb,sc,sd) in enumerate(rows):
    nm = g(names[i]) if names[i] else extra[0]
    c1 = g(cells[0][i]) if i<4 else extra[1]
    c2 = g(cells[1][i]) if i<4 else extra[2]
    c3 = g(cells[2][i]) if i<4 else extra[3]
    settext(nm, a, size=sa); settext(c1,b,size=sb); settext(c2,c,size=sc); settext(c3,d,size=sd)
save(prs); print("ok")
