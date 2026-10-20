# -*- coding: utf-8 -*-
import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs()

TOC = "A 常用术语\nB 公司档案\nC 订阅与省流版\nD 资讯渠道\nE 分享与工具\nF 原理与历史\nZ 参考资料与来源"

# 1) s5 附录卡
s5 = prs.slides[4]
for sh in s5.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("A 常用术语"):
        settext(sh, TOC); print("s5 card ok")

# 2) 各附录分隔页侧栏
for n in (54,59,80,88,93,110,121):
    s=prs.slides[n-1]
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text.strip()
        if sh.left/EMU>22 and sh.top/EMU>4.9 and re.match(r"^[A-Z]\s{2}", t):
            key=t[0]
            new={"A":"A  常用术语","B":"B  公司档案","C":"C  订阅与省流版","D":"D  资讯渠道",
                 "E":"E  分享与工具","F":"F  原理与历史","Z":"Z  参考资料与来源"}[key]
            if t!=new: settext(sh,new); print(n,"side",t,"->",new)

# 3) s52 / s53 附录索引行
s52=prs.slides[51]
for sh in s52.shapes:
    if sh.has_text_frame and "附录" in sh.text_frame.text and "A1–A4" in sh.text_frame.text:
        settext(sh,"附录  常用术语 A1–A4 · 公司档案 B1–B13（含专题页）· 订阅与省流 C1–C5（含 C1b、C4b）· 资讯渠道 D1–D4 · 分享与工具 E1–E4 · 原理与历史 F1–F9（含 F2b）· 参考资料来源 Z", size=14)
        print("s52 ok")
s53=prs.slides[52]
for sh in s53.shapes:
    if sh.has_text_frame and "附录 A–F" in sh.text_frame.text:
        settext(sh,"附录 A–F：常用术语、公司档案、订阅与省流、资讯渠道、分享与工具、原理与历史；Z 为参考资料与来源。配套文件夹里有十份鹈鹕原始 HTML、图纸与报告的核对材料。", size=14.5)
        print("s53 ok")
save(prs)
