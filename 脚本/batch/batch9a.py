# -*- coding: utf-8 -*-
import sys, copy
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs()

NO_CNT = {
 54:  "模型本身、怎么交代任务、Agent 与工具、常见误解。",
 80:  "价格与地区政策以官方页面为准：四件事自查、照着做七步、地区与支付、注册、价格、省钱的路子、省流版。",
 88:  "社区看风向，GitHub 看版本，B 站看演示，X 看一手。",
 93:  "都可以直接拿去用：降智测试（附两个自查工具） · 中转站与代理池 · 护栏与破限 · 开源落地",
 110: "发展史 · 神经网络怎么学 · 神经网络结构 · token 与向量 · Transformer 一层 · 注意力 · 稠密与 MoE · 规模定律 · 对齐与推理 · 多模态。",
 121: "Z1 按编号归档；Z2、Z3 是可以直接点开的原链接；正文页脚都指到这里。",
}
import re
for n, txt in NO_CNT.items():
    s = prs.slides[n-1]
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t = sh.text_frame.text.strip()
        if sh.left/EMU < 22 and 12.0 < sh.top/EMU < 14.5 and ("共" in t or "页" in t or "资料" in t):
            settext(sh, txt); print("sub", n, "->", txt[:30])

# 侧栏去掉（N 页）
for n in (54,59,80,88,93,110,121):
    s = prs.slides[n-1]
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t = sh.text_frame.text.strip()
        if (sh.left/EMU > 22 and sh.top/EMU > 4.9 and re.search(r"（[一二三四五六七八九十]+页）", t)):
            new = re.sub(r"（[一二三四五六七八九十]+页）", "", t)
            if t.startswith("C  订阅与省流"): new = "C  订阅与省流版"
            if t.startswith("Z  参考资料"):   new = "Z  参考资料与来源"
            settext(sh, new); print("side", n, "->", new)

# 目录页 s5 标签统一
s5 = prs.slides[4]
for sh in s5.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("A 常用术语"):
        settext(sh, "A 常用术语\nB 公司档案\nC 订阅与省流版\nD 资讯渠道\nE 分享与工具\nF 原理与历史\nZ 参考资料与来源")
        print("s5 toc unified")
save(prs)
