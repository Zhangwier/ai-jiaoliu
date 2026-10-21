# -*- coding: utf-8 -*-
"""同步目录/分隔页：页数、F2b 编号、C 副标、B 副标"""
import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
prs = openprs()

# ---- 1) 目录与侧栏：公司档案（十七页）->（二十页）；原理与历史（九页）->（十页）
cnt = 0
for i, sl in enumerate(prs.slides, start=1):
    for sh in sl.shapes:
        if not sh.has_text_frame: continue
        tf = sh.text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                o = r.text
                n = o.replace("B  公司档案（十七页）", "B  公司档案（二十页）") \
                     .replace("F  原理与历史（九页）", "F  原理与历史（十页）")
                if n != o:
                    r.text = n; cnt += 1
print("sidebar fixed", cnt)

# ---- 2) s5 目录说明补「奖励黑客」等
s = prs.slides[4]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("原理，以及"):
        settext(sh, "五个短板与它们的原理成因", size=15)
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("鹈鹕盲测"):
        settext(sh, "鹈鹕盲测、模型 × 外壳、榜单、各家脾气", size=15)
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("什么活适合"):
        settext(sh, "三类活、任务书、上下文、资料库、工作流", size=15)

# ---- 3) B 分隔页副标（s59）
s = prs.slides[58]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("B1–B13"):
        settext(sh, "B1–B13 每家一页：模型线、可用入口、强项、边界；OpenAI、Anthropic、Meta（Muse）、DeepSeek、Qwen 另各有专题页，降智与封号紧跟在对应公司之后。", size=15)

# ---- 4) C 分隔页副标（s80）
s = prs.slides[79]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("价格与地区政策"):
        settext(sh, "价格与地区政策以官方页面为准。C1–C4b、C7 共 7 页：四件事自查、照着做七步、地区与支付、注册、价格、省钱的路子、省流版。", size=15)

# ---- 5) F 分隔页副标（s110）
s = prs.slides[109]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip().startswith("共 9 页"):
        settext(sh, "共 10 页：F1 发展史 · F2 神经网络怎么学 · F2b 神经网络结构 · F3 token 与向量 · F4 Transformer 一层 · F5 注意力 · F6 稠密与 MoE · F7 规模定律 · F8 对齐与推理 · F9 多模态。", size=15)

# ---- 6) 神经网络页（现 s113）改成 F2b 并补页码
s = prs.slides[112]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.strip() == "神经网络：很多层“加权求和＋非线性”叠起来":
        settext(sh, "F2b  神经网络的结构：很多层“加权求和＋非线性”叠起来", size=25)
# 补页码框（照抄标准位置）
import copy
src = None
for sh in prs.slides[111].shapes:
    if sh.has_text_frame and sh.left/360000>29 and sh.top/360000<1.6 and re.fullmatch(r"\d+", sh.text_frame.text.strip()):
        src = sh; break
if src is not None:
    new = copy.deepcopy(src._element)
    s.shapes._spTree.append(new)
    ns = s.shapes[-1]
    ns.text_frame.paragraphs[0].runs[0].text = "113"
print("done")
save(prs)
