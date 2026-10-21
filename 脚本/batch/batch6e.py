# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
prs = openprs(); s = prs.slides[85]
def g(n): return find(s, name=n)

def rewrite(nm, lines, size=11.5, ls=1.32):
    sh = g(nm); tf = sh.text_frame
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs)[1:]: r._r.getparent().remove(r._r)
    r0 = p0.runs[0]
    r0.text = lines[0]; r0.font.size = Pt(size); r0.font.bold = False
    r0.font.color.rgb = RGBColor(0x2B,0x34,0x45); r0.font.name = "微软雅黑"
    p0.line_spacing = ls
    for ln in lines[1:]:
        p = tf.add_paragraph(); p.line_spacing = ls
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = False
        r.font.color.rgb = RGBColor(0x2B,0x34,0x45); r.font.name = "微软雅黑"

rewrite("TextBox 13", [
 "Cline：ClinePass 首月 $1（之后 $9.99/月）",
 "Codex：Go 档 $8/月，够轻度用",
 "GLM Coding：Lite ¥118/月，限指定工具",
 "各家“首月半价”：先看续费价",
])
rewrite("TextBox 20", [
 "WorkBuddy：每月 500 积分＋每日签到",
 "Trae：每天 100 次（会员 200）",
 "Qoder：2 周试用＋300 次",
 "Codex 免费账号：只有小模型",
])
rewrite("TextBox 27", [
 "开源中转与公益站（L 站常见那几个）",
 "闲鱼共享号、代充、拼车：几元/月",
 "共同点：提示词和资料过别人的服务器",
 "维护靠人，随时可能停",
])

note = {
 "TextBox 15": ["适合：试一个月，不打算长期绑定。", "关键：设日历提醒，别被自动续费扣住。"],
 "TextBox 22": ["适合：任务不重、能每天手动领一次。", "代价：额度有上限，赶 deadline 不一定够。"],
 "TextBox 29": ["适合：只跑公开材料、不涉单位数据。", "别用于：合同、图纸、报告、客户信息。"],
}
for nm, lines in note.items():
    sh = g(nm); tf = sh.text_frame
    for p in list(tf.paragraphs)[1:]: p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    for r in list(p0.runs)[1:]: r._r.getparent().remove(r._r)
    r0 = p0.runs[0]; r0.text = lines[0]
    r0.font.size = Pt(11.5); r0.font.bold = False
    r0.font.color.rgb = NAVY; r0.font.name="微软雅黑"
    p0.line_spacing = 1.32
    p = tf.add_paragraph(); p.line_spacing = 1.32
    r = p.add_run(); r.text = lines[1]
    r.font.size = Pt(11.5); r.font.bold=False; r.font.color.rgb = NAVY; r.font.name="微软雅黑"
    sh.top = Cm(9.22); sh.height = Cm(1.90)

# body 高度 2.60 → 底 8.75；分隔线 8.90
for nm in ("TextBox 13","TextBox 20","TextBox 27"):
    g(nm).top = Cm(6.10); g(nm).height = Cm(2.60)
for nm in ("Rectangle 14","Rectangle 21","Rectangle 28"):
    g(nm).top = Cm(8.92); g(nm).height = Cm(0.04)
# 三张卡高 14.85
for nm in ("Rectangle 9","Rectangle 16","Rectangle 23"):
    g(nm).height = Cm(10.32)
save(prs); print("ok")
