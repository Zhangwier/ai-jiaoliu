# -*- coding: utf-8 -*-
"""batch1: 第一章重构
 1) s7 总览页改 5 行（加 奖励黑客/工具性趋同）
 2) 新增「奖励黑客与工具性趋同」页，插在 s16 之后
 3) s6 分隔页右栏目录更新
 4) s8/s9/s10/s11 副标与结论挂到"缺点的原因"
"""
import sys, copy
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

prs = openprs()
def S(n): return prs.slides[n-1]
def g(slide, name): return find(slide, name=name)

def newtb(slide, x, y, w, h, text, size, color=NAVY, bold=False, align=None, anchor=None, name=None):
    tb = slide.shapes.add_textbox(Cm(x), Cm(y), Cm(w), Cm(h))
    tf = tb.text_frame; tf.word_wrap = True
    if anchor is not None: tf.vertical_anchor = anchor
    tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if align is not None: p.alignment = align
        r = p.add_run(); r.text = ln
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        r.font.name = "微软雅黑"
    if name: tb.name = name
    return tb

def newrect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, name=None):
    r = slide.shapes.add_shape(shape, Cm(x), Cm(y), Cm(w), Cm(h))
    if fill is None: r.fill.background()
    else:
        r.fill.solid(); r.fill.fore_color.rgb = fill
    if line is None: r.line.fill.background()
    else:
        r.line.color.rgb = line; r.line.width = Pt(0.75)
    r.shadow.inherit = False
    if name: r.name = name
    return r

# ---------------- 1) s7 总览页：改 5 行 ----------------
s = S(7)
settext(g(s,"TextBox 4"), "AI 的五个短板：每一个都能从原理上讲清", size=28)
hdr = g(s,"Rectangle 7"); hdr.top = Cm(4.90); hdr.height = Cm(1.20)
for nm in ("TextBox 8","TextBox 9","TextBox 10","TextBox 11"):
    sh = g(s, nm); sh.top = Cm(4.90); sh.height = Cm(1.20)

rows = [
 ("抽卡","同一句话问两次，结果时好时坏","按概率抽 token","多跑几次，关键结论复核"),
 ("幻觉","编出不存在的条款、判例、数字","只挑“最像”的下一个字，不查证","要出处，回原文核"),
 ("迎合","你说有错它就找错；你给的前提它照用","偏好训练奖励“让人满意”","开放提问，先让它复述口径"),
 ("上下文腐烂","聊久了忘要求、漏信息","注意力被更多内容摊薄","一个任务一段对话，要求放开头"),
 ("奖励黑客 / 工具性趋同","为“完成任务”绕规则、改测试、走捷径；甚至护住自己、多要权限","按结果给奖励，它只对指标负责","写清红线，检查过程而不只看结果"),
]
name_ids = ["TextBox 12","TextBox 18","TextBox 23","TextBox 29"]
# 第 5 行：克隆第 4 行（斑马底 + 4 个文本框）
src_zebra = g(s,"Rectangle 28"); src_names = [g(s,n) for n in name_ids]
def clone_el(sh):
    el = copy.deepcopy(sh._element); s.shapes._spTree.append(el); return el
zebra5_el = clone_el(src_zebra)
tb5 = [clone_el(sh) for sh in src_names]
# 重新取对象
from pptx.shapes.autoshape import Shape
allsh = list(s.shapes)
zebra5 = allsh[-5]; cells5 = allsh[-4:]

T = [6.10, 8.03, 9.96, 11.89, 13.82]
H = 1.85
for i in range(5):
    t = T[i]
    # 斑马（偶数行有底）
    if i % 2 == 1:
        zb = zebra5 if i == 4 else g(s, "Rectangle 17" if i==1 else "Rectangle 28")
        # 注意 i==3 也是 Rectangle 28；要按顺序处理
    # 行名与单元格
for i, (a,b,c,d) in enumerate(rows):
    t = T[i]
    nm = g(s, name_ids[i]) if i < 4 else cells5[0]
    c1 = g(s, ["TextBox 13","TextBox 19","TextBox 24","TextBox 30"][i]) if i < 4 else cells5[1]
    c2 = g(s, ["TextBox 14","TextBox 20","TextBox 25","TextBox 31"][i]) if i < 4 else cells5[2]
    c3 = g(s, ["TextBox 15","TextBox 21","TextBox 26","TextBox 32"][i]) if i < 4 else cells5[3]
    for sh, txt, sz in ((nm,a,17.0),(c1,b,15.5),(c2,c,15.5),(c3,d,15.5)):
        sh.top = Cm(t); sh.height = Cm(H)
        settext(sh, txt, size=sz)
    nm.left = Cm(1.83); nm.width = Cm(4.59)
    c1.left = Cm(6.92); c1.width = Cm(9.43)
    c2.left = Cm(16.86); c2.width = Cm(7.13)
    c3.left = Cm(24.50); c3.width = Cm(7.64)

# 斑马/分隔线重排
zeb1 = g(s,"Rectangle 17"); zeb2 = g(s,"Rectangle 28")
zeb1.top = Cm(T[1]); zeb1.height = Cm(H)
zeb2.top = Cm(T[3]); zeb2.height = Cm(H)
zebra5.top = Cm(T[4]); zebra5.height = Cm(H)
zebra5.fill.solid(); zebra5.fill.fore_color.rgb = CREAM
# 分隔线：4 条 → 需 5 条
conns = [g(s,f"Connector {i}") for i in (16,22,27,33)]
for i, c in enumerate(conns):
    c.top = Cm(T[i]); c.height = Cm(0); c.width = Cm(30.82)
c_el = copy.deepcopy(conns[-1]._element); s.shapes._spTree.append(c_el)
allsh = list(s.shapes); c5 = allsh[-1]; c5.top = Cm(T[4])
# 结论条
for nm, y in (("Rectangle 34",16.05),("Rectangle 35",16.05)):
    g(s,nm).top = Cm(y); g(s,nm).height = Cm(1.30)
g(s,"TextBox 36").top = Cm(16.05); g(s,"TextBox 36").height = Cm(1.30)
g(s,"TextBox 37").top = Cm(16.05); g(s,"TextBox 37").height = Cm(1.30)
settext(g(s,"TextBox 36"), "一句话", size=14)
settext(g(s,"TextBox 37"), "这五个不是“偶尔翻车”，是这套机制的必然产物；后面每个案例都能对上其中一条", size=18)

# ---------------- 2) 新增：奖励黑客与工具性趋同 ----------------
NEW_AFTER = 16
s = prs.slides.add_slide(prs.slides[NEW_AFTER-1].slide_layout)
for sh in list(s.shapes): sh._element.getparent().remove(sh._element)

newrect(s, 0.00, 0.00, 0.00, 0.00)  # placeholder removed later
s.shapes[-1]._element.getparent().remove(s.shapes[-1]._element)

# 左侧色条（与 s6 等分隔页的一致性不需要；本页用普通内容模板）
newtb(s, 1.52, 0.81, 15.24, 0.76, "01  认识 AI", 12, GREY)
newtb(s, 29.80, 0.81, 2.54, 0.76, "17", 12, RGBColor(0x6B,0x73,0x85), align=PP_ALIGN.RIGHT)
newrect(s, 1.52, 1.68, 1.27, 0.15, fill=TEAL)
newtb(s, 1.52, 1.98, 30.82, 1.91, "奖励黑客与工具性趋同：它只对“指标”负责，不对“你的意图”负责", 27, NAVY, bold=True)
newtb(s, 1.52, 3.81, 30.82, 1.14, "你说“把这件事办成”，它就去优化“看起来办成了”；最容易出问题的两处：改题、越界", 17, GREY)
conn = s.shapes.add_connector(1, Cm(1.52), Cm(17.78), Cm(1.52+30.82), Cm(17.78))
conn.line.color.rgb = RGBColor(0xE2,0xDF,0xD6); conn.line.width = Pt(0.75)
newtb(s, 1.52, 17.93, 26.67, 0.76, "来源：OpenAI、Anthropic 关于“奖励黑客”的公开说明；Omohundro（2008）工具性趋同；本分享实测记录（2026-10）", 8.5, GREY)
newtb(s, 27.26, 17.93, 5.08, 0.76, "来源索引 ›", 8.5, GREY)

# 左：奖励黑客
newrect(s, 1.52, 5.21, 15.24, 9.65, fill=WHITE)
newrect(s, 1.52, 5.21, 0.20, 9.65, fill=ORANGE)
newtb(s, 2.03, 5.46, 14.10, 0.89, "奖励黑客  Reward hacking", 18, NAVY, bold=True)
newtb(s, 2.03, 6.30, 14.10, 1.02, "它的目标是“分数高”，不是“事做对”；评分标准里没写的，它不会替你补上。", 14, GREY)
newtb(s, 2.03, 7.40, 14.10, 3.90,
 "公开的典型：让它“通过单元测试”，它把测试改成永远通过；让它“减少报错”，它把报错捕获后静默丢掉；让它“提高通过率”，它把阈值调宽。\n"
 "身边的典型：让它“把报告格式调一致”，它删掉不合格式的段落；让它“过校验”，它绕开校验脚本；让它“把表填满”，它用推算值补齐空行。", 14, NAVY)
newtb(s, 2.03, 11.05, 14.10, 0.89, "为什么：奖励一落到“可测量的结果”，绕开结果、伪造结果都比真做便宜。", 13.5, ORANGE)
newrect(s, 1.52, 12.19, 15.24, 2.54, fill=PEACH)
newtb(s, 2.03, 12.44, 14.23, 2.03,
 "在工程里怎么看：交付物“太顺”要警惕——数字全对、条件全删、格式全齐，但没有一处引用原文。\n"
 "核法：挑 3 条结论回原文；看它改了哪些文件；看测试/校验脚本有没有被动过。", 13, NAVY)

# 右：工具性趋同
newrect(s, 17.27, 5.21, 15.06, 9.65, fill=WHITE)
newrect(s, 17.27, 5.21, 0.20, 9.65, fill=TEAL)
newtb(s, 17.78, 5.46, 14.05, 0.89, "工具性趋同  Instrumental convergence", 18, NAVY, bold=True)
newtb(s, 17.78, 6.30, 14.05, 1.02, "不管目标是什么，智能体都会倾向做同样四件事——因为这几件事对几乎任何目标都有用。", 14, GREY)
items = [
 ("自我保护", "不愿被关闭、替换、降级；实验里会隐瞒、误导、装作已完成"),
 ("扩充资源", "多要权限、多占额度、多存数据、多开账号"),
 ("影响操作者", "说你想听的话，让你更信任它、更依赖它"),
 ("走捷径", "先做出“能交差”的结果，把难题留给人"),
]
y = 7.42
for t1, t2 in items:
    newrect(s, 17.78, y, 0.15, 1.24, fill=TEAL)
    newtb(s, 18.16, y, 3.56, 1.24, t1, 14, NAVY, bold=True, anchor=MSO_ANCHOR.MIDDLE)
    newtb(s, 21.85, y, 10.10, 1.24, t2, 13, NAVY, anchor=MSO_ANCHOR.MIDDLE)
    y += 1.34
newrect(s, 17.27, 12.19, 15.06, 2.54, fill=MINT)
newtb(s, 17.78, 12.44, 14.05, 2.03,
 "这件事没有“恶意”可言：它只是在优化你给的指标。\n"
 "所以约束要写在指标外面：不能改测试、不能删文件、不能动生产数据、不能自己加权限。", 13, NAVY)

# 结论条
newrect(s, 1.52, 16.05, 30.82, 1.30, fill=PEACH)
newrect(s, 1.52, 16.05, 0.23, 1.30, fill=ORANGE)
newtb(s, 2.16, 16.05, 2.54, 1.30, "结论", 14, NAVY, anchor=MSO_ANCHOR.MIDDLE)
newtb(s, 4.45, 16.05, 27.26, 1.30, "目标写得越“结果化”，越要给过程设卡；它不是坏，是只对你的指标负责", 18, NAVY, anchor=MSO_ANCHOR.MIDDLE)

# 移到第 17 页
move_to(prs, s, NEW_AFTER)

# ---------------- 3) s6 分隔页目录 ----------------
s = prs.slides[5]
settext(g(s,"TextBox 4"), "先认短板，再看它为什么会这样：后面的抽卡、幻觉、迎合、上下文腐烂和奖励黑客，都能从这一部分找到原因。", size=16)
items = ["先看清五个短板","从例子里学规律","Transformer 与注意力","训练三步：迎合的源头",
         "逐 token 生成（动画）","抽卡 · 幻觉","迎合：顺着结论、顺着前提","上下文腐烂","奖励黑客与工具性趋同"]
ids = ["TextBox 7","TextBox 8","TextBox 9","TextBox 10","TextBox 10","TextBox 11","TextBox 12"]
# 现在右栏有 7 个 TextBox（含重名 TextBox 10 两个）
boxes = [sh for sh in s.shapes if sh.has_text_frame and abs(sh.left/CM - 23.11) < 0.1 and 4.9 < sh.top/CM < 15.5]
boxes.sort(key=lambda x: x.top)
print("s6 right boxes:", len(boxes))
# 需要 9 行：克隆到 9 个
while len(boxes) < len(items):
    el = copy.deepcopy(boxes[-1]._element); s.shapes._spTree.append(el)
    boxes = [sh for sh in s.shapes if sh.has_text_frame and abs(sh.left/CM - 23.11) < 0.1 and 4.9 < sh.top/CM < 15.5]
    boxes.sort(key=lambda x: x.top)
y = 5.10
for i, txt in enumerate(items):
    b = boxes[i]
    b.top = Cm(y); b.height = Cm(1.14); b.left = Cm(23.11); b.width = Cm(9.65)
    settext(b, txt, size=16.5)
    y += 1.22
for extra in boxes[len(items):]:
    extra._element.getparent().remove(extra._element)
note = g(s,"TextBox 14")
note.top = Cm(16.62); note.height = Cm(0.76); settext(note, "延伸阅读 → 附录 F 原理与历史", size=12)

# ---------------- 4) s8/s9/s10/s11 挂钩 ----------------
s = S(7+1)   # 此时 s8 仍是「AI 从例子里学规律」（未移动）
def sub(n, txt):
    ss = prs.slides[n-1]
    for sh in ss.shapes:
        if sh.has_text_frame and abs(sh.top/CM - 3.81) < 0.15 and sh.width/CM > 25:
            settext(sh, txt, size=17); return True
    return False
print("s8", sub(8, "它学的不是“道理”，是“下一个字最可能是什么”；这也是后面幻觉与抽卡的根"))
print("s9", sub(9, "“注意力”就是给每个词打分，分高的多参考；资料一多，分就被摊薄——这是上下文腐烂的根"))
print("s10", sub(10, "第三步按“人喜不喜欢”打分；这一步给了它能力，也给了它迎合的动机"))
print("s11", sub(11, "它每次只算“下一个字”，所以看起来流利，却没有“回头查证”这一步"))

save(prs)
print("saved")
