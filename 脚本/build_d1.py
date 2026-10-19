# -*- coding: utf-8 -*-
"""重写 D1：四种测试题（鹈鹕 / 糖果 / 多合一 / 模型指纹），并加两页实拍。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

from build_v3 import (put, rect, hline, vline, dot, chrome, section_bar,
                      callout, rowband, style_run, L, W,
                      C_TITLE, C_BODY, C_MUTED, C_TEAL, C_ORANGE, C_HEAD,
                      C_LINE, C_WHITE, C_WARM, C_COOL, C_ZEBRA)


def clear(slide):
    spTree = slide.shapes._spTree
    for child in list(spTree):
        if child.tag in (qn("p:sp"), qn("p:cxnSp"), qn("p:pic"),
                         qn("p:graphicFrame"), qn("p:grpSp")):
            spTree.remove(child)


def scale(slide, idx, page):
    """左上角 A4 式比例尺：不解释图表时也能让人知道“这不是真实数据”。"""
    pass


# ================================================================ D1-a  四题
def page_four_tests(slide, page):
    chrome(
        slide, page,
        "D1  四种测试题：各测一件事",
        "同一批模型，四种题考的是四种不同的东西；每题都要写清“能推出什么、推不出什么”",
        "来源：本地十配置盲测；CLIProxyAPI 讨论 #3937（糖果题原题与 21/29 两种答案）；"
        "lm.ikale.io 与 ai-model-fingerprint（指纹方法）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页是后面三页的目录，也是整个附录的评价口径。\n\n"
            "四种题，各测一件事：\n"
            "一，鹈鹕动画——一句话出可用代码。考的是“一次成型”的能力：它要同时写出 SVG/HTML、"
            "把自行车结构画对、让动作循环。判据是动作连贯、部件对得上。它推不出工程判断力。\n\n"
            "二，糖果题——组合计数加最坏情况推理。原题是三种口味乘两种形状，问最少取几颗。"
            "正确答案 21，取 9 颗圆形加 12 颗五角星。**这题真正的陷阱在一句话**："
            "“形状靠手感可以分辨”——意思是你可以挑形状。漏读这句，就会按“完全随机摸”去算，"
            "得出 29。所以在社区里，答 29 常被当成“理解力掉了一档”的标志。\n\n"
            "三，多合一题——长指令遵循。把七八个互不相干的元素塞进一句话：鹈鹕在孙悟空的飞机上、"
            "驮着秦始皇、秦始皇骑北极熊、手里拿螺丝刀、气泡里喊糖课题的答案。考的是“每个元素是否都落实、"
            "关系有没有搞错”。它不测深度推理，只测听不听话。\n\n"
            "四，模型指纹——不看答案对不对，看它“随机”选数的分布。原理下一页讲。\n\n"
            "最后一行的评价纪律是原来 D1 的内容，压缩成一句：模型、客户端、设置、原始输出、评价规则都要留档，"
            "单次结果不能代表稳定水平。"
        ))

    rect(slide, L, 1.90, W, 0.40, C_HEAD)
    for x, w, t in [(0.72, 2.35, "测试题"), (3.17, 2.95, "考的是什么"),
                    (6.22, 3.30, "怎么判"), (9.62, 3.05, "推不出什么")]:
        put(slide, x, 1.90, w - 0.10, 0.40, t, size=13, bold=True, color=C_WHITE)

    rows = [
        ("鹈鹕动画", "一句话出可用代码；视觉常识（自行车结构对不对）",
         "动作是否连贯、部件是否对得上、能不能跑", "工程判断力；长任务稳定性", C_TEAL),
        ("糖果题", "组合计数 + 最坏情况；更关键是读懂“形状可凭手感挑”",
         "答案 21（取 9 圆 + 12 星）；答 29 = 漏读那句", "整体数学水平", C_ORANGE),
        ("多合一题", "长指令遵循：七八个互不相干的元素同时落位",
         "元素齐不齐、关系对不对（谁骑谁、喊什么）", "深度推理——它只测“听不听话”", C_ORANGE),
        ("模型指纹", "它“随机”选数的分布特征，也就是模型的“口音”",
         "三次 1–355 选数，比对分布形状", "证不了具体版本，只能缩小到家族", C_HEAD),
    ]
    y = 2.30
    for i, (a, b, c, d, col) in enumerate(rows):
        rowband(slide, L, y, W, 0.86, i)
        rect(slide, L, y, 0.06, 0.86, col)
        put(slide, 0.74, y, 2.30, 0.86, a, size=13.5, bold=True, color=C_TITLE,
            spacing=1.1)
        put(slide, 3.17, y, 2.85, 0.86, b, size=10.5, color=C_BODY, spacing=1.18)
        put(slide, 6.22, y, 3.25, 0.86, c, size=10.5, color=C_BODY, spacing=1.18)
        put(slide, 9.62, y, 3.05, 0.86, d, size=10.5, color=C_MUTED, spacing=1.18)
        y += 0.86

    rect(slide, L, 5.80, W, 0.44, C_ZEBRA)
    put(slide, 0.78, 5.80, W - 0.30, 0.22, [
        ("评价纪律：", 9, True, C_TEAL),
        ("模型、客户端、设置、原始输出、评价规则全部留档；单次结果不代表稳定水平", 9,
         False, C_BODY)], spacing=1.0, anchor=MSO_ANCHOR.TOP)
    put(slide, 0.78, 6.00, W - 0.30, 0.22, [
        ("案例口径：", 9, True, C_ORANGE),
        ("图纸——mm/cm 经复核不计为问题、21 条待核未计；报告——汇总表口径，112 元差额单独说明", 9,
         False, C_BODY)], spacing=1.0, anchor=MSO_ANCHOR.TOP)

    callout(slide, "要点", "一道题只能测一件事；说清“推不出什么”，结论才立得住")


# ================================================================ D1-b  多合一
def page_allinone(slide, page, img):
    chrome(
        slide, page,
        "D1  多合一题：一张图看完“听不听话”",
        "七八个互不相干的元素塞进一句话：它是不是每个都落实了，关系有没有搞错",
        "来源：本地盲测产出（DeepSeek V4.1 Flash 出图）；提示词见本页备注",
        notes=(
            "讲解与会后阅读\n\n"
            "这一题的设计思路：把“听得懂”拆成可数的七八个点。每个点都能在这张图上直接指出来。\n\n"
            "提示词大意（讲的时候念一遍就够）：\n"
            "“画一张孙悟空驾驶的飞机（机身写‘筋斗云 WUKONG AIR’），飞机上有一辆自行车，"
            "一只鹈鹕骑着自行车，鹈鹕背上驮着秦始皇，秦始皇骑着一只北极熊，"
            "秦始皇手里拿着螺丝刀在打螺丝，鹈鹕嘴里冒出气泡，气泡里写‘21！’；"
            "背景是蓝天白云，画面里还要有一句‘THE IMPOSSIBLE SKY EXPEDITION / 云上奇航’。”\n\n"
            "判据就是逐条对：孙悟空在不在驾驶舱（✓）、飞机机身有没有那个名字（✓）、"
            "鹈鹕是不是在骑自行车（✓）、自行车是不是在飞机上（✓）、秦始皇是不是骑在北极熊上（✓）、"
            "螺丝刀和螺丝有没有（✓）、气泡里是不是 21（✓）、那句英文标题有没有（✓）。\n\n"
            "关键点：这一题几乎不考“聪明”，考的是**一次把八件事都记住**。"
            "元素一多，模型常见的失败是“丢件”——画了鹈鹕忘了秦始皇，或者把“骑北极熊”画成“站在旁边”。"
            "所以它特别适合当“长指令稳定性”的指标。\n\n"
            "局限也要讲：这种题判分靠人眼逐条看，主观成分大（画得好不好看会干扰判断）。"
            "所以它适合“数元素”，不适合“打分排名”。"
        ))

    pic = slide.shapes.add_picture(img, Inches(0.60), Inches(1.90), height=Inches(4.02))
    pic.line.color.rgb = RGBColor.from_string(C_LINE)
    pic.line.width = Pt(0.75)
    h = pic.height / 914400

    put(slide, 0.62, 5.98, pic.width / 914400, 0.34,
        "八件事都落位；注意“骑”这个关系——鹈鹕骑车、秦始皇骑熊，不是站在旁边。",
        size=9, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    # 右：逐项判据
    rx = 7.40
    section_bar(slide, rx, 1.90, 5.33, "逐条对：八个点", h=0.36, size=12.5)
    checks = [
        ("孙悟空在驾驶舱", "人＋机舱，不是站在机翼上"),
        ("机身有“筋斗云 / WUKONG AIR”", "文字要落在机身上，不是飘在空中"),
        ("鹈鹕骑着自行车", "腿要踩踏板，不是悬空"),
        ("自行车在飞机上", "层级关系：飞机 > 车 > 鸟 > 人 > 熊"),
        ("秦始皇骑北极熊", "又是“骑”，且熊要有四足着地感"),
        ("手拿螺丝刀在打螺丝", "两个物体＋一个动作，最容易丢"),
        ("气泡里写 21！", "数字与标点都要对"),
        ("英文标题与中文副题", "两行字都在，且拼写对"),
    ]
    y = 2.32
    for i, (k, v) in enumerate(checks):
        put(slide, rx, y, 0.22, 0.22, "✓", size=9, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 0.24, y - 0.01, 3.30, 0.22, k, size=9.5, bold=True,
            color=C_TITLE, anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 0.24, y + 0.19, 5.05, 0.30, v, size=8.5, color=C_MUTED,
            anchor=MSO_ANCHOR.TOP)
        y += 0.46

    rect(slide, rx, 6.02, 5.33, 0.26, C_WARM)
    put(slide, rx + 0.14, 6.02, 5.10, 0.26,
        "局限：判分靠人眼数件，画面好看会干扰——适合数元素，不适合排名",
        size=8.5, color=C_BODY, anchor=MSO_ANCHOR.MIDDLE)

    callout(slide, "结论", "元素越多，越容易“丢件”：这是长指令稳定性最直白的指标")


# ================================================================ D1-c  指纹
def page_fingerprint(slide, page, img):
    chrome(
        slide, page,
        "D1  模型指纹：355 个数字里的“口音”",
        "让模型自己“随机”报 294 个 1–355 的整数，三次一比对：不同的模型，报数的偏好不一样",
        "来源：lm.ikale.io（模型指纹检测）；github.com/qqqqqf-q/ai-model-fingerprint；"
        "CLIProxyAPI 讨论 #3937（516 现象）",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页回答一个很实际的问题：中转站说它是 Opus，你怎么知道真的是？\n\n"
            "办法不是问它“你是谁”（它会照系统提示词回答），而是让它做一件它自己控制不了的事——"
            "“随机”报数。\n\n"
            "具体三道题：让它从 1 到 355 里挑 294 个整数，要求每个位置独立选择、不许从 1 开始数、"
            "不许递增递减、不许用等差循环。三道题分别用中文、英文、日文（或韩文）出，"
            "避免它靠“题型记忆”作答。\n\n"
            "为什么要这么设计：大模型不是真随机数发生器，它的“随机”是从训练分布里采出来的，"
            "于是每个模型都有自己的偏好——某些数字出现得特别多、某些区段被系统性回避。"
            "把三次输出画成直方图，形状就是这只模型的“口音”。同一家的不同代际、"
            "甚至同一模型的不同量化版本，形状都可能不一样。\n\n"
            "能干什么：识别是不是被换了模型、是不是套壳、是不是被降智。"
            "这正好接上 D2 的中转站风险——上一页说“便宜模型冒充贵模型，很难证明”，这一页就是“怎么证”的一种办法。\n\n"
            "边界必须讲：一，它只能把范围缩小到“家族”，证不了具体版本；"
            "二，结果依赖采样温度，温度设成 0 就没意义了；"
            "三，样本要够（一次 294 个数），太短的输出统计不出来；"
            "四，这只是间接证据，不能当合同依据。\n\n"
            "附带一句社区观察（原帖里的）：有人发现走第三方客户端时，推理 token 经常卡在 516，"
            "而直连官方的同一道题能正常思考——这被当作“被截断”的旁证，但官方从未承认。"
            "讲的时候要标清这是“社区观察”，不是结论。"
        ))

    pic = slide.shapes.add_picture(img, Inches(0.60), Inches(1.90), width=Inches(6.72))
    pic.line.color.rgb = RGBColor.from_string(C_LINE)
    pic.line.width = Pt(0.75)
    h = pic.height / 914400

    put(slide, 0.62, 1.90 + h + 0.08, 6.72, 0.44,
        "三道题，中文 / 英文 / 日文各一；右边三个输入框粘回复，点“检测”比对分布",
        size=9, color=C_MUTED, spacing=1.15, anchor=MSO_ANCHOR.TOP)

    rx = 7.55
    section_bar(slide, rx, 1.90, 5.18, "原理三步", h=0.36, size=12.5)
    steps = [
        ("① 出题", "1–355 里独立选 294 个；禁止递增、等差、循环"),
        ("② 采样", "三次（中/英/日），温度不要设 0"),
        ("③ 比对", "画成直方图，与候选库的分布形状比"),
    ]
    y = 2.32
    for k, v in steps:
        put(slide, rx, y, 1.05, 0.22, k, size=9.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 1.10, y - 0.01, 4.05, 0.42, v, size=9, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.52

    section_bar(slide, rx, 3.92, 5.18, "它能干什么 / 不能干什么", color=C_ORANGE,
                h=0.36, size=12.5)
    can = ["看是不是被换了模型、被套壳、被降智", "中转站说“官方原版”时，这是少数能自查的办法"]
    cannot = ["证不了具体版本，只能缩到家族", "温度 0、样本太短，都测不出来"]
    y = 4.34
    for t in can:
        put(slide, rx, y, 0.20, 0.20, "+", size=9, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 0.22, y - 0.01, 4.90, 0.40, t, size=9, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.44
    for t in cannot:
        put(slide, rx, y, 0.20, 0.20, "−", size=9, bold=True, color=C_ORANGE,
            anchor=MSO_ANCHOR.TOP)
        put(slide, rx + 0.22, y - 0.01, 4.90, 0.40, t, size=9, color=C_MUTED,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.44

    rect(slide, rx, 5.86, 5.18, 0.42, C_ZEBRA)
    put(slide, rx + 0.14, 5.86, 4.92, 0.42,
        "社区观察：走第三方时 reasoning token 常卡在 516，官方未承认——只能当旁证",
        size=8.5, color=C_BODY, spacing=1.1, anchor=MSO_ANCHOR.MIDDLE)

    callout(slide, "一句话", "别问它“你是谁”——让它报数，看它的口音", kind="cool")


# ================================================================ assembly
def shift_footers(prs, after_idx, delta):
    for i, s in enumerate(prs.slides):
        if i <= after_idx:
            continue
        for sh in s.shapes:
            if not sh.has_text_frame or sh.left is None:
                continue
            if abs(sh.left / 914400 - 11.73) > 0.05 or sh.width / 914400 > 1.25:
                continue
            p = sh.text_frame.paragraphs[0]
            if p.runs and p.runs[0].text.strip().isdigit():
                p.runs[0].text = str(int(p.runs[0].text.strip()) + delta)


def main(src, dst):
    prs = Presentation(src)
    blank = next(l for l in prs.slide_layouts if l.name == "Blank")

    # 1) 重写 D1（index 69）
    footer = None
    for sh in prs.slides[69].shapes:
        if sh.has_text_frame and sh.left is not None:
            if abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.25:
                footer = sh.text_frame.text.strip()
    clear(prs.slides[69])
    page_four_tests(prs.slides[69], footer or 68)
    print("rebuilt D1 index 69 footer", footer)

    # 2) 追加两页，随后挪到 D1 之后
    s2 = prs.slides.add_slide(blank)
    page_allinone(s2, 71, "assets/shot_allinone.png")
    i2 = len(prs.slides) - 1

    s3 = prs.slides.add_slide(blank)
    page_fingerprint(s3, 72, "assets/shot_fp.jpg")
    i3 = len(prs.slides) - 1

    lst = prs.slides._sldIdLst
    items = list(lst)
    for el in items:
        lst.remove(el)
    order = (items[:70]
             + [items[i2], items[i3]]
             + items[70:i2] + items[i3+1:])
    for el in order:
        lst.append(el)

    # 3) 页码：新页之后 +2，新页本身补上页号
    shift_footers(prs, 69, 2)

    # 4) D 扉页的页数
    for sh in prs.slides[68].shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if "共 13 页" in r.text:
                    r.text = r.text.replace("共 13 页", "共 15 页")

    prs.save(dst)
    print("saved", dst, len(prs.slides))
    for i, s in enumerate(prs.slides, 1):
        if 68 <= i <= 75:
            t = ""
            for sh in s.shapes:
                if sh.has_text_frame and sh.name == "TextBox 4":
                    t = sh.text_frame.text.split("\n")[0]
            print(f"  {i:3d}  {t[:54]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
