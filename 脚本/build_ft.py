# -*- coding: utf-8 -*-
"""在 51（两个项目）之后插入“要不要自己训”一页。"""
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

from build_v3 import (put, rect, hline, vline, dot, chrome, section_bar,
                      callout, rowband, style_run, L, W,
                      C_TITLE, C_BODY, C_MUTED, C_TEAL, C_ORANGE, C_HEAD,
                      C_LINE, C_WHITE, C_WARM, C_COOL, C_ZEBRA)


def page_finetune(slide, page):
    chrome(
        slide, page,
        "要不要自己训一个模型：先算清这笔账",
        "多数场景“通用模型 ＋ 检索 ＋ 工作流”就够了；自训的门槛不在技术，在数据、评测和维护",
        "来源：本稿两案例的实测；各厂微调与上下文计费文档；公开的企业落地经验（九成知识问答走 RAG）；本稿整理",
        notes=(
            "讲解与会后阅读\n\n"
            "这一页回答一个很常见的冲动：既然通用模型不懂我们的业务，是不是该自己训一个？\n\n"
            "先把话说明白：微调（fine-tune）改的是“说话方式和任务习惯”，不是“往里灌知识”。"
            "想让模型知道你们单位的定额口径，靠微调效果很差——那是检索该干的事。"
            "这一点很多人搞反，投了三个月才发现模型还是不知道那份文件写了什么。\n\n"
            "左边四个条件，全中才值得做：\n"
            "一，数据不能出内网，且不允许调外部 API；\n"
            "二，任务极窄且量大——比如每天几千条固定格式的分类，边际成本才划得来；\n"
            "三，有稳定的标注和评测能力，能判断“训完到底好没好”；\n"
            "四，通用模型确实做不了——不是“做得不够完美”，是做不了。\n\n"
            "右边三条是不值得的理由，第一条最致命：通用模型大约半年一代，"
            "你花三个月训出来的专用模型，很可能上线就被新一代追平；而你的三个月不会退回来。\n\n"
            "第二是维护：模型要跟着数据漂移重训，要有回归测试，要有人管版本；"
            "第三是评测：如果没有那 50 道题，你根本说不清微调之后是变好还是变差。\n\n"
            "底部的升级顺序是我建议的路径：先用通用模型把任务跑通 → 加检索解决“不知道” → "
            "把固定步骤写成工作流解决“不稳” → 最后如果还不够，再考虑微调。"
            "顺序反了，很容易在错误的地方投入三个月。\n\n"
            "最后一句可以留给听众：先问“我们缺的是知识、是稳定，还是听话”，"
            "三个答案对应三种做法，微调只对应第三个。"
        ))

    # 左：四个条件
    section_bar(slide, L, 1.86, 6.05, "什么时候值得自己训：四条全中", h=0.36, size=13)
    conds = [
        ("数据不能出内网", "且不允许走外部 API，这一条是硬约束，不是偏好"),
        ("任务极窄且量大", "比如每天几千条固定格式分类，边际成本才划得来"),
        ("有标注与评测", "能判断“训完到底好没好”，否则只是换个说法"),
        ("通用模型做不到", "注意：是做不到，不是“做得不够完美”"),
    ]
    y = 2.28
    for k, v in conds:
        dot(slide, 0.70, y + 0.12, 0.10, C_TEAL)
        put(slide, 0.88, y, 1.85, 0.22, k, size=10.5, bold=True, color=C_TEAL,
            anchor=MSO_ANCHOR.TOP)
        put(slide, 2.78, y - 0.01, 3.75, 0.42, v, size=9.5, color=C_BODY,
            spacing=1.15, anchor=MSO_ANCHOR.TOP)
        y += 0.52

    rect(slide, L, 4.46, 6.05, 0.74, C_WARM)
    put(slide, 0.74, 4.50, 5.65, 0.22, "先纠正一个常见误解", size=9.5, bold=True,
        color=C_ORANGE, anchor=MSO_ANCHOR.TOP)
    put(slide, 0.74, 4.74, 5.65, 0.44,
        "微调改的是“说话方式与任务习惯”，不是“往里灌知识”。想让模型知道你们的定额口径，"
        "该做的是检索，不是微调。", size=9, color=C_BODY, spacing=1.18,
        anchor=MSO_ANCHOR.TOP)

    # 右：三条代价
    section_bar(slide, 6.95, 1.86, 5.78, "为什么多数不值得：三条硬账", color=C_ORANGE,
                h=0.36, size=13)
    costs = [
        ("追不上换代", "通用模型约半年一代；训三个月刚上线，可能已被新一代追平"),
        ("维护不断", "数据会漂移，要重训、要回归测试、要管版本，都是常年投入"),
        ("说不清好坏", "没有固定评测题，训完只能说“感觉好一点”，没法对上级交代"),
    ]
    y = 2.28
    for k, v in costs:
        rect(slide, 6.95, y, 5.78, 0.80, C_WHITE)
        rect(slide, 6.95, y, 0.05, 0.80, C_ORANGE)
        put(slide, 7.10, y + 0.07, 1.55, 0.22, k, size=10.5, bold=True,
            color=C_TITLE, anchor=MSO_ANCHOR.TOP)
        put(slide, 8.70, y + 0.07, 3.88, 0.66, v, size=9, color=C_BODY, spacing=1.18,
            anchor=MSO_ANCHOR.TOP)
        y += 0.88

    rect(slide, 6.95, 4.92, 5.78, 0.28, C_ZEBRA)
    put(slide, 7.10, 4.92, 5.50, 0.28,
        "量级参考：三个月的调优人力，通常够把一条工作流做扎实并跑完两轮评测。",
        size=8.5, color=C_MUTED, anchor=MSO_ANCHOR.MIDDLE)

    # 底：升级顺序
    rect(slide, L, 5.44, W, 0.70, C_COOL)
    put(slide, 0.80, 5.48, 2.20, 0.22, "建议的升级顺序", size=9.5, bold=True,
        color=C_TEAL, anchor=MSO_ANCHOR.TOP)
    steps = ["① 通用模型跑通任务", "② 加检索解决“不知道”",
             "③ 写工作流解决“不稳”", "④ 还不够再考虑微调"]
    x = 0.80
    for i, t in enumerate(steps):
        w = 2.86
        put(slide, x, 5.72, w, 0.26, t, size=9, bold=(i == 3),
            color=C_TEAL if i < 3 else C_ORANGE, anchor=MSO_ANCHOR.TOP)
        x += w + 0.02

    callout(slide, "判断", "先问“缺的是知识、稳定，还是听话”——微调只对应最后一个", kind="warn")


def main(src, dst):
    prs = Presentation(src)
    blank = next(l for l in prs.slide_layouts if l.name == "Blank")

    s = prs.slides.add_slide(blank)
    page_finetune(s, 52)
    new = len(prs.slides) - 1

    lst = prs.slides._sldIdLst
    items = list(lst)
    for el in items:
        lst.remove(el)
    order = items[:51] + [items[new]] + items[51:new]
    for el in order:
        lst.append(el)

    # 插在第 51 页之后：后面页码 +1（原 52 起的 ops 页变成 53）
    for i, sl in enumerate(prs.slides):
        if i <= 50:
            continue
        for sh in sl.shapes:
            if not sh.has_text_frame or sh.left is None:
                continue
            if abs(sh.left / 914400 - 11.73) > 0.05 or sh.width / 914400 > 1.25:
                continue
            p = sh.text_frame.paragraphs[0]
            if p.runs and p.runs[0].text.strip().isdigit():
                p.runs[0].text = str(int(p.runs[0].text.strip()) + 1)

    prs.save(dst)
    print("saved", dst, len(prs.slides))
    for i, sl in enumerate(prs.slides, 1):
        if 48 <= i <= 55:
            t = ""
            for sh in sl.shapes:
                if sh.has_text_frame and sh.name == "TextBox 4":
                    t = sh.text_frame.text.split("\n")[0]
            n = None
            for sh in sl.shapes:
                if sh.has_text_frame and sh.left is not None and \
                        abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.25:
                    n = sh.text_frame.text
            print(f"  {i:3d} [{n}] {t[:50]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
