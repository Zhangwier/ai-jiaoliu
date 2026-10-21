# -*- coding: utf-8 -*-
"""第二轮审稿修正：软化 P8、标题缩短、韩国案例事实/解读分离、判断框架补适用条件。"""
import sys
from pptx import Presentation
from pptx.util import Pt


def by_name(slide, name):
    for sh in slide.shapes:
        if sh.name == name and sh.has_text_frame:
            return sh
    return None


def setp(shape, text, size=None):
    p = shape.text_frame.paragraphs[0]
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)
    if size:
        for r in p.runs:
            r.font.size = Pt(size)
    return True


def main(src, dst):
    prs = Presentation(src)
    S = prs.slides
    n = 0

    # ---- P8：冲淡也是推测，改成中性
    sh = by_name(S[7], "TextBox 86")
    if sh:
        setp(sh, "窗口越长，要同时顾及的内容越多；关键信息能否被稳定取用，是实测问题，不能由权重结构直接推出")
        print("ok P8 结论"); n += 1

    # ---- 标题缩短
    ttl = {18: "同一牌子，网页版、会员和 API 常不是同一个模型",
           25: "Opus 5.5 一句话出视频：写代码画帧，再合成 MP4",
           86: "D3  案例（一）：Hugging Face 事件与诉讼",
           87: "D3  案例（二）：人指挥 Agent 入侵韩国金融机构",
           49: "个人 Agent 和企业 Agent 的四个差别"}
    for idx, t in ttl.items():
        sh = by_name(S[idx], "TextBox 4")
        if sh:
            setp(sh, t); print("ok 标题", idx + 1); n += 1

    # ---- P87 韩国案例：副标题与右栏改为事实 / 解读分开
    sh = by_name(S[87], "TextBox 5")
    if sh:
        setp(sh, "9 月底至 10 月初，韩国七家金融机构遭入侵；据公开披露与厂商归因，"
                 "操作者是人，Agent 是其使用的工具")
        print("ok P88 副标题"); n += 1
    sh = by_name(S[87], "TextBox 26")
    if sh:
        setp(sh, "和上一页差在哪（本稿解读）")
        print("ok P88 右栏标题"); n += 1
    sh = by_name(S[87], "TextBox 37")
    if sh:
        setp(sh, "工作启示（本稿建议）：内部工具收回公网；认证与会话逐页校验；"
                 "按发布节奏复测，不是一年一次")
        print("ok P88 页脚"); n += 1
    sh = by_name(S[87], "TextBox 41")
    if sh:
        setp(sh, "同一套工具，一边由 AI 自主行动，一边由人指挥：公开信息里的差别在于操作者与授权，"
                 "责任归属仍以各自披露为准")
        print("ok P88 教训"); n += 1

    # ---- P50 / P52：判断框架补适用条件
    sh = by_name(S[48], "TextBox 12")
    if sh:
        setp(sh, "四个真需求：可作为场景筛选的参考（适用于有交付与验收流程的组织）", 12)
        print("ok P50 右标题"); n += 1
    sh = by_name(S[48], "TextBox 42")
    if sh:
        setp(sh, "反过来判断：什么时候该用个人版", 9.5)
        print("ok P50 右下"); n += 1

    # P52 副标题补适用条件
    sh = by_name(S[51], "TextBox 5")
    if sh:
        setp(sh, "多数场景“通用模型 ＋ 检索 ＋ 工作流”就够了；自训门槛不在技术，"
                 "在数据、评测和维护（本稿判断，条件随规模与合规要求变化）")
        print("ok P52 副标题"); n += 1

    prs.save(dst)
    print("applied", n, "->", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
