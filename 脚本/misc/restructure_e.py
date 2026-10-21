# -*- coding: utf-8 -*-
"""重排 E 段结构：
   1) E0b（模型架构）→ 移到 01 末尾（四个常见问题 之前）
   2) E0（按任务挑公司）→ 移到 02 里「各家脾气」之后
   3) E 段只留公司介绍：E1…E13；
      OpenAI 数学页紧跟 E1，Opus 出视频页紧跟 E2
"""
import sys
from pptx import Presentation
from pptx.oxml.ns import qn

from build_v3 import (put, rect, chrome, L, W, C_TITLE, C_BODY, C_MUTED,
                      C_TEAL, C_ORANGE, C_HEAD, C_ZEBRA, C_WHITE)


def titles(prs):
    out = []
    for s in prs.slides:
        t = ""
        for sh in s.shapes:
            if sh.has_text_frame and sh.name == "TextBox 4":
                t = sh.text_frame.text.split("\n")[0]
                break
        out.append(t)
    return out


def idx_of(prs, prefix):
    for i, t in enumerate(titles(prs)):
        if t.startswith(prefix):
            return i
    raise SystemExit("not found: " + prefix)


def idx_contains(prs, key):
    """只在“标题”（TextBox 4）里找，避免命中扉页说明文字。"""
    for i, t in enumerate(titles(prs)):
        if t.startswith(key):
            return i
    raise SystemExit("not found: " + key)


def move(prs, frm, to):
    lst = prs.slides._sldIdLst
    items = list(lst)
    el = items[frm]
    lst.remove(el)
    rest = list(lst)
    if to >= len(rest):
        lst.append(el)
    else:
        rest[to].addprevious(el)


def set_title(slide, text):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.name == "TextBox 4":
            p = sh.text_frame.paragraphs[0]
            p.runs[0].text = text
            for r in p.runs[1:]:
                r._r.getparent().remove(r._r)
            return True
    return False


def set_header(slide, text):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.name == "TextBox 1":
            sh.text_frame.paragraphs[0].runs[0].text = text
            return True
    return False


def main(src, dst):
    prs = Presentation(src)

    # ---------- 1) 架构页：改标题与页眉，移到 01 末尾
    a = idx_of(prs, "E0b")
    set_title(prs.slides[a], "模型架构：稠密、MoE、稀疏，差在哪")
    set_header(prs.slides[a], "01  认识 AI")
    # 移到「上下文腐烂」之后（即 01 末尾、四个常见问题之前）
    ctx = idx_contains(prs, "上下文腐烂")
    if a < ctx:
        ctx -= 1                      # 先移除会左移一位
    move(prs, a, ctx + 1)
    print("moved 架构 -> after", titles(prs)[ctx][:22])

    # ---------- 2) 按任务挑公司：改标题与页眉，移到 02 的「各家脾气」之后
    b = idx_of(prs, "E0  按任务挑公司")
    set_title(prs.slides[b], "按任务挑公司：六种活，各找各的人")
    set_header(prs.slides[b], "02  模型与工具")
    mood = idx_contains(prs, "各家模型的“脾气”")
    move(prs, b, mood + 1)
    print("moved 挑公司 -> after", titles(prs)[mood][:20])

    # ---------- 3) 两页进展插到对应公司后面
    e1 = idx_of(prs, "E1  OpenAI")
    m14 = idx_of(prs, "E14")
    move(prs, m14, e1 + 1)

    e2 = idx_of(prs, "E2  Anthropic")
    m15 = idx_of(prs, "E15")
    move(prs, m15, e2 + 1)

    # 改成公司内的编号（E1a / E2a 更清楚）
    for i, t in enumerate(titles(prs)):
        if t.startswith("E14"):
            set_title(prs.slides[i], "E1b  OpenAI 数学：从翻车到可验证的进展")
        if t.startswith("E15"):
            set_title(prs.slides[i], "E2b  Opus 出视频：一句话，靠写代码画帧")

    # ---------- 4) 页码重排（隐藏页不参与）
    page = 0
    for s in prs.slides:
        hidden = s._element.get("show") == "0"
        if not hidden:
            page += 1
        for sh in s.shapes:
            if not sh.has_text_frame or sh.left is None:
                continue
            if abs(sh.left / 914400 - 11.73) > 0.05 or sh.width / 914400 > 1.3:
                continue
            p = sh.text_frame.paragraphs[0]
            if p.runs and p.runs[0].text.strip().isdigit():
                p.runs[0].text = str(page)

    # ---------- 5) E 扉页文案与计数
    e_div = None
    for i, s in enumerate(prs.slides):
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() == "E" and \
                    sh.width is not None and sh.width / 914400 > 3:
                e_div = s
    if e_div is not None:
        for sh in e_div.shapes:
            if not sh.has_text_frame:
                continue
            t = sh.text_frame.text
            if "E1–E15" in t or "E0b" in t:
                sh.text_frame.paragraphs[0].runs[0].text = (
                    "E1–E13 每家一页：模型线、外壳、强项、边界，以及它在本次分享里出现过的地方；"
                    "OpenAI 与 Anthropic 各多一页今年的进展。")
            if t.strip().startswith("E  一页一家公司"):
                sh.text_frame.paragraphs[0].runs[0].text = "E  一页一家公司（十三页）"

    # ---------- 6) 01 / 02 扉页索引补两项
    def add_rail(slide, after_label, new_label):
        """在 after_label 这一行后面插入 new_label，并把其下的行下移。"""
        rows = [sh for sh in slide.shapes
                if sh.has_text_frame and sh.left is not None and
                sh.left / 914400 > 9.0 and sh.has_text_frame]
        rows = [sh for sh in rows if sh.text_frame.text.strip() and
                sh.text_frame.text.strip() != "这一部分"]
        rows.sort(key=lambda sh: sh.top)
        tgt = next((sh for sh in rows if sh.text_frame.text.strip() == after_label), None)
        if tgt is None:
            print("  ! rail row not found:", after_label)
            return False
        import copy
        new_el = copy.deepcopy(tgt._element)
        tgt._element.addnext(new_el)
        new_sh = [sh for sh in slide.shapes if sh._element is new_el][0]
        new_sh.text_frame.paragraphs[0].runs[0].text = new_label
        new_sh.top = tgt.top
        pitch = 0.56 * 914400
        for sh in rows:
            if sh.top > tgt.top:
                sh.top = int(sh.top + pitch)
        return True

    for s in prs.slides:
        hdr = None
        for sh in s.shapes:
            if sh.has_text_frame and sh.name == "TextBox 1":
                hdr = sh.text_frame.text
        for sh in s.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip() == "01  认识 AI":
                pass

    # 01 扉页：在「训练三步」后加「模型架构」
    for s in prs.slides:
        big = [sh for sh in s.shapes if sh.has_text_frame and
               sh.text_frame.text.strip() == "01" and sh.width and sh.width / 914400 > 3]
        if big:
            add_rail(s, "训练三步", "模型架构：稠密与 MoE")
            break

    # 02 扉页：在「各家脾气」后加「按任务挑公司」
    for s in prs.slides:
        big = [sh for sh in s.shapes if sh.has_text_frame and
               sh.text_frame.text.strip() == "02" and sh.width and sh.width / 914400 > 3]
        if big:
            add_rail(s, "榜单、全球格局、各家脾气", "按任务挑公司")
            break

    prs.save(dst)
    print("saved", dst, len(prs.slides))
    for i, t in enumerate(titles(prs), 1):
        if 13 <= i <= 34 or 88 <= i <= 108:
            print(f"  {i:3d} {t[:52]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
