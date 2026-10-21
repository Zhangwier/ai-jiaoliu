# -*- coding: utf-8 -*-
"""把附录 E（E 分隔页 + E0–E14）从 成品16+附录E 迁到 成品19，并重排页码。
   目标顺序：…… D4 → E 分隔页 → E0–E14 → Z 分隔页(Z1) → Z → Z2 → Z3 → 隐藏页"""
import sys
import copy
from pptx import Presentation
from pptx.oxml.ns import qn


def title_of(slide):
    for sh in slide.shapes:
        if sh.has_text_frame and sh.name == "TextBox 4":
            return sh.text_frame.text.split("\n")[0]
    return ""


def find(slides, pred):
    for i, s in enumerate(slides):
        if pred(title_of(s), s):
            return i
    return None


def clone_slides(src_prs, dst_prs, idxs):
    """按顺序把 src 的幻灯片复制到 dst 末尾，返回新 sldId 元素列表。"""
    lst = dst_prs.slides._sldIdLst
    made = []
    for i in idxs:
        s = src_prs.slides[i]
        blank = s.slide_layout
        new = dst_prs.slides.add_slide(blank)
        # 清空新建页的占位符
        for sh in list(new.shapes):
            sh._element.getparent().remove(sh._element)
        # 复制形状
        for sh in s.shapes:
            new.shapes._spTree.append(copy.deepcopy(sh._element))
        # 复制背景
        bg = s._element.find(qn("p:cSld")).find(qn("p:bg"))
        if bg is not None:
            csld = new._element.find(qn("p:cSld"))
            old = csld.find(qn("p:bg"))
            if old is not None:
                csld.remove(old)
            csld.insert(0, copy.deepcopy(bg))
        # 备注
        if s.has_notes_slide:
            new.notes_slide.notes_text_frame.text = \
                s.notes_slide.notes_text_frame.text
        made.append(list(lst)[-1])
    return made


def main(src19, src16, dst):
    a = Presentation(src19)          # 主体（含路线页与案例三）
    b = Presentation(src16)          # 附录 E 来源

    e_div = find(b.slides, lambda t, s: any(
        sh.has_text_frame and sh.text_frame.text.strip() == "E" for sh in s.shapes))
    e0 = find(b.slides, lambda t, s: t.startswith("E0"))
    e14 = find(b.slides, lambda t, s: t.startswith("E14"))
    idxs = [e_div] + list(range(e0, e14 + 1))
    print("E pages:", [i + 1 for i in idxs])

    # Z 分隔页在 a 里的位置（big letter Z）
    z_idx = find(a.slides, lambda t, s: any(
        sh.has_text_frame and sh.text_frame.text.strip() == "Z" for sh in s.shapes))
    print("Z divider at", z_idx + 1)

    new_ids = clone_slides(b, a, idxs)

    lst = a.slides._sldIdLst
    items = list(lst)
    for el in items:
        lst.remove(el)
    keep = [el for el in items if el not in new_ids]
    order = keep[:z_idx] + new_ids + keep[z_idx:]
    for el in order:
        lst.append(el)

    # 重排页码：从 Z 分隔页开始顺延
    z_new = z_idx + len(new_ids)
    page = None
    for i, s in enumerate(a.slides):
        if i == z_new - len(new_ids):
            # E 分隔页：无页码
            pass
    # 简单做法：E0 起始页号 = Z 分隔页原本应显示的号
    base = z_idx  # 0-based；Z 分隔页无号，E0 号 = z_idx+1
    for i, s in enumerate(a.slides):
        for sh in s.shapes:
            if not sh.has_text_frame or sh.left is None:
                continue
            if abs(sh.left / 914400 - 11.73) > 0.05 or sh.width / 914400 > 1.3:
                continue
            p = sh.text_frame.paragraphs[0]
            if not p.runs or not p.runs[0].text.strip().isdigit():
                continue
            p.runs[0].text = str(i + 1)

    # E 分隔页副标题里的页数
    for s in a.slides:
        for sh in s.shapes:
            if not sh.has_text_frame:
                continue
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if "E1–E14" in r.text and "共" not in r.text:
                        pass

    a.save(dst)
    print("saved", dst, len(a.slides))
    for i, s in enumerate(a.slides, 1):
        if i >= 85:
            n = None
            for sh in s.shapes:
                if sh.has_text_frame and sh.left is not None and \
                        abs(sh.left / 914400 - 11.73) < 0.05 and sh.width / 914400 < 1.3:
                    n = sh.text_frame.text
            print(f"  {i:3d} [{n}] {title_of(s)[:40]}")


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], sys.argv[3])
