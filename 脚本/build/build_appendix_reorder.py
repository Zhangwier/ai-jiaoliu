# -*- coding: utf-8 -*-
"""
附录重排：B E A C D（Z 保持最后），并把编号字母一并改成新顺序。

原：A 订阅与省流 58-65 / B 常用术语 66-70 / C 资讯渠道 71-75
    D 案例、风险与破限 76-92 / E 一页一家公司 93-110 / Z 111-114
新顺序：旧 B → 新 A；旧 E → 新 B；旧 A → 新 C；旧 C → 新 D；旧 D → 新 E；Z 不变
"""
import re
from pptx import Presentation

F = r"D:\MG\program\ai交流\AI交流分享-成品26.pptx"

L2L = {"A": "C", "B": "A", "C": "D", "D": "E", "E": "B"}

# 旧字母 -> 原页范围（1-based，含分节页）
RANGES = {
    "B": [66, 67, 68, 69, 70],
    "E": list(range(93, 111)),
    "A": list(range(58, 66)),
    "C": list(range(71, 76)),
    "D": list(range(76, 93)),
}
Z_RANGE = list(range(111, 115))
TAIL = list(range(115, 125))          # 十宫格放大播放页
HEAD = list(range(1, 58))             # 正文 1-57

DIVIDERS = {58: "A", 66: "B", 71: "C", 76: "D", 93: "E", 111: "Z"}

# 新字母 -> 目录用说明（页数按实际内容页，不含分节页）
NAV_DESC = {
    "A": "常用术语（四页）",
    "B": "一页一家公司（十七页）",
    "C": "订阅与省流（七页）",
    "D": "资讯渠道（四页）",
    "E": "案例、风险与破限（十六页）",
}

NEW_LETTER_OF_OLD = {"A": "C", "B": "A", "C": "D", "D": "E", "E": "B"}

CODE = re.compile(r"(?<![A-Za-z0-9])([A-E])(\d{1,2})(b?)(?![0-9A-Za-z])")


def remap(t):
    return CODE.sub(lambda m: L2L[m.group(1)] + m.group(2) + m.group(3), t)


def skip_run(t):
    """跳过账号示意标签 A01..A11（附录编号从不用前导零）。"""
    return re.fullmatch(r"[A-E]0\d", t) is not None


prs = Presentation(F)
slides = list(prs.slides)

# ========== 1) 页内编号与交叉引用重贴字母 ==========
n = 0
for i in range(58, 115):
    s = slides[i - 1]
    for sh in s.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in p.runs:
                if skip_run(r.text.strip()):
                    continue
                nt = remap(r.text)
                if nt != r.text:
                    r.text = nt
                    n += 1
    if s.has_notes_slide:
        tf = s.notes_slide.notes_text_frame
        for p in tf.paragraphs:
            for r in p.runs:
                if skip_run(r.text.strip()):
                    continue
                nt = remap(r.text)
                if nt != r.text:
                    r.text = nt
                    n += 1
print("codes relabelled:", n)


def set_single(shape, text):
    p = shape.text_frame.paragraphs[0]
    runs = p.runs
    if not runs:
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


# ========== 2) 六个分节页：大字 + 右侧目录 ==========
for page_no, old_letter in DIVIDERS.items():
    s = slides[page_no - 1]
    new_letter = NEW_LETTER_OF_OLD.get(old_letter, "Z")

    # 大字
    for sh in s.shapes:
        if not sh.has_text_frame or sh.top is None:
            continue
        if 1.4 <= sh.top / 914400 <= 3.3 and sh.text_frame.text.strip() == old_letter:
            set_single(sh, new_letter)

    # 右侧目录：按 y 排序后重写成 新顺序 + Z
    items = []
    for sh in s.shapes:
        if not sh.has_text_frame or sh.left is None or sh.top is None:
            continue
        if abs(sh.left / 914400 - 9.1) > 0.06:
            continue
        if sh.top / 914400 < 1.9:
            continue
        if sh.text_frame.text.strip() == "这一部分":
            continue
        items.append(sh)
    items.sort(key=lambda sh: sh.top)

    order = ["A", "B", "C", "D", "E", "Z"]   # 新字母顺序
    if len(items) != 6:
        print("!! divider", page_no, "nav count", len(items), "- abort")
    else:
        for sh, letter in zip(items, order):
            if letter == "Z":
                set_single(sh, "Z  参考资料与来源")
            else:
                set_single(sh, "%s  %s" % (letter, NAV_DESC[letter]))
        # 重排 y
        y = 2.10
        for sh in items:
            sh.top = int(y * 914400)
            y += 0.50
    print("divider", page_no, "->", new_letter)

# ========== 3) 第 3 页总目录的附录清单 ==========
s3 = slides[2]
for sh in s3.shapes:
    if not sh.has_text_frame:
        continue
    t = sh.text_frame.text
    if "订阅与省流版" in t and "常用术语" in t:
        tf = sh.text_frame
        for extra in tf.paragraphs[1:]:
            extra._p.getparent().remove(extra._p)
        lines = ["A 常用术语", "B 一页一家公司", "C 订阅与省流版",
                 "D 资讯渠道", "E 案例、风险与破限", "Z 参考资料来源"]
        first = tf.paragraphs[0]
        runs = first.runs
        base = runs[0]
        base.text = lines[0]
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
        for line in lines[1:]:
            p = tf.add_paragraph()
            r = p.add_run()
            r.text = line
            r.font.size = base.font.size
            r.font.name = base.font.name
            r.font.color.rgb = base.font.color.rgb
            p.line_spacing = first.line_spacing
            p.space_before = first.space_before
            p.space_after = first.space_after
            p.alignment = first.alignment
        print("TOC appendix list rewritten")

# ========== 4) 重排 slide order ==========
order_pages = HEAD + [p for k in ["B", "E", "A", "C", "D"] for p in RANGES[k]] + Z_RANGE + TAIL
assert sorted(order_pages) == list(range(1, 125)), "order mismatch"

sldIdLst = prs.slides._sldIdLst
elems = list(sldIdLst)
for e in elems:
    sldIdLst.remove(e)
for pg in order_pages:
    sldIdLst.append(elems[pg - 1])

# ========== 5) 页码重编 ==========
fixed = 0
for i, sl in enumerate(prs.slides, 1):
    for sh in sl.shapes:
        if not sh.has_text_frame or sh.left is None or sh.top is None:
            continue
        if not sh.text_frame.text.strip().isdigit():
            continue
        if sh.left / 914400 > 11.4 and sh.top / 914400 < 0.6:
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.text.strip().isdigit():
                        r.text = str(i)
            fixed += 1
print("page numbers:", fixed)

prs.save(F)
print("saved", len(prs.slides.__iter__.__self__._sldIdLst), "slides")
