# -*- coding: utf-8 -*-
"""修正链接 run 的 rPr 子元素顺序：solidFill 必须在 hlinkClick 之前，否则 PowerPoint
会忽略我们设的灰色，改按主题超链接色（蓝）渲染。"""
import sys
from pptx import Presentation
from pptx.oxml.ns import qn, nsdecls
from pptx.oxml import parse_xml

ORDER = ["a:ln", "a:solidFill", "a:effectLst", "a:uLnTx", "a:uFillTx",
         "a:latin", "a:ea", "a:cs", "a:sym", "a:hlinkClick",
         "a:hlinkMouseOver", "a:rtl", "a:extLst"]


def fix_rpr(rPr):
    kids = list(rPr)
    changed = False
    # 必须有 hlinkClick 才算链接 run
    if not any(k.tag == qn("a:hlinkClick") for k in kids):
        return False
    for k in kids:
        rPr.remove(k)
    have = {k.tag: k for k in kids}
    for tag in ORDER:
        el = have.get(qn(tag))
        if el is not None:
            rPr.append(el)
            changed = True
    # 补 solidFill 和字体
    if qn("a:solidFill") not in have:
        rPr.append(parse_xml(
            '<a:solidFill %s><a:srgbClr val="6B7385"/></a:solidFill>' % nsdecls("a")))
        # 重新排一次
        kids2 = list(rPr)
        for k in kids2:
            rPr.remove(k)
        have2 = {k.tag: k for k in kids2}
        for tag in ORDER:
            el = have2.get(qn(tag))
            if el is not None:
                rPr.append(el)
    for tag in ("a:latin", "a:ea", "a:cs"):
        if qn(tag) not in have:
            rPr.append(parse_xml(
                '<%s %s typeface="微软雅黑"/>' % (tag, nsdecls("a"))))
    return changed


def main(src, dst):
    prs = Presentation(src)
    n = 0
    for s in prs.slides:
        for r in s._element.iter(qn("a:r")):
            rPr = r.find(qn("a:rPr"))
            if rPr is not None and fix_rpr(rPr):
                n += 1
    prs.save(dst)
    print("fixed", n, "runs ->", dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
