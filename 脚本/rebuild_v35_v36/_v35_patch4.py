# -*- coding: utf-8 -*-
p = "ds.py"
s = open(p, encoding="utf-8").read()
s = s.replace("""    for e in tblPr.findall(qn('a:tableStyleId')):
        tblPr.remove(e)""",
"""    _NoStyle = '{2D5ABB26-0587-4C30-8999-92F81FD0307C}'
    sid = tblPr.find(qn('a:tableStyleId'))
    if sid is None:
        sid = etree.SubElement(tblPr, qn('a:tableStyleId'))
    sid.text = _NoStyle""")
open(p, "w", encoding="utf-8").write(s)
print("ok")
