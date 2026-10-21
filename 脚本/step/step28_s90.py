# -*- coding: utf-8 -*-
from pptx import Presentation
from pptx.util import Pt
P = r"D:\MG\program\ai交流\成品\_work29\base.pptx"
prs = Presentation(P)
s = prs.slides[89]
S = {sh.shape_id: sh for sh in s.shapes}
def set2(sid, text):
    sh = S[sid]
    tf = sh.text_frame
    tf.text = text
    for p in tf.paragraphs:
        for r in p.runs: r.font.size = Pt(13)
    # 保持首行 13pt 与行距
for sid, txt in ((17,"鹈鹕动画\n（测试一）"),
                 (23,"糖果题\n（测试二）"),
                 (29,"多合一题\n（工具一 · E1b）"),
                 (35,"模型指纹\n（工具二 · E1c）")):
    set2(sid, txt)
S[11].text_frame.text = "四道题"
S[6].text_frame.text = "前两道是判分用的测试题，后两道是可以直接拿来用的自查工具；四道各测一件事"
for p in S[6].text_frame.paragraphs:
    for r in p.runs: r.font.size = Pt(14)
prs.save(P)
print("ok")
