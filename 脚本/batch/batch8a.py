import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s=prs.slides[62]
for sh in s.shapes:
    if sh.has_text_frame and "见 C5" in sh.text_frame.text:
        settext(sh, "没有免费版，Claude Pro $20 起；对中国大陆用户与中资实体有地区限制（见 B2c）；最强档跑一道题约 $6。")
        print("fix1")
s=prs.slides[66]
for sh in s.shapes:
    if sh.has_text_frame and "Grok 并入 X 订阅" in sh.text_frame.text:
        settext(sh, "Grok 已并入 X 订阅（含 Cursor）；GLM Coding 只能在支持的编程工具里用；买前看续费价")
# verify C4 covers grok
s4=prs.slides[84]
for sh in s4.shapes:
    if sh.has_text_frame and "Grok" in sh.text_frame.text:
        print("C4:", sh.text_frame.text[:120].replace("\n"," / "))
save(prs)
