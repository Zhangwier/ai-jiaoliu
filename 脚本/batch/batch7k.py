import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
s = prs.slides[51]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("附录  "):
        settext(sh, "附录  常用术语 A1–A4 · 公司档案 B1–B13（含 7 个专题页）· 订阅与省流 C1–C5（含 C1b、C4b）· 资讯渠道 D1–D4 · 分享与工具 E1–E4（含 10 个分页）· 原理与历史 F1–F9（含 F2b）· 来源 Z1–Z3", size=15)
        print("s52 ok", sh.text_frame.text[:60])
save(prs); print("saved")
