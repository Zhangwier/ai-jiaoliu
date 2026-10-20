import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[70]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("① 开源不只放权重"):
        settext(sh, "① 开源的是方法，不只是权重：报告、算子、评测代码一起放（见 B7b）。② 价格最低一档：输出 8 元/百万 token，闲时半价，约 217 token/s。③ Flash 1017 已贴到 5.6 Sol 那一档：官方自报 DeepSWE 77.0。", size=11.5)
save(prs); print("ok")
