import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[38]
for sh in s.shapes:
    if sh.has_text_frame and sh.text_frame.text.startswith("先给一套判断标准"):
        settext(sh, "第一类省搜索：帮你找到不知道存在的路子。\n第二类省人力：量大、门槛低、目标明确，不需要复核。\n第三类省判断：综合任务，要拆步、要留痕、要人签字。", size=16)
        print("ok", sh.top/360000, sh.height/360000)
save(prs)
