# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs(); s=prs.slides[58]
for sh in s.shapes:
    if sh.has_text_frame and sh.left/EMU<22 and 12.0<sh.top/EMU<14.6 and sh.height/EMU>1.0:
        settext(sh,"模型线 · 可用入口 · 强项 · 边界；OpenAI、Anthropic、Meta Muse、DeepSeek、Qwen 另有专题页；降智与封号紧跟其后", size=14.5)
        print("ok")
save(prs)
