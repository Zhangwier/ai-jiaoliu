# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *

prs = openprs()
slides = list(prs.slides)
N = len(slides)
print("total", N)

def idx(key):
    for i, s in enumerate(slides):
        if key in title_of(s):
            return i
    raise KeyError(key)

i7   = idx("AI 是从例子里学规律")
i8   = idx("神经网络")
i9   = idx("Transformer")
i10  = idx("训练分三步")
i11  = idx("回答是一个 token")
i12  = idx("抽卡")
i13  = idx("幻觉：它会一本正经地编")
i14  = idx("迎合：你说")
i15  = idx("迎合的第二种形态")
i16  = idx("上下文腐烂")
i17  = idx("四个问题，都能在原理里找到原因")
i60  = idx("B1b OpenAI 数学")
i62  = idx("B2b Opus 生成视频")
i81  = idx("Anthropic 对中国用户")
i82  = idx("官方承认过 bug")
i115 = idx("F9 多模态")
print("idx", i7,i8,i9,i10,i11,i12,i13,i14,i15,i16,i17,i60,i62,i81,i82,i115)
