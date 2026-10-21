# -*- coding: utf-8 -*-
import sys, os
sys.path.insert(0, r'_v35')
import ds
from ds import *
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

prs = new_deck()

# A. serif title
s = slide(prs)
header(s, '01  认识 AI', 7)
title(s, 'AI 的五个短板，每一个都能从原理上讲清', 0.80, 31, INK, SERIF)
lede(s, '这五个不是偶发故障，而是这套生成机制的必然产物。后面每个案例都能对上其中一条。')
footer(s, '来源：本部分各页来源')

# B. yahei bold title
s = slide(prs)
header(s, '02  模型与工具', 19)
title(s, '同一句提示词，十个配置给出十种结果', 0.80, 31, INK, SANS, bold=True)
lede(s, '一次生成、没有历史记忆；差别不在提示词，而在模型档位、专精方向与外壳能力。')
footer(s, '来源：10 份原始 HTML 录屏（统一视口，去黑边）')

# C. simsun title
s = slide(prs)
header(s, '03  两个真实案例', 30)
title(s, '十个配置审同一张图：三个找到关键问题，另有三十条假错误', 0.80, 31, INK, 'SimSun')
lede(s, '说得越具体，越像真的——这是假错误最难被发现的地方。')
footer(s, '来源：本地图纸审查案例记录')

# D. table + rules style check
s = slide(prs)
header(s, '03  两个真实案例', 31)
title(s, '读图方式决定命中率：只看 OCR 的两组零命中', 0.80, 30, INK, SERIF)
rows = [
 ['配置', '耗时', 'K1 分段', 'K2 符号', '其他', '假错误', '条数'],
 ['Opus 5.5 · Hark', '3分03秒', '命中', '命中', '2', '2', '11'],
 ['GPT 6.1 Sol high · Codex', '14分34秒', '命中', '—', '4', '1', '8'],
 ['Qwen 3.8 Flash · Qoder', '43分09秒', '—', '命中', '2', '5', '24'],
 ['Seed 2.1 Turbo · Trae', '12分47秒', '—', '—', '3', '10', '10'],
]
table(s, ML, 2.10, 7.4, rows, [2.6,1.2,1.0,1.0,0.7,0.9,0.7], rowh=0.42, size=10.5,
      head_size=9.5, aligns=[PP_ALIGN.LEFT,PP_ALIGN.RIGHT,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER,PP_ALIGN.CENTER])
stat(s, 8.60, 2.10, '3 / 10', '', '命中关键问题的配置数', RED, 40, 3.4)
stat(s, 8.60, 3.70, '30', '条', '另一组同时给出的假错误', INK, 40, 3.4)
footer(s, '来源：本地图纸审查记录（21 条待核未计）')

# E. dark slide
s = slide(prs, DARKBG)
header(s, '开场', 3, dark=True)
text(s, ML, 1.40, 9.9, 2.4, '交付物越像样，\n越要知道它会在哪里错。', 46, WHITE, SERIF, False, PP_ALIGN.LEFT, 1.30)
text(s, ML, 4.30, 5.4, 0.9, '两个真实任务、十组配置、一百多条可核对的结论——\n把它们放在一起看，规律就出来了。', 13, C(0xB6,0xBE,0xC9), SANS, False, PP_ALIGN.LEFT, 1.55)
footer(s, '来源：本地分享材料', dark=True)

prs.save(r'_v35\probe2.pptx')
print('warn', len(ds.WARN))
for w in ds.WARN: print(' ', w)
