# -*- coding: utf-8 -*-
"""Insert logo= entries into build_appendix_e.py company data blocks."""
import io
import re

P = "build_appendix_e.py"
s = io.open(P, encoding="utf-8").read()

pairs = [
    ('art=imgs["gpt"]', 'openai'),
    ('art=imgs["claude"]', 'anthropic'),
    ('art=imgs["gemini"]', 'gemini'),
    ('art=imgs["grok"]', 'grok'),
    ('art=imgs["muse"]', 'meta'),
    ('art=imgs["mistral"]', 'mistral'),
    ('art=imgs["deepseek"]', 'deepseek'),
    ('art=imgs["qwen"]', 'qwen'),
    ('art=imgs["glm"]', 'glm'),
    ('art=imgs["kimi"]', 'kimi'),
    ('art=imgs["minimax"]', 'minimax'),
    ('art=imgs["hy"]', 'hy'),
    ('art=imgs["baidu"]', 'baidu'),
    ('art=None, word="dots"', 'dots'),
]
for anchor, key in pairs:
    assert anchor in s, anchor
    s = s.replace(anchor, 'logo=logos["%s"], %s' % (key, anchor), 1)

io.open(P, "w", encoding="utf-8").write(s)
print("patched", len(pairs))
