# -*- coding: utf-8 -*-
from pptx import Presentation
p = r"D:\MG\program\ai交流\成品\_work29\base.pptx"
prs = Presentation(p)
BR = ["OpenAI","Anthropic","Google","Meta","xAI","Mistral","DeepSeek","阿里","Qwen","千问","智谱","GLM","Kimi","月之暗面","MiniMax","腾讯","混元","Hy","百度","文心","字节","Seed","豆包","ChatGPT","Claude","Gemini","Grok","Muse","Hark","Trae","Qoder","Cline","Codex","WorkBuddy","Cursor","Kiro"]
for i,s in enumerate(prs.slides):
    # title = shape with sz>=26 near top
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        r = sh.text_frame.paragraphs[0].runs
        sz = r[0].font.size.pt if r and r[0].font.size else 0
        if sz and sz>=24 and sh.top/360000 < 4:
            t = sh.text_frame.text.replace("\n"," / ")
            hits = [b for b in BR if b in t]
            npic = sum(1 for x in s.shapes if x.shape_type==13)
            print(f"s{i+1:>3} sz{sz:.0f} pics={npic} {hits} {t[:60]!r}")
            break
