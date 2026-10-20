# -*- coding: utf-8 -*-
import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
EMU=360000
prs = openprs(); s=prs.slides[108]
for sh in s.shapes:
    if not sh.has_text_frame: continue
    n=sh.name
    if n=="TextBox 1": settext(sh,"附录  速查", size=12); print("hdr fixed")
    elif n=="TextBox 4": settext(sh,"E4  开源落地：从权重到能用的四层", size=26)
    elif n=="TextBox 5": settext(sh,"模型权重只是第一层；从权重到“同事能上手”，中间还有三层要拼——每层都有成熟的开源项目，难的不是装，是运维和权限", size=15)
    elif n=="TextBox 7": settext(sh,"来源：各项目 GitHub 仓库 README（vLLM、SGLang、Ollama、LiteLLM、One-API、Open WebUI、Dify、RAGFlow、Langfuse）；2026-10-20 查询；本稿为选型清单，未做部署实测", size=9)
save(prs); print("saved")
