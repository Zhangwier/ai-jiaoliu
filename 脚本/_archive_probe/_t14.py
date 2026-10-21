import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
keys=["vLLM","SGLang","Dify","RAGFlow","LiteLLM","Ollama","One API","Open WebUI","Langfuse","llama.cpp","私有化","内网"]
for i,s in enumerate(prs.slides,1):
    for sh in s.shapes:
        if sh.has_text_frame:
            t=sh.text_frame.text
            for k in keys:
                if k in t: print(i, k, "|", t[:80].replace("\n"," / "))
