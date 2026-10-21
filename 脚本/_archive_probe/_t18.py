import sys, re
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
pat = re.compile(r"[（(][一二三四五六七八九十]+页[）)]|共\s*\d+\s*页")
for i,s in enumerate(prs.slides,1):
    for sh in s.shapes:
        if sh.has_text_frame and pat.search(sh.text_frame.text):
            print(i, "|", sh.text_frame.text.strip()[:100].replace("\n"," / "))
