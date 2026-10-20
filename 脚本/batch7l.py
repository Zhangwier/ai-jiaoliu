import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[38]
m = {
 "什么活适合交给 AI": "三类活：先分对，再决定检查强度",
 "找图纸：先建索引再检索": "第一类实例：一百多个文件里找一张图",
 "方法总览：对应哪个问题": "五种方法：各自对付一个问题",
}
for sh in s.shapes:
    if sh.has_text_frame:
        t = sh.text_frame.text.strip()
        if t in m:
            old_sz = sh.text_frame.paragraphs[0].runs[0].font.size.pt if sh.text_frame.paragraphs[0].runs and sh.text_frame.paragraphs[0].runs[0].font.size else 17
            settext(sh, m[t], size=min(old_sz, 15))
            print("->", m[t])
save(prs); print("saved")
