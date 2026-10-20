import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs(); s = prs.slides[39]
m = {
 "TextBox 19": "怎么做：先问“有哪些路子”，再让它列成对比表",
 "TextBox 30": "怎么做：一次跑一百个，先抽 3 个验格式",
 "TextBox 41": "怎么做：拆成小步，每步留一个能查的结果",
}
for sh in s.shapes:
    if sh.name in m:
        settext(sh, m[sh.name], size=11.5)
save(prs); print("ok")
