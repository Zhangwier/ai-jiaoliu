import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
prs = openprs()
print("total", len(prs.slides))
for i, sl in enumerate(prs.slides, start=1):
    pn = ""
    ttl = ""
    for sh in sl.shapes:
        if sh.has_text_frame:
            t = sh.text_frame.text.strip()
            if sh.name == "TextBox 2" and t.isdigit(): pn = t
            if sh.name == "TextBox 4": ttl = t.replace("\n"," / ")[:52]
    print(f"{i:3d} pg[{pn:>3s}] {ttl}")
