import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from _lib import *
from pptx.util import Cm
from pptx.enum.shapes import MSO_SHAPE
EMU=360000
prs = openprs(); s = prs.slides[68]
# 去掉 last 行的白底矩形（第5行 i=4 偶数 -> 有白底，但超出）
kill=[]
for sh in s.shapes:
    if sh.shape_type == MSO_SHAPE.RECTANGLE and abs(sh.top/EMU-14.75)<0.06 and sh.width/EMU>20:
        kill.append(sh)
for k in kill: k._element.getparent().remove(k._element)
# 行高改为 2.06，从小 2.26 起算
y = 4.83
for i in range(5):
    pass
rows=[sh for sh in s.shapes if sh.has_text_frame and abs(sh.left/EMU-9.19)<0.05]
rows.sort(key=lambda x:x.top)
labels=[sh for sh in s.shapes if sh.has_text_frame and abs(sh.left/EMU-12.14)<0.05]
labels.sort(key=lambda x:x.top)
newY=[4.83,6.98,9.13,11.28,13.43]
for i,(a,b) in enumerate(zip(rows,labels)):
    a.top=Cm(newY[i]); b.top=Cm(newY[i]); b.height=Cm(1.95)
    a.height=Cm(1.30)
# 分隔线重排
conns=[sh for sh in s.shapes if sh.shape_type==9 and sh.top/EMU>4.8 and sh.top/EMU<17]
conns.sort(key=lambda x:x.top)
for i,c in enumerate(conns):
    if i < 4: c.top = Cm(newY[i+1]-0.05)
save(prs); print("ok", len(rows), len(labels), len(conns))
