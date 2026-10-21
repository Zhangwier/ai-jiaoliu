import sys
sys.path.insert(0, r"D:\MG\program\ai交流\脚本")
from pptx import Presentation
p=Presentation(r"D:\MG\program\ai交流\成品\AI交流分享-成品32.pptx")
EMU=360000
bad=[]
for i,s in enumerate(p.slides):
    for sh in s.shapes:
        if not sh.has_text_frame: continue
        t=sh.text_frame.text
        for b in ("会后","一页一家公司","开源平台对照","案例、风险与破限"):
            if b in t: bad.append((i+1,sh.name,b))
print("stale:",bad)
print("slides:",len(p.slides))
# E4 + s5 + s108 检查
for idx in (4,108):
    print("---",idx+1)
    for sh in p.slides[idx].shapes:
        if sh.has_text_frame and sh.text_frame.text.strip() and sh.top/EMU<4.2:
            print("   ",sh.text_frame.text.strip()[:60])
