import io, os
from pptx import Presentation
from pptx.util import Emu
from PIL import Image
p = r"D:\MG\program\ai交流\成品\_work29\base.pptx"
prs = Presentation(p)
# slides 78-82 各页 picture 清单
for idx in [19,20,21,22,23,24,25,77,78,79,80,92]:
    s = prs.slides[idx]
    txt = " ".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)[:60].replace("\n"," ")
    pics=[]
    for sh in s.shapes:
        if sh.__class__.__name__ == "Picture":
            try:
                im = Image.open(io.BytesIO(sh.image.blob))
                pics.append(f"{sh.shape_id}:{sh.name} {im.size} {im.mode} l{Emu(sh.left).cm:.1f} t{Emu(sh.top).cm:.1f} w{Emu(sh.width).cm:.2f}")
            except: pass
    print(f"=== s{idx+1} {txt}")
    for x in pics: print("   ", x)
