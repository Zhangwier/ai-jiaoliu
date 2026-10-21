import os
from PIL import Image, ImageDraw
m=r"_v34\media"
spec=[("s001_1.png",900),("s060_2.jpg",600),("s062_1.png",800),("s037_1.png",900),("s034_1.png",900),("s096_1.jpg",800)]
for f,wanted in spec:
    im=Image.open(os.path.join(m,f)).convert("RGB")
    w,h=im.size
    s=min(1.0, wanted/w)
    im=im.resize((int(w*s),int(h*s)),Image.LANCZOS)
    im.save(rf"_v35\look_{f.split('.')[0]}.png")
    print(f, w,h,"->",im.size)
