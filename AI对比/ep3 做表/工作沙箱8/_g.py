import pymupdf
from PIL import Image
import numpy as np
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def grab(r,dpi=900):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r))
    return np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,pix.n)
# N12 spec cell region and N15 spec cell region
a=grab((275,101,355,113))
b=grab((275,138,355,150))
print(a.shape,b.shape)
