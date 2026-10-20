import pymupdf
from PIL import Image
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
pix=pg.get_pixmap(dpi=340, clip=pymupdf.Rect(161,76,210,760)); pix.save(r"work\names_raw.png")
im=Image.open(r"work\names_raw.png").rotate(-90, expand=True)
im.save(r"work\names_rot.png"); print(im.size)
