import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
pix=pg.get_pixmap(dpi=340, clip=pymupdf.Rect(560,770,900,815)); pix.save(r"work\TITLEBLOCK.png"); print(pix.width,pix.height)
