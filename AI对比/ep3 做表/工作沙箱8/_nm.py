import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
pix=pg.get_pixmap(dpi=300, clip=pymupdf.Rect(155,76,215,460)); pix.save(r"work\NAME1.png"); print(pix.width,pix.height)
pix=pg.get_pixmap(dpi=300, clip=pymupdf.Rect(155,455,215,760)); pix.save(r"work\NAME2.png"); print(pix.width,pix.height)
pix=pg.get_pixmap(dpi=330, clip=pymupdf.Rect(486,76,540,300)); pix.save(r"work\NAME3.png"); print(pix.width,pix.height)
