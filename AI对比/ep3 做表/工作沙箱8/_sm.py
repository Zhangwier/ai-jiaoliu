import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
pix=pg.get_pixmap(dpi=260, clip=pymupdf.Rect(440,395,910,605)); pix.save(r"work\SUMMARY.png"); print(pix.width,pix.height)
pix=pg.get_pixmap(dpi=300, clip=pymupdf.Rect(300,55,380,70)); pix.save(r"work\TITLE.png"); print(pix.width,pix.height)
