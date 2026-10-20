import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
pix = pg.get_pixmap(dpi=400)
pix.save(r"work\p3_400.png")
print(pix.width, pix.height)
