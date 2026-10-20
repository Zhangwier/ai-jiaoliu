import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
for dpi in (150,):
    pix = pg.get_pixmap(dpi=dpi)
    out = rf"work\page3_{dpi}.png"
    pix.save(out)
    print(out, pix.width, pix.height)
