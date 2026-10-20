import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
S=400/72
# left table px 1163..2860 (incl 备注), rows 560..4210 ; convert to pt
r = pymupdf.Rect(1163/S, 560/S, 2870/S, 4215/S)
pix = pg.get_pixmap(dpi=300, clip=r); pix.save(r"work\table_left.png"); print(r, pix.width, pix.height)
r2 = pymupdf.Rect(2981/S, 560/S, 4560/S, 1670/S)
pix = pg.get_pixmap(dpi=300, clip=r2); pix.save(r"work\table_right.png"); print(r2, pix.width, pix.height)
