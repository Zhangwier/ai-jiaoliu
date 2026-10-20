import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
crops = {
 "L1": (100, 52, 292, 160),
 "L2": (100, 152, 292, 262),
 "L3": (100, 254, 292, 366),
 "L4": (100, 358, 292, 470),
 "R1": (296, 52, 456, 168),
 "R2": (296, 160, 456, 262),
 "S1": (450, 258, 600, 412),
}
for k,(x0,y0,x1,y1) in crops.items():
    pix = pg.get_pixmap(dpi=400, clip=pymupdf.Rect(x0,y0,x1,y1))
    out = rf"work\p3_{k}.png"
    pix.save(out); print(out, pix.width, pix.height)
