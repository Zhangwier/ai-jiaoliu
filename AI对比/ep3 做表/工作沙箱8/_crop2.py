import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
def crops(name, x0,x1, ys, dpi=360):
    for i,(y0,y1) in enumerate(ys,1):
        pix = pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(x0,y0,x1,y1))
        out = rf"work\{name}{i}.png"; pix.save(out); print(out, pix.width, pix.height)
crops("L", 143,442, [(93,205),(200,305),(300,405),(400,505),(500,605),(600,672)])
crops("R", 545,845, [(93,205),(200,305),(300,405)])
crops("S", 730,880, [(395,545)])
