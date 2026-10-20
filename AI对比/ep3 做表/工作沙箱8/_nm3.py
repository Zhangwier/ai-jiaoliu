import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def c(n,r,dpi=340):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r)); pix.save(rf"work\{n}.png"); print(n,pix.width,pix.height)
c("NAMECOL",(486,76,540,300))
c("FOOT",(480,240,830,302))
