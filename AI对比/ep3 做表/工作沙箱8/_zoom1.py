import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def c(n,r,dpi=700):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r)); pix.save(rf"work\{n}.png"); print(n,pix.width,pix.height)
c("NOTE24",(456,311,492,325))
c("N73",(600,175,686,189))
c("N84",(600,237,686,251))
c("SUMROW",(700,259,790,300))
