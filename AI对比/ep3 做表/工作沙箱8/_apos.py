import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def c(n,r,dpi=600):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r)); pix.save(rf"work\{n}.png"); print(n,pix.width,pix.height)
c("HDR2",(370,76,820,89))
c("APOS1",(205,140,280,180),dpi=700)
c("APOS2",(205,325,280,400),dpi=700)
