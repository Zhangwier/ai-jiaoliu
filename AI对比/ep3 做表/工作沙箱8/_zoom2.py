import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def c(n,r,dpi=700):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r)); pix.save(rf"work\{n}.png"); print(n,pix.width,pix.height)
c("N82",(600,212,686,226))
c("N71",(600,88,690,101))
c("N73A",(600,187,690,200))
c("N31",(275,435,360,448))
