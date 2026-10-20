import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
def crop(name,r,dpi=300):
    pix=pg.get_pixmap(dpi=dpi, clip=pymupdf.Rect(*r)); pix.save(rf"work\{name}.png"); print(name,pix.width,pix.height)
crop("HDR",(80,98,500,128))
crop("ROWFULL",(80,100.8,500,175),dpi=300)
