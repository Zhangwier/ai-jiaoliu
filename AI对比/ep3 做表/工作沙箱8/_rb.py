import pymupdf
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p); pg=d[2]
rows=[76.1,88.5,101.0,113.4,125.7,138.2,150.6,162.9,175.4,187.7,200.2,212.6,225.0,237.3,249.8,262.2,274.5,287.0,299.3]
for k in range(0,len(rows)-1,6):
    a=rows[k]; b=rows[min(k+6,len(rows)-1)]
    pix=pg.get_pixmap(dpi=330, clip=pymupdf.Rect(520,a,830,b)); out=rf"work\R{k//6+1:02d}.png"; pix.save(out); print(out,round(a,1),round(b,1))
