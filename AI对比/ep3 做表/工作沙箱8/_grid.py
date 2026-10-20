import pymupdf, collections
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
print("page rect", pg.rect, "cropbox", pg.cropbox, "mediabox", pg.mediabox)
dr = pg.get_drawings()
print("n drawings", len(dr))
xs=collections.Counter(); ys=collections.Counter()
for it in dr:
    r = it["rect"]
    for k in ("l","re"):
        pass
    for item in it["items"]:
        if item[0]=="l":
            p1,p2 = item[1],item[2]
            if abs(p1.x-p2.x)<0.6: xs[round(p1.x,1)]+=1
            if abs(p1.y-p2.y)<0.6: ys[round(p1.y,1)]+=1
        elif item[0]=="re":
            r2=item[1]
            xs[round(r2.x0,1)]+=1; xs[round(r2.x1,1)]+=1
            ys[round(r2.y0,1)]+=1; ys[round(r2.y1,1)]+=1
print("X lines:", sorted(xs)[:60])
print("Y lines:", sorted(ys)[:60])
