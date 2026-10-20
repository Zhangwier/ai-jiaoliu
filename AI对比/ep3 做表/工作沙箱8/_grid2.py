import pymupdf, collections
p = r"主线G6-2 新沙高架桥-仅含数量表.pdf"
d = pymupdf.open(p)
pg = d[2]
dr = pg.get_drawings()
vs=collections.Counter(); hs=collections.Counter()
for it in dr:
    for item in it["items"]:
        if item[0]=="l":
            a,b = item[1],item[2]
            if abs(a.x-b.x)<0.8 and abs(a.y-b.y)>60: vs[round((a.x+b.x)/2,1)]+=1
            if abs(a.y-b.y)<0.8 and abs(a.x-b.x)>60: hs[round((a.y+b.y)/2,1)]+=1
        elif item[0]=="re":
            r=item[1]
            if r.height>60 and r.width<2: vs[round((r.x0+r.x1)/2,1)]+=1
            if r.width>60 and r.height<2: hs[round((r.y0+r.y1)/2,1)]+=1
print("V:", sorted(vs.items()))
print("H:", sorted(hs.items()))
